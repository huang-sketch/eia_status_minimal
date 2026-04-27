import json
import os
import re
import urllib.error
import urllib.request
from typing import Any, Dict, List, Optional


FIELDS = [
    "source_type",
    "monitor_type",
    "point",
    "factor",
    "value",
    "unit",
    "standard_class",
    "evidence",
    "confidence",
]

DEFAULT_ENDPOINT = "https://api.openai.com/v1/chat/completions"
DEFAULT_MODEL = "gpt-4o-mini"

NUMBER_RE = re.compile(r"[<≤>=]?\d+(?:\.\d+)?(?:\s*[-~～]\s*\d+(?:\.\d+)?)?")
UNIT_RE = re.compile(r"mg/L|mg/m3|µg/m3|ug/m3|dB\(A\)|dB|无量纲", re.IGNORECASE)
FACTOR_RE = re.compile(
    r"COD|BOD5|氨氮|总磷|总氮|pH|PM10|PM2\.5|SO2|NO2|噪声|昼间|夜间",
    re.IGNORECASE,
)


def extract_records_from_chunk(chunk: Dict[str, Any]) -> List[Dict[str, Any]]:
    """
    Extract records from one chunk only.

    If EIA_LLM_API_KEY is configured, call an OpenAI-compatible chat
    completions endpoint. Otherwise use a small mock extractor so the project
    can run locally without network or credentials.
    """
    if os.getenv("EIA_LLM_API_KEY"):
        raw_records = call_llm(chunk["text"])
    else:
        raw_records = mock_extract(chunk["text"])

    return normalize_and_validate(raw_records, chunk)


def call_llm(chunk_text: str) -> List[Dict[str, Any]]:
    prompt = build_prompt(chunk_text)
    payload = {
        "model": os.getenv("EIA_LLM_MODEL", DEFAULT_MODEL),
        "temperature": 0,
        "messages": [
            {
                "role": "system",
                "content": (
                    "You extract environmental monitoring data. Return strict "
                    "JSON only. Never invent data."
                ),
            },
            {"role": "user", "content": prompt},
        ],
    }

    endpoint = os.getenv("EIA_LLM_ENDPOINT", DEFAULT_ENDPOINT)
    request = urllib.request.Request(
        endpoint,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {os.getenv('EIA_LLM_API_KEY')}",
            "Content-Type": "application/json",
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            data = json.loads(response.read().decode("utf-8"))
    except urllib.error.URLError as exc:
        raise RuntimeError(f"LLM request failed: {exc}") from exc

    content = data["choices"][0]["message"]["content"]
    return parse_json_records(content)


def build_prompt(chunk_text: str) -> str:
    fields = ", ".join(FIELDS)
    return f"""
从以下“单个切片原文”中提取环评现状监测数据。

必须遵守：
1. 只提取当前切片原文中明确出现的数据。
2. 不允许编造、补全或跨切片推断。
3. 缺失字段返回 null。
4. evidence 必须是当前切片原文中的短文本片段、完整表格行或能定位到记录的原文摘录。
5. confidence 是 0 到 1，越确定越高。
6. 只返回 JSON 数组，不要 Markdown，不要解释。
字段：{fields}

字段含义：
- source_type: 报告、监测方案、验收监测、未知等
- monitor_type: 环境空气、地表水、地下水、声环境、土壤、废气、废水等
- point: 监测点位或断面
- factor: 监测因子
- value: 监测值，保持原文数值或范围字符串
- unit: 单位
- standard_class: 标准类别或标准限值类别
- evidence: 支撑该条记录的原文
- confidence: 置信度

单个切片原文：
{chunk_text}
""".strip()


def parse_json_records(content: str) -> List[Dict[str, Any]]:
    content = content.strip()
    if content.startswith("```"):
        content = re.sub(r"^```(?:json)?", "", content).strip()
        content = re.sub(r"```$", "", content).strip()

    parsed = json.loads(content)
    if isinstance(parsed, dict) and "records" in parsed:
        parsed = parsed["records"]
    if not isinstance(parsed, list):
        raise ValueError("LLM response must be a JSON array or {'records': [...]}.")
    return parsed


def normalize_and_validate(
    records: List[Dict[str, Any]], chunk: Dict[str, Any]
) -> List[Dict[str, Any]]:
    normalized: List[Dict[str, Any]] = []
    chunk_text = chunk["text"]

    for record in records:
        item = {field: record.get(field) for field in FIELDS}
        if not evidence_is_supported(item, chunk_text):
            continue

        item["confidence"] = coerce_confidence(item.get("confidence"))
        item["chunk_id"] = chunk["chunk_id"]
        item["source_file"] = chunk["source_file"]
        normalized.append(item)

    return normalized


def evidence_is_supported(record: Dict[str, Any], chunk_text: str) -> bool:
    evidence = str(record.get("evidence") or "").strip()
    if not evidence:
        return False

    chunk_norm = normalize_text(chunk_text)
    evidence_norm = normalize_text(evidence)
    if evidence_norm and evidence_norm in chunk_norm:
        return True

    numbers = unique(NUMBER_RE.findall(str(record.get("value") or "")))
    if not numbers:
        numbers = meaningful_numbers(evidence)
    units = unique([str(record.get("unit") or "")])
    if not units:
        units = unique(match.group(0) for match in UNIT_RE.finditer(evidence))
    keywords = evidence_keywords(evidence, record)

    number_hits = sum(1 for number in numbers if normalize_text(number) in chunk_norm)
    unit_hits = sum(1 for unit in units if normalize_text(unit) in chunk_norm)
    keyword_hits = sum(1 for keyword in keywords if normalize_text(keyword) in chunk_norm)

    has_value_support = bool(numbers) and number_hits > 0
    has_unit_support = not units or unit_hits > 0
    has_keyword_support = keyword_hits >= min(2, len(keywords)) if keywords else True

    return has_value_support and has_unit_support and has_keyword_support


def meaningful_numbers(text: str) -> List[str]:
    numbers: List[str] = []
    for match in NUMBER_RE.finditer(text):
        start, end = match.span()
        before = text[start - 1] if start > 0 else ""
        after = text[end] if end < len(text) else ""
        if before.isalnum() or after.isalnum():
            continue
        numbers.append(match.group(0))
    return unique(numbers)


def normalize_text(text: str) -> str:
    text = text.lower()
    text = text.replace("μ", "µ")
    return re.sub(r"[\s|:：,，;；。()\[\]{}<>《》“”\"'`]+", "", text)


def evidence_keywords(evidence: str, record: Dict[str, Any]) -> List[str]:
    candidates: List[str] = []
    for field in ("point", "factor", "standard_class", "monitor_type", "source_type"):
        value = record.get(field)
        if value:
            candidates.append(str(value))

    candidates.extend(match.group(0) for match in FACTOR_RE.finditer(evidence))
    candidates.extend(re.findall(r"[\u4e00-\u9fffA-Za-z#]{2,}", evidence))

    ignored = {"row", "col", "单位", "监测值", "标准类别", "监测因子", "点位"}
    return [item for item in unique(candidates) if item not in ignored]


def unique(items: Any) -> List[str]:
    result: List[str] = []
    seen: set[str] = set()
    for item in items:
        text = str(item).strip()
        key = normalize_text(text)
        if not text or key in seen:
            continue
        seen.add(key)
        result.append(text)
    return result


def coerce_confidence(value: Any) -> float:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return 0.5
    return max(0.0, min(1.0, number))


def mock_extract(text: str) -> List[Dict[str, Any]]:
    """
    Local fallback for smoke tests and demos only.

    It extracts obvious table rows with headers or simple factor/value lines.
    Real extraction should be done by the LLM path.
    """
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    source_type = guess_source_type(text)
    monitor_type = guess_monitor_type(text)

    table_records = extract_from_table_lines(lines, source_type, monitor_type)
    if table_records:
        return table_records

    return extract_from_plain_lines(lines, source_type, monitor_type)


def extract_from_table_lines(
    lines: List[str], source_type: Optional[str], monitor_type: Optional[str]
) -> List[Dict[str, Any]]:
    records: List[Dict[str, Any]] = []
    header: Optional[Dict[str, int]] = None

    for line in lines:
        cells = split_table_line(line)
        if len(cells) < 3:
            continue

        detected_header = detect_header(cells)
        if detected_header:
            header = detected_header
            continue

        record = extract_from_cells(cells, line, source_type, monitor_type, header)
        if record:
            records.append(record)

    return records


def split_table_line(line: str) -> List[str]:
    line = re.sub(r"^row\s+\d+:\s*", "", line, flags=re.IGNORECASE)
    cells = []
    for cell in line.split("|"):
        cell = re.sub(r"^col\s+\d+:\s*", "", cell.strip(), flags=re.IGNORECASE)
        cells.append(cell.strip())
    return cells


def detect_header(cells: List[str]) -> Optional[Dict[str, int]]:
    header: Dict[str, int] = {}
    for index, cell in enumerate(cells):
        if re.search(r"点位|断面|位置|编号", cell):
            header["point"] = index
        if re.search(r"因子|项目|指标", cell):
            header["factor"] = index
        if re.search(r"监测值|检测值|浓度|结果|数值", cell):
            header["value"] = index
        if re.search(r"单位", cell):
            header["unit"] = index
        if re.search(r"标准|类别|限值", cell):
            header["standard_class"] = index

    return header if "value" in header and ("factor" in header or "point" in header) else None


def extract_from_cells(
    cells: List[str],
    evidence: str,
    source_type: Optional[str],
    monitor_type: Optional[str],
    header: Optional[Dict[str, int]],
) -> Optional[Dict[str, Any]]:
    if header:
        value_cell = get_cell(cells, header.get("value"))
        value, unit = split_value_unit(value_cell or "")
        if not value or not NUMBER_RE.search(value):
            return None

        unit = unit or parse_unit(get_cell(cells, header.get("unit")) or "")
        return {
            "source_type": source_type,
            "monitor_type": monitor_type,
            "point": get_cell(cells, header.get("point")),
            "factor": get_cell(cells, header.get("factor")),
            "value": value,
            "unit": unit,
            "standard_class": get_cell(cells, header.get("standard_class")),
            "evidence": evidence,
            "confidence": 0.62,
        }

    value_index = find_value_cell(cells)
    if value_index is None:
        return None

    value, unit = split_value_unit(cells[value_index])
    factor = find_factor(cells)
    point = find_point(cells, value_index, factor)

    return {
        "source_type": source_type,
        "monitor_type": monitor_type,
        "point": point,
        "factor": factor,
        "value": value,
        "unit": unit,
        "standard_class": find_standard_class(cells),
        "evidence": evidence,
        "confidence": 0.45,
    }


def get_cell(cells: List[str], index: Optional[int]) -> Optional[str]:
    if index is None or index >= len(cells):
        return None
    return cells[index] or None


def extract_from_plain_lines(
    lines: List[str], source_type: Optional[str], monitor_type: Optional[str]
) -> List[Dict[str, Any]]:
    records: List[Dict[str, Any]] = []
    pattern = re.compile(
        r"(?P<factor>COD|BOD5|氨氮|总磷|总氮|pH|PM10|PM2\.5|SO2|NO2|噪声|昼间|夜间)"
        r"[^\d<≤>=-]{0,20}"
        r"(?P<value>[<≤>=]?\d+(?:\.\d+)?(?:\s*[-~～]\s*\d+(?:\.\d+)?)?)"
        r"\s*(?P<unit>mg/L|mg/m3|µg/m3|ug/m3|dB\(A\)|dB|无量纲)?",
        re.IGNORECASE,
    )

    for line in lines:
        for match in pattern.finditer(line):
            records.append(
                {
                    "source_type": source_type,
                    "monitor_type": monitor_type,
                    "point": None,
                    "factor": match.group("factor"),
                    "value": match.group("value"),
                    "unit": match.group("unit"),
                    "standard_class": find_standard_class([line]),
                    "evidence": line,
                    "confidence": 0.5,
                }
            )
    return records


def find_value_cell(cells: List[str]) -> Optional[int]:
    for index, cell in enumerate(cells):
        stripped = cell.strip()
        if re.search(r"年份|日期|频次|编号|序号", stripped):
            continue
        if NUMBER_RE.fullmatch(stripped) or re.fullmatch(
            NUMBER_RE.pattern + r"\s*(?:" + UNIT_RE.pattern + r")?",
            stripped,
            re.IGNORECASE,
        ):
            return index
    return None


def split_value_unit(text: str) -> tuple[Optional[str], Optional[str]]:
    match = re.search(
        r"(?P<value>[<≤>=]?\d+(?:\.\d+)?(?:\s*[-~～]\s*\d+(?:\.\d+)?)?)\s*"
        r"(?P<unit>mg/L|mg/m3|µg/m3|ug/m3|dB\(A\)|dB|无量纲)?",
        text,
        re.IGNORECASE,
    )
    if not match:
        return None, None
    return match.group("value"), match.group("unit")


def parse_unit(text: str) -> Optional[str]:
    match = UNIT_RE.fullmatch(text.strip())
    return match.group(0) if match else None


def find_factor(cells: List[str]) -> Optional[str]:
    for cell in cells:
        match = FACTOR_RE.search(cell)
        if match:
            return match.group(0)
    return None


def find_point(cells: List[str], value_index: int, factor: Optional[str]) -> Optional[str]:
    for index, cell in enumerate(cells):
        if index == value_index or not cell:
            continue
        if factor and normalize_text(cell) == normalize_text(factor):
            continue
        if NUMBER_RE.search(cell) or UNIT_RE.search(cell) or find_standard_class([cell]):
            continue
        if re.search(r"点|断面|#|[A-Z]\d+", cell, re.IGNORECASE):
            return cell
    return None


def guess_source_type(text: str) -> Optional[str]:
    if "监测方案" in text:
        return "监测方案"
    if "监测报告" in text or "检测报告" in text:
        return "监测报告"
    if "验收监测" in text:
        return "验收监测"
    return None


def guess_monitor_type(text: str) -> Optional[str]:
    candidates = [
        "环境空气",
        "地表水",
        "地下水",
        "声环境",
        "土壤",
        "废气",
        "废水",
        "噪声",
    ]
    for item in candidates:
        if item in text:
            return "声环境" if item == "噪声" else item
    return None


def find_standard_class(cells: List[str]) -> Optional[str]:
    text = " ".join(cells)
    patterns = [
        r"(?:III|II|IV|V|I)类",
        r"(?:III|II|IV|V|I)级",
        r"[一二三四五1-5ⅠⅡⅢⅣⅤ]类",
        r"[一二三四五1-5ⅠⅡⅢⅣⅤ]级",
        r"GB\s*\d+(?:\.\d+)?[-—]\d+",
        r"《[^》]+》",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(0)
    return None
