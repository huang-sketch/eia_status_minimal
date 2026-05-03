import json
import http.client
import os
import re
import ssl
import time
import urllib.error
import urllib.request
from typing import Any, Dict, Iterable, List, Optional

import certifi

from models import ExtractedRecord, MonitorType


FIELDS = [
    "source_type",
    "monitor_type",
    "noise_type",
    "noise_type_label",
    "point",
    "sample_date",
    "factor",
    "value",
    "unit",
    "standard_class",
    "evidence",
    "confidence",
]

DEFAULT_ENDPOINT = "https://api.openai.com/v1/chat/completions"
DEFAULT_MODEL = "gpt-4o-mini"
LLM_TIMEOUT_SECONDS = int(os.getenv("EIA_LLM_TIMEOUT_SECONDS", "25"))
LLM_MAX_RETRIES = int(os.getenv("EIA_LLM_MAX_RETRIES", "1"))
LLM_RETRY_DELAY_SECONDS = 2
SSL_CONTEXT = ssl.create_default_context(cafile=certifi.where())

NUMBER_RE = re.compile(r"[<≤>=]?\d+(?:\.\d+)?(?:\s*[-~–]\s*\d+(?:\.\d+)?)?")
UNIT_RE = re.compile(
    r"mg/L|mg/m3|mg/m³|μg/m3|μg/m³|ug/m3|dB\(A\)|dB|无量纲|%",
    re.IGNORECASE,
)
FACTOR_RE = re.compile(
    r"COD|BOD5|氨氮|总磷|总氮|pH|溶解氧|高锰酸盐指数|石油类|"
    r"PM10|PM2\.5|SO2|NO2|CO|O3|TSP|非甲烷总烃|"
    r"LAeq|L10|L50|L90|Lmax|Lmin|噪声|昼间|夜间",
    re.IGNORECASE,
)
MONITOR_HINT_RE = re.compile(
    r"监测|检测|采样|断面|点位|因子|项目|浓度|结果|标准|限值|达标|"
    r"地表水|地下水|环境空气|声环境|噪声|土壤|废气|废水"
)


def extract_records_from_chunk(
    chunk: Dict[str, Any], logger: Optional[Any] = None
) -> List[Dict[str, Any]]:
    """
    Unified extraction entry point.

    Every chunk enters this function. The LLM decides whether the chunk contains
    monitoring data; rules are used as fallback when the LLM is unavailable,
    fails, or appears to miss obvious monitoring values.
    """
    text = chunk.get("text") or ""
    raw_records: List[Dict[str, Any]] = []
    llm_records: List[Dict[str, Any]] = []
    rule_records: List[Dict[str, Any]] = []
    used_method = "rule"

    if llm_enabled():
        try:
            llm_records = call_llm(text, chunk=chunk, logger=logger)
            raw_records.extend(mark_method(llm_records, "llm"))
            used_method = "llm"
        except Exception as exc:
            if logger:
                logger.log_failed_chunk(chunk, "llm_extract", exc, fallback_used=True)

    should_fallback = (
        not raw_records
        or chunk_looks_like_monitoring_data(text)
        and len(raw_records) < 2
    )
    if should_fallback:
        rule_records = fallback_extract(text)
        raw_records = merge_records(raw_records, mark_method(rule_records, "rule"))
        if rule_records and llm_records:
            used_method = "merged"

    normalized = normalize_and_validate(raw_records, chunk, default_method=used_method)
    if logger and chunk_looks_like_monitoring_data(text) and not normalized:
        logger.log_warning(chunk, "monitoring_like_chunk_no_records")
    return normalized


def llm_enabled() -> bool:
    return bool(os.getenv("EIA_LLM_API_KEY")) and os.getenv("EIA_LLM_DISABLE") != "1"


def detect_result_table_type(text: str) -> Optional[str]:
    """
    Backward-compatible hint only.

    This is no longer a gate for extraction. It is kept for callers/tests that
    still want a weak table-type hint.
    """
    first_lines = "\n".join(text.splitlines()[:3])
    if all(keyword in first_lines for keyword in ("采样位置", "采样日期", "检测项目", "检测结果")):
        return "surface_water"
    if all(keyword in first_lines for keyword in ("检测位置", "测量时间", "LAeq")):
        return "noise"
    return None


def call_llm(
    chunk_text: str,
    chunk: Optional[Dict[str, Any]] = None,
    logger: Optional[Any] = None,
) -> List[Dict[str, Any]]:
    prompt = build_prompt(chunk_text)
    payload = {
        "model": (chunk or {}).get("llm_model") or os.getenv("EIA_LLM_MODEL", DEFAULT_MODEL),
        "temperature": 0,
        "max_tokens": int(
            (chunk or {}).get("llm_max_tokens")
            or os.getenv("EIA_LLM_MAX_TOKENS", "2048")
        ),
        "messages": [
            {
                "role": "system",
                "content": (
                    "You are an environmental impact assessment data extraction "
                    "engine. Return strict JSON only. Never invent data."
                ),
            },
            {"role": "user", "content": prompt},
        ],
    }

    endpoint = os.getenv("EIA_LLM_ENDPOINT", DEFAULT_ENDPOINT)

    for attempt in range(1, LLM_MAX_RETRIES + 1):
        started = time.time()
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
            print(f"LLM call attempt {attempt}/{LLM_MAX_RETRIES}", flush=True)
            with urllib.request.urlopen(
                request,
                timeout=LLM_TIMEOUT_SECONDS,
                context=SSL_CONTEXT,
            ) as response:
                data = json.loads(response.read().decode("utf-8"))

            content = data["choices"][0]["message"]["content"]
            records = parse_json_records(content)
            if logger:
                logger.log_llm_call(
                    {
                        "chunk_id": (chunk or {}).get("chunk_id"),
                        "attempt": attempt,
                        "model": payload["model"],
                        "success": True,
                        "latency_ms": int((time.time() - started) * 1000),
                        "record_count": len(records),
                        "error": None,
                    }
                )
            print("LLM call succeeded", flush=True)
            return records
        except (
            urllib.error.URLError,
            http.client.RemoteDisconnected,
            OSError,
            TimeoutError,
            json.JSONDecodeError,
            KeyError,
            ValueError,
        ) as exc:
            if logger:
                logger.log_llm_call(
                    {
                        "chunk_id": (chunk or {}).get("chunk_id"),
                        "attempt": attempt,
                        "model": payload["model"],
                        "success": False,
                        "latency_ms": int((time.time() - started) * 1000),
                        "record_count": 0,
                        "error": str(exc),
                    }
                )
            print(f"LLM call failed on attempt {attempt}/{LLM_MAX_RETRIES}: {exc}", flush=True)
            if attempt < LLM_MAX_RETRIES:
                time.sleep(LLM_RETRY_DELAY_SECONDS * attempt)

    raise RuntimeError("LLM call failed after all retries")


def build_prompt(chunk_text: str) -> str:
    fields = ", ".join(FIELDS)
    return f"""
你是环评现状监测数据抽取助手。

只返回 JSON 数组，不要返回对象，不要 Markdown，不要解释。
如果没有监测数据，返回 []。

只抽取当前文本中明确出现、且有 evidence 支撑的记录。
不要编造、补全或跨 part 推断；缺失字段返回 null。

每条记录必须包含字段：{fields}

字段说明：
- source_type: 监测报告、监测方案、验收监测或未知
- monitor_type: surface_water, groundwater, ambient_air, acoustic, soil, wastewater, waste_gas, unknown
- point: 监测点位、断面或位置
- sample_date: 采样日期或监测时间
- factor: 监测因子
- value: 原文监测值
- unit: 原文单位
- standard_class: 标准类别、限值类别或标准名称
- evidence: 当前文本中的原文短句或完整表格行
- confidence: 0 到 1 的数字

当前文本：
{chunk_text}
""".strip()


def parse_json_records(content: str) -> List[Dict[str, Any]]:
    content = content.strip()
    if content.startswith("```"):
        content = re.sub(r"^```(?:json)?", "", content).strip()
        content = re.sub(r"```$", "", content).strip()

    parsed = json.loads(content)
    if isinstance(parsed, dict):
        if parsed.get("contains_monitoring_data") is False:
            return []
        parsed = parsed.get("records", [])
    if not isinstance(parsed, list):
        raise ValueError("LLM response must be a JSON array or an object with records.")
    return [item for item in parsed if isinstance(item, dict)]


def fallback_extract(text: str) -> List[Dict[str, Any]]:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    records = extract_from_table_lines(lines)
    plain_lines = [
        line
        for line in lines
        if not re.search(r"\brow\s+\d+:\s*|\bcol\s+\d+:", line, re.IGNORECASE)
    ]
    records.extend(extract_from_plain_lines(plain_lines))
    return dedupe_records(records)


def extract_from_table_lines(lines: List[str]) -> List[Dict[str, Any]]:
    records: List[Dict[str, Any]] = []
    header: Optional[Dict[str, int]] = None
    carry_point: Optional[str] = None
    carry_date: Optional[str] = None

    for line in lines:
        cells = split_table_line(line)
        if len(cells) < 2:
            continue

        detected_header = detect_header(cells)
        if detected_header:
            header = detected_header
            continue

        record = extract_from_cells(cells, line, header, carry_point, carry_date)
        if not record:
            continue

        if record.get("point"):
            carry_point = record["point"]
        if record.get("sample_date"):
            carry_date = record["sample_date"]
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
        if re.search(r"点位|断面|位置|编号|测点|采样点", cell):
            header["point"] = index
        if re.search(r"日期|时间|采样", cell):
            header["sample_date"] = index
        if re.search(r"因子|项目|指标|参数", cell):
            header["factor"] = index
        if re.search(r"监测值|检测值|浓度|结果|数值|测量值", cell):
            header["value"] = index
        if re.search(r"单位", cell):
            header["unit"] = index
        if re.search(r"标准|类别|限值|级别", cell):
            header["standard_class"] = index

    if "value" in header and ("factor" in header or "point" in header):
        return header
    return None


def extract_from_cells(
    cells: List[str],
    evidence: str,
    header: Optional[Dict[str, int]],
    carry_point: Optional[str],
    carry_date: Optional[str],
) -> Optional[Dict[str, Any]]:
    if looks_like_method_or_limit_row(cells):
        return None

    if header:
        value_cell = get_cell(cells, header.get("value"))
        value, unit = split_value_unit(value_cell or "")
        if not value:
            return None

        point = get_cell(cells, header.get("point")) or carry_point
        sample_date = get_cell(cells, header.get("sample_date")) or carry_date
        factor = get_cell(cells, header.get("factor")) or find_factor(cells)
        unit = unit or parse_unit(get_cell(cells, header.get("unit")) or "")

        return build_record(
            evidence=evidence,
            point=point,
            sample_date=sample_date,
            factor=factor,
            value=value,
            unit=unit,
            standard_class=get_cell(cells, header.get("standard_class")),
            confidence=0.62,
        )

    factor = find_factor(cells)
    value_index = find_value_cell(cells)
    if value_index is None or not factor:
        return None

    value, unit = split_value_unit(cells[value_index])
    if not value:
        return None

    return build_record(
        evidence=evidence,
        point=find_point(cells, value_index, factor) or carry_point,
        sample_date=find_date(cells) or carry_date,
        factor=factor,
        value=value,
        unit=unit,
        standard_class=find_standard_class(cells),
        confidence=0.48,
    )


def extract_from_plain_lines(lines: List[str]) -> List[Dict[str, Any]]:
    records: List[Dict[str, Any]] = []
    pattern = re.compile(
        r"(?P<factor>COD|BOD5|氨氮|总磷|总氮|pH|PM10|PM2\.5|SO2|NO2|"
        r"LAeq|L10|L50|L90|Lmax|Lmin|噪声|昼间|夜间)"
        r"[^\d<≤>=]{0,30}"
        r"(?P<value>[<≤>=]?\d+(?:\.\d+)?(?:\s*[-~–]\s*\d+(?:\.\d+)?)?)"
        r"\s*(?P<unit>mg/L|mg/m3|mg/m³|μg/m3|μg/m³|ug/m3|dB\(A\)|dB|无量纲|%)?",
        re.IGNORECASE,
    )

    for line in lines:
        for match in pattern.finditer(line):
            records.append(
                build_record(
                    evidence=line,
                    point=find_point([line], -1, match.group("factor")),
                    sample_date=find_date([line]),
                    factor=match.group("factor"),
                    value=match.group("value"),
                    unit=match.group("unit"),
                    standard_class=find_standard_class([line]),
                    confidence=0.5,
                )
            )
    return records


def build_record(
    evidence: str,
    point: Optional[str],
    sample_date: Optional[str],
    factor: Optional[str],
    value: Optional[str],
    unit: Optional[str],
    standard_class: Optional[str],
    confidence: float,
) -> Dict[str, Any]:
    return {
        "source_type": guess_source_type(evidence),
        "monitor_type": guess_monitor_type(
            " ".join(str(item or "") for item in [evidence, factor, unit, point])
        ),
        "point": point,
        "sample_date": sample_date,
        "factor": factor,
        "value": value,
        "unit": unit,
        "standard_class": standard_class,
        "evidence": evidence,
        "confidence": confidence,
    }


def normalize_and_validate(
    records: List[Dict[str, Any]],
    chunk: Dict[str, Any],
    default_method: str = "llm",
) -> List[Dict[str, Any]]:
    normalized: List[Dict[str, Any]] = []
    chunk_text = chunk.get("text") or ""

    for record in records:
        item = {field: record.get(field) for field in FIELDS}
        if not item.get("evidence"):
            continue
        if not evidence_is_supported(item, chunk_text):
            item["needs_review"] = True

        item["confidence"] = coerce_confidence(item.get("confidence"))
        item["chunk_id"] = chunk["chunk_id"]
        item["source_file"] = chunk["source_file"]
        item["extraction_method"] = record.get("extraction_method") or default_method
        item["monitor_type"] = normalize_monitor_type(item.get("monitor_type"))
        item["needs_review"] = bool(item.get("needs_review")) or item["confidence"] < 0.55

        try:
            normalized.append(ExtractedRecord(**item).model_dump(mode="json"))
        except AttributeError:
            normalized.append(ExtractedRecord(**item).dict())
        except Exception:
            continue

    return dedupe_records(normalized)


def evidence_is_supported(record: Dict[str, Any], chunk_text: str) -> bool:
    raw_evidence = record.get("evidence")
    if isinstance(raw_evidence, (dict, list)):
        evidence = json.dumps(raw_evidence, ensure_ascii=False)
    else:
        evidence = str(raw_evidence or "").strip()
    if not evidence:
        return False

    chunk_norm = normalize_text(chunk_text)
    evidence_norm = normalize_text(evidence)
    if evidence_norm and evidence_norm in chunk_norm:
        return True

    numbers = unique(NUMBER_RE.findall(str(record.get("value") or "")))
    if not numbers:
        numbers = meaningful_numbers(evidence)
    keywords = evidence_keywords(evidence, record)

    number_hits = sum(1 for number in numbers if normalize_text(number) in chunk_norm)
    keyword_hits = sum(1 for keyword in keywords if normalize_text(keyword) in chunk_norm)

    has_value_support = bool(numbers) and number_hits > 0
    has_keyword_support = keyword_hits >= 1 if keywords else True
    return has_value_support and has_keyword_support


def chunk_looks_like_monitoring_data(text: str) -> bool:
    return bool(
        MONITOR_HINT_RE.search(text)
        and NUMBER_RE.search(text)
        and (UNIT_RE.search(text) or FACTOR_RE.search(text))
    )


def mark_method(records: Iterable[Dict[str, Any]], method: str) -> List[Dict[str, Any]]:
    result = []
    for record in records:
        item = dict(record)
        item["extraction_method"] = method
        result.append(item)
    return result


def merge_records(
    primary: List[Dict[str, Any]],
    fallback: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    return dedupe_records([*primary, *fallback])


def dedupe_records(records: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    result: List[Dict[str, Any]] = []
    seen: set[str] = set()
    for record in records:
        key = "|".join(
            normalize_text(str(record.get(field) or ""))
            for field in ("point", "sample_date", "factor", "value", "unit", "evidence")
        )
        if key in seen:
            continue
        seen.add(key)
        result.append(record)
    return result


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
    text = text.lower().replace("μ", "u").replace("µ", "u")
    return re.sub(r"[\s|:：；;，,。()\[\]{}<>《》“”\"'`]+", "", text)


def evidence_keywords(evidence: str, record: Dict[str, Any]) -> List[str]:
    candidates: List[str] = []
    for field in ("point", "factor", "standard_class", "monitor_type", "source_type"):
        value = record.get(field)
        if value:
            candidates.append(str(value))
    candidates.extend(match.group(0) for match in FACTOR_RE.finditer(evidence))
    candidates.extend(re.findall(r"[\u4e00-\u9fffA-Za-z#]{2,}", evidence))
    ignored = {"row", "col", "单位", "监测值", "检测值", "标准类别", "监测因子", "点位"}
    return [item for item in unique(candidates) if item not in ignored]


def unique(items: Iterable[Any]) -> List[str]:
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


def normalize_monitor_type(value: Any) -> str:
    text = str(value or "").strip()
    mapping = {
        "地表水": MonitorType.surface_water.value,
        "surface_water": MonitorType.surface_water.value,
        "地下水": MonitorType.groundwater.value,
        "groundwater": MonitorType.groundwater.value,
        "环境空气": MonitorType.ambient_air.value,
        "ambient_air": MonitorType.ambient_air.value,
        "空气": MonitorType.ambient_air.value,
        "声环境": MonitorType.acoustic.value,
        "噪声": MonitorType.acoustic.value,
        "acoustic": MonitorType.acoustic.value,
        "土壤": MonitorType.soil.value,
        "soil": MonitorType.soil.value,
        "废水": MonitorType.wastewater.value,
        "wastewater": MonitorType.wastewater.value,
        "废气": MonitorType.waste_gas.value,
        "waste_gas": MonitorType.waste_gas.value,
    }
    return mapping.get(text, MonitorType.unknown.value)


def get_cell(cells: List[str], index: Optional[int]) -> Optional[str]:
    if index is None or index >= len(cells):
        return None
    return cells[index] or None


def find_value_cell(cells: List[str]) -> Optional[int]:
    for index, cell in enumerate(cells):
        stripped = cell.strip()
        if re.search(r"年份|日期|时间|频次|编号|序号", stripped):
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
        r"(?P<value>[<≤>=]?\d+(?:\.\d+)?(?:\s*[-~–]\s*\d+(?:\.\d+)?)?)\s*"
        r"(?P<unit>mg/L|mg/m3|mg/m³|μg/m3|μg/m³|ug/m3|dB\(A\)|dB|无量纲|%)?",
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


def looks_like_method_or_limit_row(cells: List[str]) -> bool:
    text = " ".join(cells)
    has_method_standard = bool(
        re.search(r"\b(?:GB/T|GB|HJ|SL)\s*[/T]*\s*\d+", text, re.IGNORECASE)
    )
    has_method_words = bool(re.search(r"测定|方法|检出限|仪器|标准号|分析方法", text))
    has_result_context = bool(
        re.search(r"点位|断面|采样|监测时间|监测日期|20\d{2}[-/.年]\d{1,2}", text)
    )
    return has_method_standard and has_method_words and not has_result_context


def find_date(cells: List[str]) -> Optional[str]:
    text = " ".join(cells)
    match = re.search(r"\d{4}[-/.年]\d{1,2}[-/.月]\d{1,2}日?", text)
    if match:
        return match.group(0)
    match = re.search(r"\d{1,2}月\d{1,2}日", text)
    return match.group(0) if match else None


def guess_source_type(text: str) -> Optional[str]:
    if "监测方案" in text:
        return "监测方案"
    if "监测报告" in text or "检测报告" in text:
        return "监测报告"
    if "验收监测" in text:
        return "验收监测"
    return None


def guess_monitor_type(text: str) -> str:
    candidates = [
        ("地表水", MonitorType.surface_water.value),
        ("地下水", MonitorType.groundwater.value),
        ("环境空气", MonitorType.ambient_air.value),
        ("空气", MonitorType.ambient_air.value),
        ("声环境", MonitorType.acoustic.value),
        ("噪声", MonitorType.acoustic.value),
        ("土壤", MonitorType.soil.value),
        ("废气", MonitorType.waste_gas.value),
        ("废水", MonitorType.wastewater.value),
    ]
    for keyword, value in candidates:
        if keyword in text:
            return value
    if re.search(r"水质|河|湖|断面|溶解氧|氨氮|总磷|总氮|高锰酸盐指数|石油类", text):
        return MonitorType.surface_water.value
    if re.search(r"LAeq|L10|L50|L90|Lmax|Lmin|dB", text, re.IGNORECASE):
        return MonitorType.acoustic.value
    return MonitorType.unknown.value


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
