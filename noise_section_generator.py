import json
import http.client
import os
import re
import time
import urllib.error
import urllib.request
from collections import OrderedDict
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

from dotenv import load_dotenv
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from llm_extractor import (
    DEFAULT_ENDPOINT,
    LLM_MAX_RETRIES,
    LLM_RETRY_DELAY_SECONDS,
    LLM_TIMEOUT_SECONDS,
    SSL_CONTEXT,
)
from word_processor import load_docx_chunks

load_dotenv()

INPUT_DIR = Path(os.getenv("EIA_INPUT_DIR", "input"))
OUTPUT_DIR = Path(os.getenv("EIA_OUTPUT_DIR", "output"))
DEBUG_DIR = OUTPUT_DIR / "debug_tables"
OUTPUT_DOCX = OUTPUT_DIR / "noise_section.docx"
ENABLE_LLM_TEXT_POLISH = os.getenv("ENABLE_LLM_TEXT_POLISH", "false").lower() == "true"

POINT_KEY = "点位"
TIME_KEY = "监测时间"
FLOW_KEY = "车流量（辆/20min）"

DAY_LIMITS = {"1类": 55, "2类": 60, "3类": 65, "4a类": 70}
NIGHT_LIMITS = {"1类": 45, "2类": 50, "3类": 55, "4a类": 55}


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    DEBUG_DIR.mkdir(parents=True, exist_ok=True)

    plan_rows = parse_noise_plan()
    area_records = load_flattened_noise("flattened_table_1.json", "area_environment_noise")
    traffic_records = load_flattened_noise("flattened_table_0.json", "traffic_noise")
    all_records = [*area_records, *traffic_records]

    validate_plan_report_match(all_records, plan_rows)
    table1 = build_monitor_points_table(plan_rows)
    table2 = build_sensitive_result_table(all_records, plan_rows)
    table3 = build_attenuation_result_table(traffic_records, plan_rows)
    summary = build_compliance_summary(table2, table3)

    write_json(DEBUG_DIR / "noise_monitor_points_table.json", table1)
    write_json(DEBUG_DIR / "noise_sensitive_points_result_table.json", table2)
    write_json(DEBUG_DIR / "traffic_noise_attenuation_table.json", table3)
    write_json(DEBUG_DIR / "noise_compliance_summary.json", summary)

    texts = build_rule_texts(table1, table2, table3, summary)
    texts = polish_noise_text_with_llm(texts, table1, table2, table3, summary)

    doc = build_docx(table1, table2, table3, summary, texts)
    doc.save(OUTPUT_DOCX)
    print(f"generated: {OUTPUT_DOCX}")
    print(f"table4.2-1 rows: {len(table1['rows'])}")
    print(f"table4.2-2 rows: {len(table2['rows'])}")
    print(f"table4.2-3 rows: {len(table3['rows'])}")
    print(f"sensitive exceed count: {summary['sensitive']['exceed_count']}")
    print(f"attenuation exceed count: {summary['attenuation']['exceed_count']}")


def parse_noise_plan() -> List[Dict[str, Any]]:
    plan_path = next(
        path for path in INPUT_DIR.glob("*.docx")
        if "方案" in path.name and not path.name.startswith("~$")
    )
    chunks = load_docx_chunks(plan_path)
    for chunk in chunks:
        if chunk.get("kind") != "table":
            continue
        title = (chunk.get("metadata") or {}).get("table_title") or ""
        if "声环境" not in title:
            continue
        rows = parse_table_rows(chunk.get("text") or "")
        headers = rows[0]
        return [
            normalize_plan_row(row_to_dict(headers, row), source_file=str(plan_path), source_table=chunk["chunk_id"])
            for row in rows[1:]
            if len(row) >= 8
        ]
    raise RuntimeError("未找到声环境现状监测方案表")


def normalize_plan_row(row: Dict[str, Any], source_file: str, source_table: str) -> Dict[str, Any]:
    code = str(row.get("监测点编号") or "").strip()
    return {
        "point_code": code,
        "station": row.get("桩号") or "",
        "point_name": row.get("监测点名称") or "",
        "district": row.get("行政区划") or "",
        "position": row.get("监测点位置") or "",
        "standard_class": row.get("现状标准") or "",
        "factor": row.get("监测因子") or "",
        "frequency": row.get("监测频次") or "",
        "is_attenuation": is_attenuation_text(" ".join(str(v) for v in row.values())),
        "source_file": source_file,
        "source_table": source_table,
    }


def load_flattened_noise(filename: str, noise_type: str) -> List[Dict[str, Any]]:
    path = DEBUG_DIR / filename
    if not path.exists():
        raise FileNotFoundError(f"缺少噪声扁平化表: {path}")
    data = json.loads(path.read_text(encoding="utf-8"))
    records: List[Dict[str, Any]] = []
    for index, row in enumerate(data.get("records") or []):
        point_text = str(row.get(POINT_KEY) or "").strip()
        laeq_key = next((key for key in row if key.endswith("_LAeq")), None)
        if not point_text or not laeq_key:
            continue
        records.append(
            {
                "noise_type": noise_type,
                "point_text": point_text,
                "raw_point_code": extract_point_code(point_text),
                "monitor_time": row.get(TIME_KEY) or "",
                "period": infer_period(row.get(TIME_KEY) or ""),
                "laeq": row.get(laeq_key),
                "traffic_flow": row.get(FLOW_KEY) or "/",
                "source_order": index,
                "raw": row,
            }
        )
    return records


def build_monitor_points_table(plan_rows: List[Dict[str, Any]]) -> Dict[str, Any]:
    headers = ["监测点编号", "桩号", "监测点名称", "行政区划", "监测点位置", "现状标准", "监测因子", "监测频次"]
    rows = [
        {
            "监测点编号": row["point_code"],
            "桩号": row["station"],
            "监测点名称": row["point_name"],
            "行政区划": row["district"],
            "监测点位置": row["position"],
            "现状标准": row["standard_class"],
            "监测因子": row["factor"],
            "监测频次": row["frequency"],
        }
        for row in plan_rows
    ]
    return {"title": "表4.2-1 声环境质量现状监测点", "headers": headers, "rows": rows}


def build_sensitive_result_table(
    records: List[Dict[str, Any]],
    plan_rows: List[Dict[str, Any]],
) -> Dict[str, Any]:
    headers = result_headers()
    matched = match_records(records, plan_rows)
    grouped = group_records_by_point(matched)
    rows: List[Dict[str, Any]] = []
    warnings: List[str] = []
    for plan in sorted([row for row in plan_rows if not row.get("is_attenuation")], key=lambda row: point_sort_key(row["point_code"])):
        code = plan["point_code"]
        if code not in grouped:
            raise ValueError(f"监测方案点位 {code} 未在噪声监测结果中匹配到数据")
        plan = grouped[code]["plan"]
        if plan.get("is_attenuation"):
            continue
        point_records = grouped[code]["records"]
        metrics = compute_point_metrics(point_records, plan)
        row = base_result_row(plan)
        row.update(metrics)
        flows = paired_records(point_records)
        row["traffic_flow_day1_day"] = flow_text(flows, 0, "day")
        row["traffic_flow_day1_night"] = flow_text(flows, 0, "night")
        row["traffic_flow_day2_day"] = flow_text(flows, 1, "day")
        row["traffic_flow_day2_night"] = flow_text(flows, 1, "night")
        row["noise_types"] = sorted({record["noise_type"] for record in point_records})
        row["needs_review"] = bool(metrics.get("warnings"))
        row["warning"] = "；".join(metrics.get("warnings") or [])
        rows.append(row)
        warnings.extend(metrics.get("warnings") or [])
    return {"title": "表4.2-2 项目沿线敏感点声环境现状监测平均值  单位：dB(A)", "headers": headers, "rows": rows, "warnings": warnings}


def build_attenuation_result_table(
    records: List[Dict[str, Any]],
    plan_rows: List[Dict[str, Any]],
) -> Dict[str, Any]:
    headers = result_headers()
    matched = match_records(records, plan_rows)
    grouped = group_records_by_point(matched)
    rows: List[Dict[str, Any]] = []
    warnings: List[str] = []
    for code in sorted(grouped, key=point_sort_key):
        plan = grouped[code]["plan"]
        if not plan.get("is_attenuation"):
            continue
        point_records = grouped[code]["records"]
        metrics = compute_point_metrics(point_records, plan)
        row = base_result_row(plan)
        row.update(metrics)
        flows = paired_records(point_records)
        row["traffic_flow_day1_day"] = flow_text(flows, 0, "day")
        row["traffic_flow_day1_night"] = flow_text(flows, 0, "night")
        row["traffic_flow_day2_day"] = flow_text(flows, 1, "day")
        row["traffic_flow_day2_night"] = flow_text(flows, 1, "night")
        row["needs_review"] = bool(metrics.get("warnings"))
        row["warning"] = "；".join(metrics.get("warnings") or [])
        rows.append(row)
        warnings.extend(metrics.get("warnings") or [])
    return {"title": "表4.2-3 现状公路交通噪声衰减断面监测结果  单位：dB(A)", "headers": headers, "rows": rows, "warnings": warnings}


def result_headers() -> List[str]:
    return [
        "监测点编号",
        "监测点名称",
        "监测点位置",
        "day1_day",
        "day1_night",
        "day2_day",
        "day2_night",
        "avg_day",
        "avg_night",
        "standard_day",
        "standard_night",
        "exceed_day",
        "exceed_night",
        "traffic_flow_day1_day",
        "traffic_flow_day1_night",
        "traffic_flow_day2_day",
        "traffic_flow_day2_night",
    ]


def match_records(records: List[Dict[str, Any]], plan_rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    result = []
    for record in records:
        plan = match_plan(record, plan_rows)
        item = dict(record)
        item["plan"] = plan
        result.append(item)
    return result


def validate_plan_report_match(records: List[Dict[str, Any]], plan_rows: List[Dict[str, Any]]) -> None:
    matched = match_records(records, plan_rows)
    grouped = group_records_by_point(matched)
    plan_codes = {row["point_code"] for row in plan_rows}
    report_codes = set(grouped)
    missing_in_report = sorted(plan_codes - report_codes, key=point_sort_key)
    extra_in_report = sorted(report_codes - plan_codes, key=point_sort_key)
    unmatched_records = [
        {
            "point_text": record.get("point_text"),
            "raw_point_code": record.get("raw_point_code"),
            "noise_type": record.get("noise_type"),
        }
        for record in matched
        if not (record.get("plan") or {}).get("point_code")
    ]
    payload = {
        "missing_in_report": missing_in_report,
        "extra_in_report": extra_in_report,
        "unmatched_records": unmatched_records,
        "plan_count": len(plan_codes),
        "report_count": len(report_codes),
    }
    write_json(DEBUG_DIR / "noise_plan_report_mismatch.json", payload)
    if missing_in_report or extra_in_report or unmatched_records:
        raise ValueError(
            "监测方案与监测报告噪声点位不一致，详见 "
            f"{(DEBUG_DIR / 'noise_plan_report_mismatch.json').resolve()}"
        )


def match_plan(record: Dict[str, Any], plan_rows: List[Dict[str, Any]]) -> Dict[str, Any]:
    raw_code = record.get("raw_point_code")
    point_text = record.get("point_text") or ""
    if raw_code:
        candidates = [row for row in plan_rows if row["point_code"] == raw_code]
        if len(candidates) == 1:
            return candidates[0]
        candidates = [row for row in plan_rows if row["point_code"].startswith(raw_code + "-")]
        if len(candidates) == 1:
            return candidates[0]
        if len(candidates) > 1:
            by_name = [row for row in candidates if row["point_name"] and row["point_name"] in point_text]
            if by_name:
                return by_name[0]
            by_name = [
                row for row in candidates
                if row["point_name"] and len(row["point_name"].rstrip("1234567890")) >= 2 and row["point_name"].rstrip("1234567890") in point_text
            ]
            if by_name:
                return by_name[0]
    for row in plan_rows:
        name = row["point_name"]
        station = row["station"]
        if name and name in point_text:
            return row
        if station and station in point_text and name and name.rstrip("12") in point_text:
            return row
    if is_attenuation_text(point_text):
        distance = extract_distance(point_text)
        for row in plan_rows:
            if row["is_attenuation"] and distance and distance in row["position"]:
                return row
    return {
        "point_code": raw_code or "",
        "station": "",
        "point_name": point_text,
        "district": "",
        "position": point_text,
        "standard_class": "",
        "factor": "",
        "frequency": "",
        "is_attenuation": is_attenuation_text(point_text),
    }


def group_records_by_point(records: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    grouped: "OrderedDict[str, Dict[str, Any]]" = OrderedDict()
    for record in records:
        plan = record["plan"]
        code = plan.get("point_code") or record.get("raw_point_code") or record.get("point_text")
        if code not in grouped:
            grouped[code] = {"plan": plan, "records": []}
        grouped[code]["records"].append(record)
    for value in grouped.values():
        value["records"].sort(key=lambda item: item["source_order"])
    return grouped


def base_result_row(plan: Dict[str, Any]) -> Dict[str, Any]:
    return {
        "监测点编号": plan.get("point_code") or "-",
        "监测点名称": plan.get("point_name") or "-",
        "监测点位置": plan.get("position") or "-",
    }


def compute_point_metrics(records: List[Dict[str, Any]], plan: Dict[str, Any]) -> Dict[str, Any]:
    pairs = paired_records(records)
    warnings: List[str] = []
    day_values: List[int] = []
    night_values: List[int] = []
    result: Dict[str, Any] = {}
    for day_index in range(2):
        day_record = pairs[day_index].get("day") if day_index < len(pairs) else None
        night_record = pairs[day_index].get("night") if day_index < len(pairs) else None
        day_value = parse_int(day_record.get("laeq")) if day_record else None
        night_value = parse_int(night_record.get("laeq")) if night_record else None
        result[f"day{day_index + 1}_day"] = value_or_dash(day_value)
        result[f"day{day_index + 1}_night"] = value_or_dash(night_value)
        if day_value is not None:
            day_values.append(day_value)
        else:
            warnings.append(f"{plan.get('point_code')} 第{day_index + 1}天昼间缺失")
        if night_value is not None:
            night_values.append(night_value)
        else:
            warnings.append(f"{plan.get('point_code')} 第{day_index + 1}天夜间缺失")
    avg_day = round_half_up(sum(day_values) / len(day_values)) if day_values else None
    avg_night = round_half_up(sum(night_values) / len(night_values)) if night_values else None
    standard_class = plan.get("standard_class")
    standard_day = DAY_LIMITS.get(standard_class)
    standard_night = NIGHT_LIMITS.get(standard_class)
    if standard_day is None or standard_night is None:
        warnings.append(f"{plan.get('point_code')} 未识别现状标准: {standard_class}")
    result["avg_day"] = value_or_dash(avg_day)
    result["avg_night"] = value_or_dash(avg_night)
    result["standard_day"] = value_or_dash(standard_day)
    result["standard_night"] = value_or_dash(standard_night)
    result["exceed_day"] = exceed_text(avg_day, standard_day)
    result["exceed_night"] = exceed_text(avg_night, standard_night)
    result["warnings"] = warnings
    return result


def paired_records(records: List[Dict[str, Any]]) -> List[Dict[str, Dict[str, Any]]]:
    pairs: List[Dict[str, Dict[str, Any]]] = []
    current: Dict[str, Dict[str, Any]] = {}
    for record in records:
        period = record.get("period")
        if period == "day":
            if current:
                pairs.append(current)
            current = {"day": record}
        elif period == "night":
            if not current:
                current = {}
            current["night"] = record
        else:
            if not current:
                current = {}
            current.setdefault("unknown", record)
    if current:
        pairs.append(current)
    return pairs[:2]


def build_compliance_summary(table2: Dict[str, Any], table3: Dict[str, Any]) -> Dict[str, Any]:
    return {
        "sensitive": summarize_rows(table2["rows"]),
        "attenuation": summarize_rows(table3["rows"]),
    }


def summarize_rows(rows: List[Dict[str, Any]]) -> Dict[str, Any]:
    exceed_items = []
    for row in rows:
        for period, field in [("昼间", "exceed_day"), ("夜间", "exceed_night")]:
            value = row.get(field)
            if isinstance(value, int) and value > 0:
                exceed_items.append(
                    {
                        "point_code": row.get("监测点编号"),
                        "point_name": row.get("监测点名称"),
                        "period": period,
                        "exceed": value,
                    }
                )
    return {
        "total_count": len(rows),
        "exceed_count": len(exceed_items),
        "exceed_items": exceed_items,
    }


def build_rule_texts(
    table1: Dict[str, Any],
    table2: Dict[str, Any],
    table3: Dict[str, Any],
    summary: Dict[str, Any],
) -> Dict[str, str]:
    return {
        "monitoring_plan_text": (
            f"根据《环境影响评价技术导则 公路建设项目》（HJ1358-2024）要求，结合项目沿线敏感点分布、现状噪声源及代表性区段，共设置{len(table1['rows'])}个声环境监测点。声环境现状监测方案见表4.2-1。"
        ),
        "monitoring_result_text": (
            "监测期间具体测量时间段、测量仪器、测量方法均按规范要求进行。测量结果以等效连续A声级和统计噪声级给出，并以等效A声级作为最终评价量。"
        ),
        "sensitive_conclusion": sensitive_conclusion(summary["sensitive"]),
        "traffic_intro": "本次评价对现状公路交通噪声衰减断面进行监测，监测结果见表4.2-3。",
        "attenuation_conclusion": attenuation_conclusion(summary["attenuation"], lead_only=False),
    }


def polish_noise_text_with_llm(
    rule_texts: Dict[str, str],
    table1: Dict[str, Any],
    table2: Dict[str, Any],
    table3: Dict[str, Any],
    summary: Dict[str, Any],
) -> Dict[str, str]:
    payload = {
        "enabled": ENABLE_LLM_TEXT_POLISH,
        "reference_text": read_reference_noise_text(),
        "rule_texts": rule_texts,
        "summary": summary,
        "table_stats": {
            "table4_2_1_rows": len(table1["rows"]),
            "table4_2_2_rows": len(table2["rows"]),
            "table4_2_3_rows": len(table3["rows"]),
        },
    }
    write_json(DEBUG_DIR / "noise_llm_text_input.json", payload)

    if not ENABLE_LLM_TEXT_POLISH:
        validation = {"used_llm": False, "valid": True, "warnings": ["ENABLE_LLM_TEXT_POLISH=false"]}
        write_json(DEBUG_DIR / "noise_llm_text_output.json", {})
        write_json(DEBUG_DIR / "noise_llm_text_validation.json", validation)
        return rule_texts

    if not os.getenv("EIA_LLM_API_KEY"):
        validation = {"used_llm": False, "valid": False, "warnings": ["EIA_LLM_API_KEY is missing"]}
        write_json(DEBUG_DIR / "noise_llm_text_output.json", {})
        write_json(DEBUG_DIR / "noise_llm_text_validation.json", validation)
        return rule_texts

    try:
        polished = call_text_polish_llm(build_text_polish_prompt(payload))
        validation = validate_polished_text(polished, summary)
        write_json(DEBUG_DIR / "noise_llm_text_output.json", polished)
        write_json(DEBUG_DIR / "noise_llm_text_validation.json", validation)
        if validation["valid"]:
            return clean_polished_texts({**rule_texts, **{key: str(polished[key]).strip() for key in rule_texts}})
        return rule_texts
    except Exception as exc:
        write_json(DEBUG_DIR / "noise_llm_text_output.json", {})
        write_json(
            DEBUG_DIR / "noise_llm_text_validation.json",
            {"used_llm": True, "valid": False, "warnings": [str(exc)]},
        )
        return rule_texts


def clean_polished_texts(texts: Dict[str, str]) -> Dict[str, str]:
    cleaned = dict(texts)
    for prefix in ("监测结果见表4.2-2。", "监测结果见表4.2-2，"):
        value = cleaned.get("sensitive_conclusion", "")
        if value.startswith(prefix):
            cleaned["sensitive_conclusion"] = value[len(prefix):].lstrip()
    return cleaned


def read_reference_noise_text() -> List[str]:
    reference_path = Path("D:/") / "华设" / "智能体" / "环评报告" / "输出" / "台儿庄高速" / "声环境现状调查与评价.docx"
    if not reference_path.exists():
        return []
    try:
        doc = Document(str(reference_path))
        return [paragraph.text.strip() for paragraph in doc.paragraphs if paragraph.text.strip()]
    except Exception:
        return []


def build_text_polish_prompt(payload: Dict[str, Any]) -> str:
    return (
        "你是环评报告声环境章节文字润色助手。只润色文字，不判断达标，不修改任何数字、表号、点位编号、超标量或标准名称。\n"
        "请参考 reference_text 的中文报告风格，润色 rule_texts 中的 5 个字段。\n"
        "硬性要求：\n"
        "1. 只输出 JSON object，不要 markdown，不要解释。\n"
        "2. JSON 必须且只能包含 monitoring_plan_text、monitoring_result_text、sensitive_conclusion、traffic_intro、attenuation_conclusion。\n"
        "3. 必须保留表4.2-1、表4.2-2、表4.2-3。\n"
        "4. 必须保留输入 summary 中所有超标点位编号和超标量。\n"
        "5. 不得新增超标原因，因为输入未提供原因字段。\n"
        "6. 若存在超标，不得写“均满足”或“均达标”。\n"
        "输入如下：\n"
        f"{json.dumps(payload, ensure_ascii=False)}"
    )


def call_text_polish_llm(prompt: str) -> Dict[str, Any]:
    endpoint = os.getenv("EIA_LLM_ENDPOINT", DEFAULT_ENDPOINT)
    model = os.getenv("EIA_LLM_MODEL", "qwen-plus")
    request_payload = {
        "model": model,
        "temperature": 0.2,
        "max_tokens": int(os.getenv("EIA_LLM_TEXT_MAX_TOKENS", "1800")),
        "messages": [
            {
                "role": "system",
                "content": "你是严谨的环评报告文本润色助手。必须输出严格 JSON，禁止编造事实。",
            },
            {"role": "user", "content": prompt},
        ],
    }
    last_error: Optional[Exception] = None
    for attempt in range(1, LLM_MAX_RETRIES + 1):
        started = time.time()
        req = urllib.request.Request(
            endpoint,
            data=json.dumps(request_payload, ensure_ascii=False).encode("utf-8"),
            headers={
                "Authorization": f"Bearer {os.getenv('EIA_LLM_API_KEY')}",
                "Content-Type": "application/json",
            },
            method="POST",
        )
        try:
            print(f"LLM text polish attempt {attempt}/{LLM_MAX_RETRIES}", flush=True)
            with urllib.request.urlopen(req, timeout=LLM_TIMEOUT_SECONDS, context=SSL_CONTEXT) as response:
                data = json.loads(response.read().decode("utf-8"))
            content = data["choices"][0]["message"]["content"]
            parsed = parse_json_object(content)
            print(f"LLM text polish succeeded in {int((time.time() - started) * 1000)}ms", flush=True)
            return parsed
        except (
            urllib.error.URLError,
            http.client.RemoteDisconnected,
            OSError,
            TimeoutError,
            json.JSONDecodeError,
            KeyError,
            ValueError,
        ) as exc:
            last_error = exc
            print(f"LLM text polish failed on attempt {attempt}/{LLM_MAX_RETRIES}: {exc}", flush=True)
            if attempt < LLM_MAX_RETRIES:
                time.sleep(LLM_RETRY_DELAY_SECONDS * attempt)
    raise RuntimeError(f"LLM text polish failed: {last_error}")


def parse_json_object(content: str) -> Dict[str, Any]:
    text = content.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
    start = text.find("{")
    end = text.rfind("}")
    if start >= 0 and end > start:
        text = text[start : end + 1]
    parsed = json.loads(text)
    if not isinstance(parsed, dict):
        raise ValueError("LLM text polish response must be a JSON object")
    return parsed


def validate_polished_text(polished: Dict[str, Any], summary: Dict[str, Any]) -> Dict[str, Any]:
    required = [
        "monitoring_plan_text",
        "monitoring_result_text",
        "sensitive_conclusion",
        "traffic_intro",
        "attenuation_conclusion",
    ]
    warnings: List[str] = []
    for key in required:
        if not isinstance(polished.get(key), str) or not polished.get(key, "").strip():
            warnings.append(f"missing or empty field: {key}")

    all_text = "\n".join(
        [
            *(str(polished.get(key, "")) for key in required),
            "监测结果见表4.2-2。",
        ]
    )
    for table_no in ("表4.2-1", "表4.2-2", "表4.2-3"):
        if table_no not in all_text:
            warnings.append(f"missing table number: {table_no}")

    for section in ("sensitive", "attenuation"):
        for item in summary.get(section, {}).get("exceed_items", []):
            point_code = str(item.get("point_code") or "")
            exceed_text = f"{item.get('exceed')}dB(A)"
            alt_exceed_text = f"{item.get('exceed')}dB（A）"
            if point_code and point_code not in all_text:
                warnings.append(f"missing exceed point: {point_code}")
            if exceed_text not in all_text and alt_exceed_text not in all_text:
                warnings.append(f"missing exceed amount: {exceed_text}")

    if "原因" in str(polished.get("sensitive_conclusion", "")) or "原因" in str(polished.get("attenuation_conclusion", "")):
        warnings.append("text adds exceed reason not present in summary")
    for key in ("sensitive", "attenuation"):
        if summary.get(key, {}).get("exceed_items"):
            target_field = "sensitive_conclusion" if key == "sensitive" else "attenuation_conclusion"
            text = str(polished.get(target_field, ""))
            if "均满足" in text or "均达标" in text:
                warnings.append(f"{target_field} says all compliant despite exceed items")

    return {"used_llm": True, "valid": not warnings, "warnings": warnings}


def build_docx(
    table1: Dict[str, Any],
    table2: Dict[str, Any],
    table3: Dict[str, Any],
    summary: Dict[str, Any],
    texts: Dict[str, str],
) -> Document:
    doc = Document()
    setup_document(doc)
    doc.add_heading("1.1.2 声环境现状调查与评价", level=1)
    doc.add_heading("1.1.2.1 环境质量调查与评价", level=2)

    doc.add_heading("监测方案", level=3)
    doc.add_paragraph(texts["monitoring_plan_text"])
    add_caption(doc, table1["title"])
    add_table(doc, table1["headers"], table1["rows"])

    doc.add_heading("监测结果", level=3)
    doc.add_paragraph(texts["monitoring_result_text"])
    doc.add_paragraph("（1）敏感点声环境质量现状")
    doc.add_paragraph("监测结果见表4.2-2。")
    add_caption(doc, table2["title"])
    add_result_table(doc, table2["rows"], include_flow=True)
    doc.add_paragraph(texts["sensitive_conclusion"])

    doc.add_paragraph("（2）交通噪声监测结果")
    doc.add_paragraph(texts["traffic_intro"])
    add_caption(doc, table3["title"])
    add_result_table(doc, table3["rows"], include_flow=True)
    doc.add_paragraph(texts["attenuation_conclusion"])
    return doc


def sensitive_conclusion(summary: Dict[str, Any]) -> str:
    if not summary["exceed_items"]:
        return "根据监测结果，项目沿线敏感点的昼间、夜间监测值均满足《声环境质量标准》相应标准限值。"
    return "根据监测结果，" + format_exceed_items(summary["exceed_items"]) + "；其他敏感点的昼间、夜间监测值满足《声环境质量标准》相应标准限值。"


def attenuation_conclusion(summary: Dict[str, Any], lead_only: bool = False) -> str:
    if lead_only:
        return ""
    if not summary["exceed_items"]:
        return "根据断面监测结果，现状公路交通噪声衰减断面昼间、夜间监测值均满足相应标准限值。"
    return "根据衰减断面监测结果，" + format_exceed_items(summary["exceed_items"]) + "。"


def format_exceed_items(items: List[Dict[str, Any]]) -> str:
    return "；".join(
        f"{item['point_name']}（{item['point_code']}）{item['period']}超标{item['exceed']}dB(A)"
        for item in items
    )


def setup_document(doc: Document) -> None:
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.8)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)
    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "宋体"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = RGBColor(0, 0, 0)
    for style_name, size in [("Heading 1", 14), ("Heading 2", 12), ("Heading 3", 11)]:
        style = styles[style_name]
        style.font.name = "黑体"
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "黑体")
        style.font.size = Pt(size)
        style.font.color.rgb = RGBColor(0, 0, 0)


def add_caption(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.bold = True
    run.font.color.rgb = RGBColor(0, 0, 0)


def add_table(doc: Document, headers: List[str], rows: List[Dict[str, Any]]) -> None:
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    for i, header in enumerate(headers):
        table.rows[0].cells[i].text = header
        shade_cell(table.rows[0].cells[i], "EDEDED")
    for row in rows:
        cells = table.add_row().cells
        for i, header in enumerate(headers):
            cells[i].text = str(row.get(header, "-"))
    style_table(table)
    doc.add_paragraph()


def add_result_table(doc: Document, rows: List[Dict[str, Any]], include_flow: bool) -> None:
    headers = result_headers()
    table = doc.add_table(rows=3, cols=len(headers))
    table.style = "Table Grid"
    build_result_table_header(table)
    for row in rows:
        cells = table.add_row().cells
        for i, header in enumerate(headers):
            value = row.get(header, "-")
            cells[i].text = "-" if value == 0 and header.startswith("exceed_") else str(value)
    style_table(table)
    doc.add_paragraph()


def build_result_table_header(table: Any) -> None:
    # Left fixed columns: vertical merge across the three header rows.
    for col, text in [(0, "监测点编号"), (1, "监测点名称"), (2, "监测点位置")]:
        cell = table.cell(0, col).merge(table.cell(2, col))
        cell.text = text

    # Top-level groups.
    table.cell(0, 3).merge(table.cell(0, 6)).text = "LAeq"
    table.cell(0, 7).merge(table.cell(0, 8)).text = "平均值"
    table.cell(0, 9).merge(table.cell(0, 10)).text = "标准值"
    table.cell(0, 11).merge(table.cell(0, 12)).text = "超标量"
    table.cell(0, 13).merge(table.cell(0, 16)).text = "车流量（辆/20min）"

    # Second-level groups.
    table.cell(1, 3).merge(table.cell(1, 4)).text = "第一天"
    table.cell(1, 5).merge(table.cell(1, 6)).text = "第二天"
    table.cell(1, 7).merge(table.cell(1, 8)).text = "平均值"
    table.cell(1, 9).merge(table.cell(1, 10)).text = "标准值"
    table.cell(1, 11).merge(table.cell(1, 12)).text = "超标量"
    table.cell(1, 13).merge(table.cell(1, 14)).text = "第一天"
    table.cell(1, 15).merge(table.cell(1, 16)).text = "第二天"

    for col, text in {
        3: "昼",
        4: "夜",
        5: "昼",
        6: "夜",
        7: "昼",
        8: "夜",
        9: "昼",
        10: "夜",
        11: "昼",
        12: "夜",
        13: "昼",
        14: "夜",
        15: "昼",
        16: "夜",
    }.items():
        table.cell(2, col).text = text

    for row in table.rows[:3]:
        for cell in row.cells:
            shade_cell(cell, "EDEDED")


def style_table(table: Any) -> None:
    for row in table.rows:
        for cell in row.cells:
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_margin(cell)
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.font.name = "宋体"
                    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
                    run.font.size = Pt(8)
                    run.font.color.rgb = RGBColor(0, 0, 0)


def shade_cell(cell: Any, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_cell_margin(cell: Any, margin: int = 70) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for side in ("top", "left", "bottom", "right"):
        node = tc_mar.find(qn(f"w:{side}"))
        if node is None:
            node = OxmlElement(f"w:{side}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(margin))
        node.set(qn("w:type"), "dxa")


def parse_table_rows(table_text: str) -> List[List[str]]:
    rows: List[List[str]] = []
    for line in table_text.splitlines():
        body = re.sub(r"^row\s+\d+:\s*", "", line.strip(), flags=re.IGNORECASE)
        cells = [
            match.group(2).strip()
            for match in re.finditer(r"col\s+(\d+):\s*(.*?)(?=\s*\|\s*col\s+\d+:|$)", body)
        ]
        if any(cells):
            rows.append(cells)
    return rows


def row_to_dict(headers: List[str], row: List[str]) -> Dict[str, Any]:
    return {headers[i]: row[i] if i < len(row) else "" for i in range(len(headers))}


def extract_point_code(text: str) -> Optional[str]:
    match = re.search(r"\bNJ\d+(?:-\d+)?\b", text or "")
    return match.group(0) if match else None


def extract_distance(text: str) -> Optional[str]:
    match = re.search(r"(\d+)\s*m", text or "", flags=re.IGNORECASE)
    return match.group(1) + "m" if match else None


def is_attenuation_text(text: str) -> bool:
    return any(item in str(text or "") for item in ["衰减断面", "距离中心线", "距离公路中心线", "距离宁徐高速公路中心线", "距离淮徐高速公路中心线"])


def infer_period(time_text: str) -> str:
    match = re.search(r"(\d{1,2}):(\d{2})", str(time_text or ""))
    if not match:
        return "unknown"
    hour = int(match.group(1))
    return "day" if 6 <= hour < 22 else "night"


def parse_int(value: Any) -> Optional[int]:
    match = re.search(r"\d+(?:\.\d+)?", str(value or ""))
    if not match:
        return None
    return int(Decimal(match.group(0)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def round_half_up(value: float) -> int:
    return int(Decimal(str(value)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def value_or_dash(value: Any) -> Any:
    return "-" if value is None else value


def exceed_text(avg: Optional[int], limit: Optional[int]) -> Any:
    if avg is None or limit is None:
        return "-"
    return max(0, avg - limit)


def flow_text(pairs: List[Dict[str, Dict[str, Any]]], day_index: int, period: str) -> str:
    if day_index >= len(pairs):
        return "/"
    record = pairs[day_index].get(period)
    if not record:
        return "/"
    return str(record.get("traffic_flow") or "/").replace(" ", "/")


def point_sort_key(code: Any) -> Tuple[int, int, str]:
    text = str(code or "")
    match = re.search(r"NJ(\d+)(?:-(\d+))?", text)
    if not match:
        return (9999, 9999, text)
    return (int(match.group(1)), int(match.group(2) or 0), text)


def write_json(path: Path, payload: Any) -> None:
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
