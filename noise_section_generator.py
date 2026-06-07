import json
import os
import re
import time
from collections import OrderedDict
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

from dotenv import load_dotenv
from docx import Document

from docx_layout import (
    HEADER_FILL,
    TABLE_STYLE,
    add_body_paragraph,
    add_caption,
    add_chapter_title,
    add_landscape_section,
    add_level3_heading,
    add_portrait_section,
    add_section_heading,
    add_table,
    create_section_document,
    finalize_section_document,
    finalize_table,
    setup_document,
    set_cell_text_style,
    shade_cell,
)
from docx_numbering import DocxNumbering, load_numbering
from llm_client import (
    LLM_TEXT_BATCH_PAUSE_SECONDS,
    LlmProfile,
    build_rule_text_fallback_validation,
    chat_completion_json_object_with_recovery,
    is_network_error,
)
from text_polish_utils import build_text_polish_prompt, load_text_polish_guidance
from table_schema_mapper import apply_header_mapping, record_data_validation, resolve_table_schema
from word_processor import load_docx_chunks

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = Path(os.getenv("EIA_INPUT_DIR", "input"))
OUTPUT_DIR = Path(os.getenv("EIA_OUTPUT_DIR", "output"))
DEBUG_DIR = OUTPUT_DIR / "debug_tables"
OUTPUT_DOCX = OUTPUT_DIR / "noise_section.docx"
ENABLE_LLM_TEXT_POLISH = os.getenv("ENABLE_LLM_TEXT_POLISH", "false").lower() == "true"
TEXT_POLISH_GUIDANCE_PATH = Path(
    os.getenv("EIA_NOISE_TEXT_POLISH_GUIDANCE", BASE_DIR / "config" / "noise_text_polish_guidance.json")
)

POINT_KEY = "点位"
TIME_KEY = "监测时间"
FLOW_KEY = "车流量（辆/20min）"
DATE_TUPLE = Tuple[int, int, int]
LAB_NAME_RE = re.compile(
    r"^[\u4e00-\u9fffA-Za-z0-9（）()·\-]{4,80}(?:检测中心|环境监测中心|环境检测中心)$"
)
DATE_TOKEN_RE = re.compile(r"(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})")
DATE_RANGE_RE = re.compile(
    r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})\s*[-~–—至]\s*(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})"
)

DAY_LIMITS = {"0类": 50, "1类": 55, "2类": 60, "3类": 65, "4a类": 70, "4b类": 70}
NIGHT_LIMITS = {"0类": 40, "1类": 45, "2类": 50, "3类": 55, "4a类": 55, "4b类": 60}

NOISE_TABLE_MONITOR_POINTS = "noise_monitor_points"
NOISE_TABLE_SENSITIVE = "noise_sensitive_results"
NOISE_TABLE_INDOOR = "noise_indoor_results"
NOISE_TABLE_ATTENUATION = "noise_attenuation"


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    DEBUG_DIR.mkdir(parents=True, exist_ok=True)

    plan_rows = parse_noise_plan()
    area_records, traffic_records = load_available_flattened_noise()
    all_records = [*area_records, *traffic_records]
    if not all_records:
        raise FileNotFoundError(f"未找到可用的噪声监测结果扁平化表: {DEBUG_DIR}")

    validate_plan_report_match(all_records, plan_rows)
    standard_warnings = [
        f"{row.get('point_code')} 未识别现状标准: {row.get('standard_class') or '空'}"
        for row in plan_rows
        if row.get("standard_class") not in DAY_LIMITS
        or row.get("standard_class") not in NIGHT_LIMITS
    ]
    position_warnings = build_plan_report_position_warnings(all_records, plan_rows)
    record_data_validation(
        OUTPUT_DIR,
        "noise_data",
        valid=not standard_warnings and not position_warnings,
        warnings=[*standard_warnings, *position_warnings],
    )
    table1 = build_monitor_points_table(plan_rows)
    table2 = build_sensitive_result_table(all_records, plan_rows)
    table_indoor = build_indoor_result_table(table2)
    table3 = build_attenuation_result_table(traffic_records, plan_rows)
    summary = build_compliance_summary(table2, table3)

    numbering = load_numbering(OUTPUT_DIR)
    numbering.begin_section("noise", "声环境现状调查与评价")
    register_noise_tables(numbering, table1, table2, table3, table_indoor)

    write_json(DEBUG_DIR / "noise_monitor_points_table.json", table1)
    write_json(DEBUG_DIR / "noise_sensitive_points_result_table.json", table2)
    write_json(DEBUG_DIR / "noise_indoor_result_table.json", table_indoor)
    write_json(DEBUG_DIR / "traffic_noise_attenuation_table.json", table3)
    monitoring_meta = build_noise_monitoring_meta(all_records, plan_rows)
    summary["monitoring_meta"] = monitoring_meta
    write_json(DEBUG_DIR / "noise_compliance_summary.json", summary)
    write_json(DEBUG_DIR / "noise_monitoring_meta.json", monitoring_meta)

    texts = build_rule_texts(table1, table2, table3, summary, plan_rows, numbering)
    texts = polish_noise_text_with_llm(texts, table1, table2, table3, summary, numbering)

    doc = build_docx(table1, table2, table3, table_indoor, summary, texts, numbering)
    finalize_section_document(doc)
    doc.save(OUTPUT_DOCX)
    write_json(DEBUG_DIR / "noise_section_texts.json", texts)
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
    defaults = extract_noise_plan_defaults(chunks)
    for chunk in chunks:
        if chunk.get("kind") != "table":
            continue
        title = (chunk.get("metadata") or {}).get("table_title") or ""
        if "声环境" not in title:
            continue
        rows = parse_table_rows(chunk.get("text") or "")
        if len(rows) < 2:
            continue
        headers = rows[0]
        schema_result = resolve_table_schema(
            "noise_plan",
            headers,
            sample_rows=rows[1:4],
            table_title=title,
            output_dir=OUTPUT_DIR,
            source_table=chunk["chunk_id"],
            job_id=OUTPUT_DIR.parent.name,
        )
        if not schema_result["validation"]["valid"]:
            continue
        parsed = [
            normalize_plan_row(
                {
                    **row_to_dict(headers, row),
                    **apply_header_mapping(headers, row, schema_result["mapping"]),
                },
                source_file=str(plan_path),
                source_table=chunk["chunk_id"],
                defaults=defaults,
            )
            for row in rows[1:]
            if len(row) >= 5
        ]
        parsed = [row for row in parsed if row.get("point_code")]
        if parsed:
            return parsed
    raise RuntimeError("未找到声环境现状监测方案表")


def normalize_plan_row(
    row: Dict[str, Any],
    source_file: str,
    source_table: str,
    defaults: Optional[Dict[str, str]] = None,
) -> Dict[str, Any]:
    defaults = defaults or {}
    code = cell_value(row, "point_code", "监测点编号", "点位编号", "编号")
    return {
        "point_code": code,
        "station": cell_value(row, "station", "桩号"),
        "point_name": cell_value(row, "point_name", "监测点名称", "敏感目标名称", "名称"),
        "district": cell_value(row, "district", "行政区划", "所在行政区"),
        "position": cell_value(row, "position", "监测点位置", "点位", "监测位置"),
        "standard_class": cell_value(
            row,
            "standard_class",
            "现状标准",
            "现状执行标准",
            "声现状执行标准",
            "声环境现状执行标准",
            "声环境标准",
            "标准类别",
            "执行标准",
        ),
        "factor": cell_value(row, "factor", "监测因子") or defaults.get("factor", ""),
        "frequency": cell_value(row, "frequency", "监测频次", "监测时间", "监测要求") or defaults.get("frequency", ""),
        "is_attenuation": is_attenuation_text(" ".join(str(v) for v in row.values())),
        "source_file": source_file,
        "source_table": source_table,
    }


def cell_value(row: Dict[str, Any], *keys: str) -> str:
    for key in keys:
        value = str(row.get(key) or "").strip()
        if value:
            return value
    return ""


def extract_noise_plan_defaults(chunks: List[Dict[str, Any]]) -> Dict[str, str]:
    text = "\n".join(str(chunk.get("text") or "") for chunk in chunks if chunk.get("kind") == "paragraphs")
    defaults: Dict[str, str] = {}
    factor_match = re.search(r"监测因子\s*[\r\n]+([^\r\n]+)", text)
    if factor_match:
        defaults["factor"] = factor_match.group(1).strip()
    frequency_match = re.search(r"监测\s*2\s*天.*?昼间、夜间.*?各监测\s*1\s*次", text, flags=re.DOTALL)
    if frequency_match:
        defaults["frequency"] = "连续监测2天，每天昼间、夜间各监测1次，每次测量时间不低于20min。"
    return defaults


def load_flattened_noise(filename: str, noise_type: str) -> List[Dict[str, Any]]:
    path = DEBUG_DIR / filename
    if not path.exists():
        raise FileNotFoundError(f"缺少噪声扁平化表: {path}")
    data = json.loads(path.read_text(encoding="utf-8"))
    return parse_flattened_noise_records(data, noise_type)


def load_available_flattened_noise() -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    area_records: List[Dict[str, Any]] = []
    traffic_records: List[Dict[str, Any]] = []
    debug_items: List[Dict[str, Any]] = []
    paths = sorted(DEBUG_DIR.glob("flattened_table_*.json"), key=flattened_table_sort_key)
    for path in paths:
        data = json.loads(path.read_text(encoding="utf-8"))
        noise_type = infer_flattened_noise_type(data)
        records = parse_flattened_noise_records(data, noise_type)
        debug_items.append(
            {
                "file": path.name,
                "noise_type": noise_type,
                "record_count": len(records),
                "table_title": data.get("table_title") or "",
                "flattened_headers": data.get("flattened_headers") or [],
            }
        )
        if noise_type == "traffic_noise":
            traffic_records.extend(records)
        else:
            area_records.extend(records)
    write_json(DEBUG_DIR / "noise_flattened_table_detection.json", debug_items)
    return area_records, traffic_records


def flattened_table_sort_key(path: Path) -> Tuple[int, str]:
    match = re.search(r"flattened_table_(\d+)\.json$", path.name)
    return (int(match.group(1)) if match else 9999, path.name)


def infer_flattened_noise_type(data: Dict[str, Any]) -> str:
    text_parts = [
        str(data.get("table_title") or ""),
        " ".join(str(item) for item in data.get("flattened_headers") or []),
        " ".join(
            str(key)
            for row in (data.get("records") or [])[:5]
            for key in row.keys()
        ),
    ]
    text = "\n".join(text_parts)
    traffic_markers = ("交通噪声", "车流量", "衰减断面", "交通量", "距路")
    if any(marker in text for marker in traffic_markers):
        return "traffic_noise"
    return "area_environment_noise"


def parse_flattened_noise_records(data: Dict[str, Any], noise_type: str) -> List[Dict[str, Any]]:
    records: List[Dict[str, Any]] = []
    raw_records = data.get("records") or []
    headers = list(data.get("flattened_headers") or [])
    if not headers and raw_records:
        headers = [str(key) for key in raw_records[0].keys() if not str(key).startswith("_")]
    schema_result = resolve_table_schema(
        "noise_result",
        headers,
        sample_rows=[
            [str(row.get(header) or "") for header in headers]
            for row in raw_records[:3]
        ],
        table_title=str(data.get("table_title") or ""),
        output_dir=OUTPUT_DIR,
        source_table=str(data.get("chunk_id") or ""),
        job_id=OUTPUT_DIR.parent.name,
    )
    mapping = schema_result.get("mapping") or {}
    point_key = mapping.get("point") or POINT_KEY
    time_key = mapping.get("monitor_time") or TIME_KEY
    mapped_laeq_key = mapping.get("laeq")
    flow_key = mapping.get("traffic_flow") or FLOW_KEY
    for index, row in enumerate(raw_records):
        point_text = str(row.get(point_key) or "").strip()
        laeq_key = mapped_laeq_key or next((key for key in row if key.endswith("_LAeq")), None)
        if not point_text or not laeq_key:
            continue
        raw_point_code = extract_point_code(point_text)
        if not raw_point_code and not has_noise_point_identity(point_text):
            continue
        records.append(
            {
                "noise_type": noise_type,
                "point_text": point_text,
                "raw_point_code": raw_point_code,
                "monitor_time": row.get(time_key) or "",
                "period": infer_period(row.get(time_key) or ""),
                "laeq": row.get(laeq_key),
                "traffic_flow": row.get(flow_key) or "/",
                "source_order": index,
                "raw": row,
            }
        )
    return records


def has_noise_point_identity(text: str) -> bool:
    value = str(text or "").strip()
    if len(value) < 4:
        return False
    return any(marker in value for marker in ("点位编号", "敏感点", "背景点", "衰减断面"))


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
    if not any(str(row.get("行政区划") or "").strip() for row in rows):
        headers = [header for header in headers if header != "行政区划"]
        for row in rows:
            row.pop("行政区划", None)
    return {
        "table_key": NOISE_TABLE_MONITOR_POINTS,
        "caption_suffix": "声环境质量现状监测点",
        "title": "声环境质量现状监测点",
        "headers": headers,
        "rows": rows,
    }


def build_sensitive_result_table(
    records: List[Dict[str, Any]],
    plan_rows: List[Dict[str, Any]],
) -> Dict[str, Any]:
    matched = match_records(records, plan_rows)
    grouped = group_sensitive_records_for_result(matched)
    rows: List[Dict[str, Any]] = []
    warnings: List[str] = []
    for plan in sorted(
        [row for row in plan_rows if not row.get("is_attenuation")],
        key=lambda row: point_sort_key(row["point_code"]),
    ):
        code = plan["point_code"]
        if not any((group["plan"].get("source_plan_code") or group["plan"].get("point_code")) == code for group in grouped.values()):
            raise ValueError(f"监测方案点位 {code} 未在噪声监测结果中匹配到数据")
    for code in sorted(grouped, key=point_sort_key):
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
    include_flow = has_traffic_flow(rows)
    if not include_flow:
        strip_traffic_flow_fields(rows)
    headers = result_headers(include_flow=include_flow)
    return {
        "table_key": NOISE_TABLE_SENSITIVE,
        "caption_suffix": "项目沿线敏感点声环境现状监测平均值  单位：dB(A)",
        "title": "项目沿线敏感点声环境现状监测平均值  单位：dB(A)",
        "headers": headers,
        "rows": rows,
        "warnings": warnings,
    }


def build_attenuation_result_table(
    records: List[Dict[str, Any]],
    plan_rows: List[Dict[str, Any]],
) -> Dict[str, Any]:
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
        row["监测点位置"] = attenuation_position_for_record(point_records[0], plan)
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
    include_flow = has_traffic_flow(rows)
    if not include_flow:
        strip_traffic_flow_fields(rows)
    headers = result_headers(include_flow=include_flow)
    return {
        "table_key": NOISE_TABLE_ATTENUATION,
        "caption_suffix": "现状公路交通噪声衰减断面监测结果  单位：dB(A)",
        "title": "现状公路交通噪声衰减断面监测结果  单位：dB(A)",
        "headers": headers,
        "rows": rows,
        "warnings": warnings,
    }


def build_indoor_result_table(table2: Dict[str, Any]) -> Dict[str, Any]:
    indoor_rows: List[Dict[str, Any]] = []
    outdoor_rows: List[Dict[str, Any]] = []
    for row in table2.get("rows") or []:
        if is_indoor_result_row(row):
            indoor_rows.append(row)
        else:
            outdoor_rows.append(row)

    table2["rows"] = outdoor_rows
    include_flow = has_traffic_flow(outdoor_rows)
    if not include_flow:
        strip_traffic_flow_fields(outdoor_rows)
    table2["headers"] = result_headers(include_flow=include_flow)

    include_indoor_flow = has_traffic_flow(indoor_rows)
    if not include_indoor_flow:
        strip_traffic_flow_fields(indoor_rows)
    return {
        "table_key": NOISE_TABLE_INDOOR,
        "caption_suffix": "室内噪声监测结果统计  单位：dB(A)",
        "title": "室内噪声监测结果统计  单位：dB(A)",
        "headers": result_headers(include_flow=include_indoor_flow),
        "rows": indoor_rows,
        "warnings": [
            warning
            for row in indoor_rows
            for warning in (row.get("warnings") or [])
        ],
    }


def is_indoor_result_row(row: Dict[str, Any]) -> bool:
    position = str(row.get("监测点位置") or "")
    return "室内" in position


def result_headers(include_flow: bool = True) -> List[str]:
    headers = [
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
    ]
    if include_flow:
        headers.extend(
            [
                "traffic_flow_day1_day",
                "traffic_flow_day1_night",
                "traffic_flow_day2_day",
                "traffic_flow_day2_night",
            ]
        )
    return headers


def has_traffic_flow(rows: List[Dict[str, Any]]) -> bool:
    flow_headers = [
        "traffic_flow_day1_day",
        "traffic_flow_day1_night",
        "traffic_flow_day2_day",
        "traffic_flow_day2_night",
    ]
    empty_values = {"", "-", "/", "—", "无", "None", "none", "null"}
    for row in rows:
        for header in flow_headers:
            value = str(row.get(header) or "").strip()
            if value not in empty_values:
                return True
    return False


def strip_traffic_flow_fields(rows: List[Dict[str, Any]]) -> None:
    for row in rows:
        for header in (
            "traffic_flow_day1_day",
            "traffic_flow_day1_night",
            "traffic_flow_day2_day",
            "traffic_flow_day2_night",
        ):
            row.pop(header, None)


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


def build_plan_report_position_warnings(
    records: List[Dict[str, Any]],
    plan_rows: List[Dict[str, Any]],
) -> List[str]:
    warnings: List[str] = []
    seen_codes = set()
    for plan in plan_rows:
        code = str(plan.get("point_code") or "").strip()
        if not code or code in seen_codes or plan.get("is_attenuation"):
            continue
        seen_codes.add(code)
        plan_floors = floor_tokens(str(plan.get("position") or ""))
        report_floors = set()
        for record in records:
            if str(record.get("raw_point_code") or "").strip() != code:
                continue
            if "背景点" in str(record.get("point_text") or ""):
                continue
            report_floors.update(floor_tokens(str(record.get("point_text") or "")))
        comparable_plan_floors, comparable_report_floors = normalize_top_floor_equivalence(
            plan_floors,
            report_floors,
        )
        if comparable_plan_floors and comparable_report_floors and comparable_plan_floors != comparable_report_floors:
            warnings.append(
                f"{code} 监测方案楼层({format_floor_tokens(plan_floors)})"
                f"与监测报告楼层({format_floor_tokens(report_floors)})不一致"
            )
    return warnings


def normalize_top_floor_equivalence(plan_floors: set, report_floors: set) -> Tuple[set, set]:
    normalized_plan = set(plan_floors)
    normalized_report = set(report_floors)
    if "顶层" in normalized_report:
        numeric_plan = [item for item in normalized_plan if str(item).isdigit()]
        if numeric_plan:
            normalized_report.remove("顶层")
            normalized_report.add(max(numeric_plan, key=int))
    return normalized_plan, normalized_report


def format_floor_tokens(tokens: set) -> str:
    numeric = sorted((item for item in tokens if str(item).isdigit()), key=int)
    other = sorted(item for item in tokens if not str(item).isdigit())
    return "、".join([*(f"{item}层" for item in numeric), *other])


def match_plan(record: Dict[str, Any], plan_rows: List[Dict[str, Any]]) -> Dict[str, Any]:
    raw_code = record.get("raw_point_code")
    point_text = record.get("point_text") or ""
    background_plan = match_background_point_plan(point_text, plan_rows)
    if background_plan:
        return background_plan
    if raw_code:
        candidates = [row for row in plan_rows if row["point_code"] == raw_code]
        if len(candidates) == 1:
            return candidates[0]
        if len(candidates) > 1:
            by_position = match_plan_by_floor_position(point_text, candidates)
            if by_position:
                return by_position
            if plan_rows_equivalent(candidates):
                return candidates[0]
        by_position = match_plan_by_floor_position(point_text, related_plan_candidates(raw_code, plan_rows))
        if by_position:
            return by_position
        if "-" in raw_code:
            parent_code = raw_code.split("-", 1)[0]
            candidates = [row for row in plan_rows if row["point_code"] == parent_code]
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


def plan_rows_equivalent(rows: List[Dict[str, Any]]) -> bool:
    if not rows:
        return False
    keys = (
        "point_code",
        "station",
        "point_name",
        "district",
        "position",
        "standard_class",
        "factor",
        "frequency",
        "is_attenuation",
    )
    signatures = {
        tuple(str(row.get(key) or "").strip() for key in keys)
        for row in rows
    }
    return len(signatures) == 1


def related_plan_candidates(raw_code: str, plan_rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    if not raw_code:
        return []
    parent_code = raw_code.split("-", 1)[0]
    return [
        row
        for row in plan_rows
        if row.get("point_code") == raw_code
        or str(row.get("point_code") or "").startswith(parent_code + "-")
    ]


def match_plan_by_floor_position(point_text: str, candidates: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    point_floors = floor_tokens(point_text)
    if not point_floors or not candidates:
        return None
    scored: List[Tuple[int, int, Dict[str, Any]]] = []
    for row in candidates:
        row_floors = floor_tokens(str(row.get("position") or ""))
        score = len(point_floors & row_floors)
        if score:
            scored.append((score, -len(row_floors), row))
    if not scored:
        return None
    scored.sort(key=lambda item: (item[0], item[1], point_sort_key(item[2].get("point_code"))), reverse=True)
    return scored[0][2]


def floor_tokens(text: str) -> set:
    value = re.sub(r"\s+", "", str(text or ""))
    tokens = set(re.findall(r"\d+(?=层)", value))
    for match in re.finditer(r"((?:\d+[、,，/和及])+(\d+))层", value):
        tokens.update(re.findall(r"\d+", match.group(1)))
    if "顶层" in value:
        tokens.add("顶层")
    return tokens


def match_background_point_plan(point_text: str, plan_rows: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    text = str(point_text or "")
    if "背景点" not in text:
        return None
    background_candidates = [
        row
        for row in plan_rows
        if any(
            marker in " ".join(str(row.get(key) or "") for key in ("point_name", "position", "station"))
            for marker in ("背景点", "中部", "空旷", "区域环境")
        )
    ]
    if not background_candidates:
        return None
    text_tokens = set(re.findall(r"[\u4e00-\u9fff]{2,}", text))
    scored: List[Tuple[int, Dict[str, Any]]] = []
    for row in background_candidates:
        row_text = " ".join(str(row.get(key) or "") for key in ("point_name", "position", "station"))
        score = sum(1 for token in text_tokens if token and token in row_text)
        if row.get("point_name") and str(row.get("point_name")).replace("小区中部", "") in text:
            score += 5
        scored.append((score, row))
    scored = sorted(scored, key=lambda item: (item[0], point_sort_key(item[1].get("point_code"))), reverse=True)
    return scored[0][1] if scored and scored[0][0] > 0 else None


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


def group_sensitive_records_for_result(records: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    grouped: "OrderedDict[str, Dict[str, Any]]" = OrderedDict()
    for record in records:
        plan = record["plan"]
        raw_code = str(record.get("raw_point_code") or "").strip()
        plan_code = str(plan.get("point_code") or "").strip()
        remapped_background = is_remapped_background_record(record, plan)
        display_code = plan_code if remapped_background else raw_code or plan_code or record.get("point_text")
        group_key = sensitive_result_group_key(record, plan)
        if group_key not in grouped:
            result_plan = dict(plan)
            result_plan["source_plan_code"] = plan_code
            result_plan["point_code"] = display_code
            result_plan["point_name"] = clean_result_point_name(record, plan)
            result_plan["position"] = clean_result_position(
                result_position_for_record(record, plan),
                record,
                plan,
            )
            result_plan["result_sort_key"] = group_key
            grouped[group_key] = {"plan": result_plan, "records": []}
        grouped[group_key]["records"].append(record)
    for value in grouped.values():
        value["records"].sort(key=lambda item: item["source_order"])
    return grouped


def is_remapped_background_record(record: Dict[str, Any], plan: Dict[str, Any]) -> bool:
    point_text = str(record.get("point_text") or "")
    raw_code = str(record.get("raw_point_code") or "").strip()
    plan_code = str(plan.get("point_code") or "").strip()
    return "背景点" in point_text and bool(raw_code and plan_code and raw_code != plan_code)


def sensitive_result_group_key(record: Dict[str, Any], plan: Dict[str, Any]) -> str:
    raw_code = str(record.get("raw_point_code") or plan.get("point_code") or record.get("point_text") or "").strip()
    if is_remapped_background_record(record, plan):
        raw_code = str(plan.get("point_code") or raw_code)
    report_sample_code = extract_report_sample_code(str(record.get("point_text") or ""))
    if report_sample_code:
        return f"{raw_code}|{report_sample_code}"
    position = result_position_for_record(record, plan)
    return f"{raw_code}|{position}"


def extract_report_sample_code(text: str) -> str:
    match = re.search(r"点位编号\s*([A-Za-z]\d+(?:-\d+)?)", str(text or ""), flags=re.IGNORECASE)
    return match.group(1).upper() if match else ""


def result_position_for_record(record: Dict[str, Any], plan: Dict[str, Any]) -> str:
    point_text = str(record.get("point_text") or "")
    raw_code = str(record.get("raw_point_code") or "")
    if is_background_record(record):
        return background_position_for_record(record, plan)
    report_position = extract_position_from_point_text(point_text, record, plan)
    if has_indoor_outdoor_marker(report_position):
        return report_position
    if report_position and has_orientation_marker(report_position):
        return report_position
    if report_position:
        return enrich_floor_position_with_plan_context(report_position, plan)
    plan_positions = split_floor_positions(str(plan.get("position") or ""))
    if raw_code and "-" in raw_code and plan_positions:
        try:
            sub_index = int(raw_code.rsplit("-", 1)[1]) - 1
        except ValueError:
            sub_index = -1
        if 0 <= sub_index < len(plan_positions):
            return plan_positions[sub_index]
    floor_position = extract_floor_position(point_text)
    if floor_position:
        return enrich_floor_position_with_plan_context(floor_position, plan)
    return plan.get("position") or point_text or "-"


def clean_result_point_name(record: Dict[str, Any], plan: Dict[str, Any]) -> str:
    plan_name = sanitize_result_point_name(plan.get("point_name"), plan)
    if plan_name:
        return plan_name
    return sanitize_result_point_name(result_name_for_record(record, plan), plan) or "-"


def sanitize_result_point_name(value: Any, plan: Dict[str, Any]) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    station = str(plan.get("station") or "").strip()
    position = str(plan.get("position") or "").strip()
    if station:
        text = text.replace(station, "")
    if position and position in text:
        text = text.replace(position, "")
    text = re.sub(r"K\s*\d+\s*[+-]\s*\d+", "", text, flags=re.IGNORECASE)
    text = re.sub(r"点位编号\s*[A-Za-z]\d+(?:-\d+)?", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\d+\s*[层楼](?:室外|室内)?(?:监测)?", "", text)
    text = re.sub(r"(?:面向|背向)[^，。；;\r\n]*?(?:首排|第二排|本项目|道路|铁路|陇海线)[^，。；;\r\n]*", "", text)
    text = re.sub(r"(?:敏感点|背景点|衰减断面|室外监测|室内监测|噪声监测)", "", text)
    text = re.sub(r"\s+", " ", text).strip(" 、，,;；")
    return text.strip()


def clean_result_position(value: Any, record: Dict[str, Any], plan: Dict[str, Any]) -> str:
    text = str(value or "").strip()
    if not text:
        return "-"
    point_text = str(record.get("point_text") or "")
    raw_code = str(record.get("raw_point_code") or "").strip()
    station = str(plan.get("station") or "").strip()
    point_name = str(plan.get("point_name") or "").strip()
    if text == point_text:
        text = strip_point_identity(text, raw_code, station, point_name)
    else:
        text = strip_point_identity(text, raw_code, station, point_name)
    text = re.sub(r"点位编号\s*[A-Za-z]*\d+(?:-\d+)?", "", text, flags=re.IGNORECASE)
    text = normalize_position_text(text)
    text = re.sub(r"\s+", "", text).strip("、，,;；")
    if is_background_record(record):
        return append_background_marker(extract_floor_position_fragment(text) or text)
    if not has_orientation_marker(text):
        text = extract_floor_position_fragment(text) or text
    return text or "-"


def extract_position_from_point_text(text: str, record: Dict[str, Any], plan: Dict[str, Any]) -> str:
    cleaned = strip_point_identity(
        str(text or ""),
        str(record.get("raw_point_code") or "").strip(),
        str(plan.get("station") or "").strip(),
        str(plan.get("point_name") or "").strip(),
    )
    cleaned = re.sub(r"点位编号\s*[A-Za-z]*\d+(?:-\d+)?", "", cleaned, flags=re.IGNORECASE)
    cleaned = normalize_position_text(cleaned)
    cleaned = re.sub(r"\s+", "", cleaned).strip("、，,;；")
    if re.search(r"(面向|背向|首排|第二排|[0-9]+[层楼]|顶层|室外|室内)", cleaned):
        if not has_orientation_marker(cleaned):
            return extract_floor_position_fragment(cleaned) or cleaned
        return cleaned
    return ""


def normalize_position_text(text: str) -> str:
    value = str(text or "")
    value = re.sub(r"(室外|室内)\s*监测", r"\1", value)
    value = re.sub(r"(室外|室内)\s*噪声\s*监测", r"\1", value)
    value = re.sub(r"(?:噪声监测|监测|进行)", "", value)
    value = re.sub(r"敏感点", "", value)
    return value


def is_background_record(record: Dict[str, Any]) -> bool:
    return "背景点" in str(record.get("point_text") or "")


def background_position_for_record(record: Dict[str, Any], plan: Dict[str, Any]) -> str:
    point_text = str(record.get("point_text") or "")
    report_floor = extract_floor_position(point_text)
    if report_floor:
        return append_background_marker(report_floor)
    plan_floor = first_floor_position_from_plan(str(plan.get("position") or ""))
    if plan_floor:
        return append_background_marker(plan_floor)
    return "背景点"


def append_background_marker(position: str) -> str:
    value = str(position or "").strip()
    if not value:
        return "背景点"
    if "背景点" in value:
        return value
    return f"{value}背景点"


def first_floor_position_from_plan(text: str) -> str:
    value = str(text or "")
    multi_floor = re.search(r"((?:\d+\s*[、,，和及]\s*)+\d+)\s*[层楼]", value)
    if multi_floor:
        first = re.search(r"\d+", multi_floor.group(1))
        if first:
            suffix = "室外" if "室外" in value else "室内" if "室内" in value else ""
            return f"{first.group(0)}层{suffix}"
    match = re.search(r"(\d+\s*[层楼](?:室外|室内)?|顶层(?:室外|室内)?)", value)
    if match:
        return re.sub(r"\s+", "", match.group(1))
    if "室外" in value:
        return "室外"
    if "室内" in value:
        return "室内"
    return ""


def extract_floor_position_fragment(text: str) -> str:
    value = str(text or "")
    matches = re.findall(r"\d+\s*[层楼](?:室外|室内)?|顶层(?:室外|室内)?", value)
    if matches:
        return re.sub(r"\s+", "", matches[-1])
    if "室外" in value:
        return "室外"
    if "室内" in value:
        return "室内"
    return ""


def has_indoor_outdoor_marker(text: str) -> bool:
    return "室内" in str(text or "") or "室外" in str(text or "")


def has_orientation_marker(text: str) -> bool:
    return bool(re.search(r"(面向|背向|首排|第二排|本项目|道路|铁路|陇海线)", str(text or "")))


def enrich_floor_position_with_plan_context(position: str, plan: Dict[str, Any]) -> str:
    value = str(position or "").strip()
    plan_position = str(plan.get("position") or "").strip()
    if not value or has_indoor_outdoor_marker(value) or has_orientation_marker(value):
        return value
    if re.fullmatch(r"\d+\s*[层楼]|顶层", value) and has_orientation_marker(plan_position):
        context = strip_plan_floor_terms(plan_position)
        if context:
            return f"{context}{value}"
    return value


def strip_plan_floor_terms(text: str) -> str:
    value = str(text or "")
    value = re.sub(r"\d+(?:[、,，和及]\d+)*\s*[层楼](?:室外|室内)?(?:监测)?", "", value)
    value = normalize_position_text(value)
    value = re.sub(r"\s+", "", value).strip("、，,;；")
    return value


def strip_point_identity(text: str, raw_code: str, station: str, point_name: str) -> str:
    result = str(text or "")
    if raw_code:
        result = re.sub(rf"\b{re.escape(raw_code)}\b", "", result, flags=re.IGNORECASE)
    if station:
        result = result.replace(station, "")
    result = re.sub(r"K\s*\d+\s*[+-]\s*\d+", "", result, flags=re.IGNORECASE)
    if point_name:
        result = result.replace(point_name, "")
        base_name = re.sub(r"\d+$", "", point_name).strip()
        if len(base_name) >= 2:
            result = result.replace(base_name, "")
    return result


def attenuation_position_for_record(record: Dict[str, Any], plan: Dict[str, Any]) -> str:
    point_text = str(record.get("point_text") or "")
    match = re.search(
        r"(距离[^，。；;\r\n]*?\d+\s*m\s*处(?:空地)?)",
        point_text,
        flags=re.IGNORECASE,
    )
    if match:
        return re.sub(r"\s+", "", match.group(1))
    return plan.get("position") or point_text or "-"


def result_name_for_record(record: Dict[str, Any], plan: Dict[str, Any]) -> str:
    point_text = str(record.get("point_text") or "")
    raw_code = str(record.get("raw_point_code") or "")
    text = point_text
    if raw_code and text.startswith(raw_code):
        text = text[len(raw_code):].strip()
    text = re.sub(r"点位编号\s*[A-Za-z]\d+(?:-\d+)?", "", text, flags=re.IGNORECASE).strip()
    text = re.sub(r"(敏感点|背景点).*", "", text).strip()
    text = re.sub(r"\d+\s*层.*", "", text).strip()
    return text or plan.get("point_name") or "-"


def extract_floor_position(text: str) -> str:
    text = str(text or "")
    match = re.search(r"(\d+\s*层(?:进行)?(?:室外|室内)?(?:噪声监测)?)", text)
    if match:
        value = re.sub(r"\s+", "", match.group(1))
        return value.replace("进行", "").replace("噪声监测", "")
    if "顶层" in text:
        return "顶层"
    return ""


def split_floor_positions(text: str) -> List[str]:
    text = re.sub(r"\s+", "", str(text or ""))
    if not text or "层" not in text:
        return []
    suffix = "室外" if "室外" in text else "室内" if "室内" in text else ""
    prefix = text.split("层", 1)[0]
    numbers = re.findall(r"\d+", prefix)
    if len(numbers) > 1:
        return [f"{number}层{suffix}" for number in numbers]
    return [extract_floor_position(text)] if extract_floor_position(text) else []


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
    exceed_row_count = 0
    for row in rows:
        row_exceeded = False
        for period, field in [("昼间", "exceed_day"), ("夜间", "exceed_night")]:
            value = row.get(field)
            if isinstance(value, int) and value > 0:
                row_exceeded = True
                exceed_items.append(
                    {
                        "point_code": row.get("监测点编号"),
                        "point_name": row.get("监测点名称"),
                        "period": period,
                        "exceed": value,
                    }
                )
        if row_exceeded:
            exceed_row_count += 1
    return {
        "total_count": len(rows),
        "exceed_count": exceed_row_count,
        "exceed_items": exceed_items,
    }


def register_noise_tables(
    numbering: DocxNumbering,
    table1: Dict[str, Any],
    table2: Dict[str, Any],
    table3: Dict[str, Any],
    table_indoor: Optional[Dict[str, Any]] = None,
) -> None:
    tables = [table1, table2]
    if table_indoor and table_indoor.get("rows"):
        tables.append(table_indoor)
    if table3.get("rows"):
        tables.append(table3)
    for table in tables:
        table["title"] = numbering.register_table(table["table_key"], table["caption_suffix"])


def build_rule_texts(
    table1: Dict[str, Any],
    table2: Dict[str, Any],
    table3: Dict[str, Any],
    summary: Dict[str, Any],
    plan_rows: List[Dict[str, Any]],
    numbering: DocxNumbering,
) -> Dict[str, str]:
    monitoring_meta = summary.get("monitoring_meta") or {}
    return {
        "monitoring_plan_text": build_monitoring_plan_text(table1, plan_rows, numbering),
        "monitoring_result_text": build_monitoring_result_text(monitoring_meta),
        "sensitive_conclusion": sensitive_conclusion(summary["sensitive"]),
        "traffic_intro": build_traffic_intro_text(plan_rows, monitoring_meta, numbering),
        "attenuation_conclusion": attenuation_conclusion(summary["attenuation"], lead_only=False),
    }


def build_monitoring_plan_text(
    table1: Dict[str, Any],
    plan_rows: List[Dict[str, Any]],
    numbering: DocxNumbering,
) -> str:
    total = len(table1["rows"])
    attenuation_count = sum(1 for row in plan_rows if row.get("is_attenuation"))
    text = (
        "根据《环境影响评价技术导则 公路建设项目》（HJ1358-2024）要求，"
        "一级评价应对评价范围内具有代表性的敏感目标的声环境质量进行实测，并对实测结果进行评价。"
        "本项目根据不同路段，贯彻“以点代线、点线结合、以代表性区段为主、反馈全线”的原则，"
        "结合项目沿线敏感点分布、现状噪声源类型及空间特征，"
        f"共设置{total}个声环境监测点"
    )
    if attenuation_count:
        text += f"（含{attenuation_count}个交通噪声衰减断面监测点）"
    text += f"。声环境现状监测方案见{numbering.table_label(NOISE_TABLE_MONITOR_POINTS)}。"
    return text


def build_monitoring_result_text(monitoring_meta: Dict[str, Any]) -> str:
    unit = str(monitoring_meta.get("monitoring_unit") or "").strip()
    date_text = str(monitoring_meta.get("monitor_date_text") or "").strip()
    lead = ""
    if unit and date_text:
        lead = f"{unit}于{date_text}对本项目沿线各监测点位的环境噪声进行了监测。"
    elif unit:
        lead = f"{unit}对本项目沿线各监测点位的环境噪声进行了监测。"
    elif date_text:
        lead = f"本项目于{date_text}对沿线各监测点位开展环境噪声现状监测。"
    tail = (
        "具体测量时间段、测量仪器、测量方法均按规范要求执行。"
        "测量结果以等效连续A声级和统计噪声级给出，并以等效A声级作为最终评价量。"
    )
    return f"{lead}{tail}" if lead else tail


def build_traffic_intro_text(
    plan_rows: List[Dict[str, Any]],
    monitoring_meta: Dict[str, Any],
    numbering: DocxNumbering,
) -> str:
    label = numbering.table_label(NOISE_TABLE_ATTENUATION)
    road_hint = str(monitoring_meta.get("road_hint") or "").strip()
    if road_hint:
        return f"本次评价在{road_hint}开展交通噪声衰减断面监测，监测结果见{label}。"
    return f"本次评价对现状公路交通噪声衰减断面进行监测，监测结果见{label}。"


def build_noise_monitoring_meta(
    records: List[Dict[str, Any]],
    plan_rows: List[Dict[str, Any]],
) -> Dict[str, Any]:
    report_path = find_report_path()
    chunks = load_docx_chunks(report_path)
    monitoring_unit = extract_monitoring_unit(chunks)
    report_ranges = extract_report_noise_date_ranges(chunks)
    record_dates = collect_monitor_dates_from_records(records)
    record_ranges = date_tuples_to_ranges(record_dates)
    merged_ranges = merge_date_ranges([*report_ranges, *record_ranges])
    return {
        "monitoring_unit": monitoring_unit,
        "monitor_date_ranges": [
            {"start": format_date_cn(*start), "end": format_date_cn(*end)}
            for start, end in merged_ranges
        ],
        "monitor_date_text": format_date_ranges_text(merged_ranges),
        "monitor_dates": [format_date_cn(*item) for item in record_dates],
        "road_hint": infer_road_hint(plan_rows),
        "source_file": str(report_path),
    }


def find_report_path() -> Path:
    return next(
        path for path in INPUT_DIR.glob("*.docx")
        if "报告" in path.name and not path.name.startswith("~$")
    )


def extract_monitoring_unit(chunks: List[Dict[str, Any]]) -> Optional[str]:
    for chunk in chunks:
        if chunk.get("kind") != "paragraphs":
            continue
        for line in str(chunk.get("text") or "").splitlines():
            candidate = line.strip()
            if LAB_NAME_RE.match(candidate):
                return candidate
    for chunk in chunks:
        if chunk.get("kind") != "table":
            continue
        text = str(chunk.get("text") or "")
        if "样品类型" not in text or "噪声" not in text:
            continue
        match = re.search(r"检测中心[\u4e00-\u9fffA-Za-z0-9（）()·\-]{0,40}", text)
        if match:
            return match.group(0)
    return None


def extract_report_noise_date_ranges(chunks: List[Dict[str, Any]]) -> List[Tuple[DATE_TUPLE, DATE_TUPLE]]:
    ranges: List[Tuple[DATE_TUPLE, DATE_TUPLE]] = []
    for chunk in chunks:
        if chunk.get("kind") != "table":
            continue
        text = str(chunk.get("text") or "")
        if "样品类型" not in text or "噪声" not in text:
            continue
        for match in DATE_RANGE_RE.finditer(text):
            start = (int(match.group(1)), int(match.group(2)), int(match.group(3)))
            end = (int(match.group(4)), int(match.group(5)), int(match.group(6)))
            ranges.append((start, end))
    return ranges


def collect_monitor_dates_from_records(records: List[Dict[str, Any]]) -> List[DATE_TUPLE]:
    dates: List[DATE_TUPLE] = []
    for record in records:
        parsed = parse_monitor_date(str(record.get("monitor_time") or ""))
        if parsed:
            dates.append(parsed)
    return sorted(set(dates))


def parse_monitor_date(text: str) -> Optional[DATE_TUPLE]:
    match = DATE_TOKEN_RE.search(str(text or ""))
    if not match:
        return None
    return int(match.group(1)), int(match.group(2)), int(match.group(3))


def date_tuples_to_ranges(dates: List[DATE_TUPLE]) -> List[Tuple[DATE_TUPLE, DATE_TUPLE]]:
    if not dates:
        return []
    ranges: List[Tuple[DATE_TUPLE, DATE_TUPLE]] = []
    start = end = dates[0]
    for current in dates[1:]:
        if date_ordinal(current) - date_ordinal(end) <= 1:
            end = current
            continue
        ranges.append((start, end))
        start = end = current
    ranges.append((start, end))
    return ranges


def merge_date_ranges(ranges: List[Tuple[DATE_TUPLE, DATE_TUPLE]]) -> List[Tuple[DATE_TUPLE, DATE_TUPLE]]:
    normalized = sorted(
        ((min(start, end), max(start, end)) for start, end in ranges),
        key=lambda item: date_ordinal(item[0]),
    )
    if not normalized:
        return []
    merged: List[Tuple[DATE_TUPLE, DATE_TUPLE]] = [normalized[0]]
    for start, end in normalized[1:]:
        last_start, last_end = merged[-1]
        if date_ordinal(start) <= date_ordinal(last_end) + 1:
            merged[-1] = (last_start, max([last_end, end], key=date_ordinal))
        else:
            merged.append((start, end))
    return merged


def format_date_cn(year: int, month: int, day: int) -> str:
    return f"{year}年{month}月{day}日"


def format_date_ranges_text(ranges: List[Tuple[DATE_TUPLE, DATE_TUPLE]]) -> str:
    parts: List[str] = []
    for start, end in ranges:
        if start == end:
            parts.append(format_date_cn(*start))
        else:
            parts.append(f"{format_date_cn(*start)}~{format_date_cn(*end)}")
    return "、".join(parts)


def date_ordinal(value: DATE_TUPLE) -> int:
    return date(value[0], value[1], value[2]).toordinal()


def infer_road_hint(plan_rows: List[Dict[str, Any]]) -> str:
    for row in plan_rows:
        if not row.get("is_attenuation"):
            continue
        text = " ".join(
            str(row.get(key) or "")
            for key in ("point_name", "position", "station", "point_code")
        )
        match = re.search(r"([\u4e00-\u9fff]{2,8}高速)", text)
        if match:
            return f"现状{match.group(1)}空旷地段"
    return ""


NOISE_POLISH_FIELD_BATCHES = (
    ("monitoring_plan_text", "monitoring_result_text"),
    ("sensitive_conclusion", "traffic_intro", "attenuation_conclusion"),
)


def _noise_polish_extra_rules(table_labels: Dict[str, str]) -> str:
    return (
        f"必须保留表号 {', '.join(table_labels.values())}。\n"
        "必须保留 summary 中所有超标点位编号和超标量。\n"
        "monitoring_result_text 必须保留 monitoring_meta 中的监测单位和监测日期，不得删改或编造。\n"
        "monitoring_plan_text 不得明显短于 style_exemplars 示例长度。\n"
        "若存在超标，不得写“均满足”或“均达标”。"
    )


def _call_noise_polish_batch(
    payload: Dict[str, Any],
    batch_keys: Tuple[str, ...],
    table_labels: Dict[str, str],
) -> Dict[str, str]:
    batch_rule_texts = {
        key: str(payload["rule_texts"].get(key, ""))
        for key in batch_keys
    }
    batch_payload = {**payload, "rule_texts": batch_rule_texts}
    polished = chat_completion_json_object_with_recovery(
        [
            {
                "role": "system",
                "content": "你是严谨的环评报告文本润色助手。必须输出严格 JSON，禁止编造事实。",
            },
            {
                "role": "user",
                "content": build_text_polish_prompt(
                    batch_payload,
                    role="声环境章节",
                    output_keys=", ".join(batch_keys),
                    extra_rules=_noise_polish_extra_rules(table_labels),
                ),
            },
        ],
        profile=LlmProfile.text_polish,
        label=f"noise_text_polish_{batch_keys[0]}",
    )
    return {
        key: str(polished.get(key, "")).strip() or batch_rule_texts[key]
        for key in batch_keys
    }


def polish_noise_text_with_llm(
    rule_texts: Dict[str, str],
    table1: Dict[str, Any],
    table2: Dict[str, Any],
    table3: Dict[str, Any],
    summary: Dict[str, Any],
    numbering: DocxNumbering,
) -> Dict[str, str]:
    table_labels = {
        table["table_key"]: numbering.table_label(table["table_key"])
        for table in (table1, table2, table3)
    }
    payload = {
        "enabled": ENABLE_LLM_TEXT_POLISH,
        "text_guidance": load_text_polish_guidance(TEXT_POLISH_GUIDANCE_PATH),
        "rule_texts": rule_texts,
        "summary": summary,
        "table_labels": table_labels,
        "monitoring_meta": summary.get("monitoring_meta") or {},
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
        polished_fields = dict(rule_texts)
        extra_rules = _noise_polish_extra_rules(table_labels)
        for batch_index, batch_keys in enumerate(NOISE_POLISH_FIELD_BATCHES):
            batch_result = _call_noise_polish_batch(payload, batch_keys, table_labels)
            polished_fields.update(batch_result)
            if batch_index + 1 < len(NOISE_POLISH_FIELD_BATCHES) and LLM_TEXT_BATCH_PAUSE_SECONDS > 0:
                time.sleep(LLM_TEXT_BATCH_PAUSE_SECONDS)
        merged = ensure_noise_table_refs(polished_fields, rule_texts, numbering)
        validation = validate_polished_text(merged, summary, numbering, rule_texts)
        write_json(DEBUG_DIR / "noise_llm_text_output.json", merged)
        write_json(DEBUG_DIR / "noise_llm_text_validation.json", validation)
        if validation["valid"]:
            return clean_polished_texts(merged, numbering)
        return rule_texts
    except Exception as exc:
        if is_network_error(exc):
            try:
                print("noise_text_polish retrying as single request after batch failure", flush=True)
                time.sleep(LLM_TEXT_BATCH_PAUSE_SECONDS)
                polished = chat_completion_json_object_with_recovery(
                    [
                        {
                            "role": "system",
                            "content": "你是严谨的环评报告文本润色助手。必须输出严格 JSON，禁止编造事实。",
                        },
                        {
                            "role": "user",
                            "content": build_text_polish_prompt(
                                payload,
                                role="声环境章节",
                                output_keys="monitoring_plan_text, monitoring_result_text, sensitive_conclusion, traffic_intro, attenuation_conclusion",
                                extra_rules=extra_rules,
                            ),
                        },
                    ],
                    profile=LlmProfile.text_polish,
                    label="noise_text_polish_single",
                    recovery_attempts=2,
                )
                polished_fields = {key: str(polished.get(key, "")).strip() for key in rule_texts}
                merged = ensure_noise_table_refs({**rule_texts, **polished_fields}, rule_texts, numbering)
                validation = validate_polished_text(merged, summary, numbering, rule_texts)
                write_json(DEBUG_DIR / "noise_llm_text_output.json", merged)
                write_json(DEBUG_DIR / "noise_llm_text_validation.json", validation)
                if validation["valid"]:
                    return clean_polished_texts(merged, numbering)
            except Exception:
                pass
        write_json(DEBUG_DIR / "noise_llm_text_output.json", {})
        write_json(
            DEBUG_DIR / "noise_llm_text_validation.json",
            build_rule_text_fallback_validation(exc),
        )
        return rule_texts


def clean_polished_texts(texts: Dict[str, str], numbering: DocxNumbering) -> Dict[str, str]:
    cleaned = dict(texts)
    sensitive_label = numbering.table_label(NOISE_TABLE_SENSITIVE)
    for prefix in (f"监测结果见{sensitive_label}。", f"监测结果见{sensitive_label}，"):
        value = cleaned.get("sensitive_conclusion", "")
        if value.startswith(prefix):
            cleaned["sensitive_conclusion"] = value[len(prefix):].lstrip()
    return cleaned


def ensure_noise_table_refs(
    texts: Dict[str, str],
    rule_texts: Dict[str, str],
    numbering: DocxNumbering,
) -> Dict[str, str]:
    result = dict(texts)
    field_table_keys = {
        "monitoring_plan_text": NOISE_TABLE_MONITOR_POINTS,
        "traffic_intro": NOISE_TABLE_ATTENUATION,
    }
    for field, table_key in field_table_keys.items():
        label = numbering.table_label(table_key)
        current = str(result.get(field, "")).strip()
        if label in current:
            continue
        rule = str(rule_texts.get(field, "")).strip()
        if label not in rule:
            continue
        match = re.search(rf"[^。]*{re.escape(label)}[^。]*。?", rule)
        if match and current:
            clause = match.group(0).strip()
            if not current.endswith("。"):
                current += "。"
            result[field] = current + clause
        else:
            result[field] = rule
    return result


def validate_polished_text(
    polished: Dict[str, Any],
    summary: Dict[str, Any],
    numbering: DocxNumbering,
    rule_texts: Dict[str, str],
) -> Dict[str, Any]:
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

    all_text = "\n".join(str(polished.get(key, "")) for key in required)
    table_field_checks = {
        NOISE_TABLE_MONITOR_POINTS: "monitoring_plan_text",
        NOISE_TABLE_ATTENUATION: "traffic_intro",
    }
    for table_key, field in table_field_checks.items():
        label = numbering.table_label(table_key)
        field_text = str(polished.get(field, ""))
        rule_text = str(rule_texts.get(field, ""))
        if label in rule_text and label not in field_text:
            warnings.append(f"missing table number: {label}")

    for item in required_exceed_checks(summary):
        point_code = str(item.get("point_code") or "")
        exceed_value = item.get("exceed")
        exceed_text = f"{exceed_value}dB(A)"
        if point_code and point_code not in all_text:
            warnings.append(f"missing exceed point: {point_code}")
        if not contains_exceed_amount(all_text, exceed_value):
            warnings.append(f"missing exceed amount: {exceed_text}")

    for key in ("sensitive", "attenuation"):
        if summary.get(key, {}).get("exceed_items"):
            target_field = "sensitive_conclusion" if key == "sensitive" else "attenuation_conclusion"
            text = str(polished.get(target_field, ""))
            if "均满足" in text or "均达标" in text:
                warnings.append(f"{target_field} says all compliant despite exceed items")

    monitoring_meta = summary.get("monitoring_meta") or {}
    result_text = str(polished.get("monitoring_result_text", ""))
    unit = str(monitoring_meta.get("monitoring_unit") or "").strip()
    if unit and unit not in result_text:
        warnings.append("monitoring_result_text missing monitoring unit from monitoring_meta")
    date_text = str(monitoring_meta.get("monitor_date_text") or "").strip()
    if date_text:
        years = sorted({str(item[0]) for item in re.findall(r"(\d{4})年", date_text)})
        if years and not any(year in result_text for year in years):
            warnings.append("monitoring_result_text missing monitoring dates from monitoring_meta")
    plan_text = str(polished.get("monitoring_plan_text", ""))
    if len(plan_text) < 150:
        warnings.append("monitoring_plan_text shorter than expected reference length")

    return {"used_llm": True, "valid": not warnings, "llm_applied": True, "warnings": warnings}


def required_exceed_checks(summary: Dict[str, Any]) -> List[Dict[str, Any]]:
    grouped: Dict[Tuple[str, str, str], Dict[str, Any]] = {}
    for section in ("sensitive", "attenuation"):
        for item in summary.get(section, {}).get("exceed_items", []):
            point_code = str(item.get("point_code") or "")
            point_name = str(item.get("point_name") or "")
            period = str(item.get("period") or "")
            key = (point_code, point_name, period)
            try:
                exceed = int(item.get("exceed"))
            except (TypeError, ValueError):
                exceed = 0
            current = grouped.get(key)
            if current is None or exceed > int(current.get("exceed") or 0):
                grouped[key] = {**item, "exceed": exceed}
    return list(grouped.values())


def contains_exceed_amount(text: str, exceed: Any) -> bool:
    try:
        value = int(exceed)
    except (TypeError, ValueError):
        return False
    normalized = (
        str(text or "")
        .replace(" ", "")
        .replace("（", "(")
        .replace("）", ")")
        .replace("－", "-")
        .replace("—", "-")
        .replace("~", "～")
    )
    exact = rf"{value}(?:\.0)?dB\(A\)"
    if re.search(exact, normalized, flags=re.IGNORECASE):
        return True
    range_pattern = r"(\d+(?:\.0)?)[～\-至到](\d+(?:\.0)?)dB\(A\)"
    for match in re.finditer(range_pattern, normalized, flags=re.IGNORECASE):
        start = int(float(match.group(1)))
        end = int(float(match.group(2)))
        if min(start, end) <= value <= max(start, end):
            return True
    return False


def build_docx(
    table1: Dict[str, Any],
    table2: Dict[str, Any],
    table3: Dict[str, Any],
    table_indoor: Dict[str, Any],
    summary: Dict[str, Any],
    texts: Dict[str, str],
    numbering: DocxNumbering,
) -> Document:
    doc = create_section_document()
    setup_document(doc)
    add_chapter_title(doc, numbering.section_title("noise", "声环境现状调查与评价"))
    add_section_heading(doc, numbering.next_level2_heading("noise", "监测方案"))
    add_body_paragraph(doc, texts["monitoring_plan_text"])
    add_caption(doc, numbering.table_caption(NOISE_TABLE_MONITOR_POINTS))
    add_monitor_points_table(doc, table1["headers"], table1["rows"])

    add_section_heading(doc, numbering.next_level2_heading("noise", "监测结果"))
    add_body_paragraph(doc, texts["monitoring_result_text"])
    add_level3_heading(doc, numbering.next_level3_heading("noise", "敏感点声环境质量现状"))
    sensitive_label = numbering.table_label(NOISE_TABLE_SENSITIVE)
    add_body_paragraph(doc, f"监测结果见{sensitive_label}。")
    add_landscape_section(doc)
    add_caption(doc, numbering.table_caption(NOISE_TABLE_SENSITIVE))
    add_result_table(doc, table2["rows"], include_flow=True)
    add_body_paragraph(doc, texts["sensitive_conclusion"])

    if table_indoor.get("rows"):
        indoor_label = numbering.table_label(NOISE_TABLE_INDOOR)
        add_body_paragraph(doc, f"室内噪声监测结果统计见{indoor_label}。")
        add_caption(doc, numbering.table_caption(NOISE_TABLE_INDOOR))
        add_result_table(doc, table_indoor["rows"], include_flow=True)

    if table3.get("rows"):
        add_level3_heading(doc, numbering.next_level3_heading("noise", "交通噪声监测结果"))
        add_body_paragraph(doc, texts["traffic_intro"])
        add_caption(doc, numbering.table_caption(NOISE_TABLE_ATTENUATION))
        add_result_table(doc, table3["rows"], include_flow=True)
        add_body_paragraph(doc, texts["attenuation_conclusion"])
    add_portrait_section(doc)
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


def add_result_table(doc: Document, rows: List[Dict[str, Any]], include_flow: bool) -> None:
    include_flow = include_flow and has_traffic_flow(rows)
    headers = result_headers(include_flow=include_flow)
    table = doc.add_table(rows=3, cols=len(headers))
    table.style = TABLE_STYLE
    build_result_table_header(table, include_flow=include_flow)
    for row in rows:
        cells = table.add_row().cells
        for i, header in enumerate(headers):
            value = row.get(header, "-")
            cells[i].text = "-" if value == 0 and header.startswith("exceed_") else str(value)
    merge_result_identity_cells(table, headers, rows)
    finalize_table(table, header_row_count=3)
    doc.add_paragraph()


def merge_result_identity_cells(table: Any, headers: List[str], rows: List[Dict[str, Any]]) -> None:
    if len(rows) <= 1:
        return
    code_header = "监测点编号"
    name_header = "监测点名称"
    if code_header not in headers or name_header not in headers:
        return
    code_col = headers.index(code_header)
    name_col = headers.index(name_header)
    data_start = 3
    start = 0
    while start < len(rows):
        code = str(rows[start].get(code_header) or "").strip()
        name = str(rows[start].get(name_header) or "").strip()
        end = start
        while (
            end + 1 < len(rows)
            and str(rows[end + 1].get(code_header) or "").strip() == code
            and str(rows[end + 1].get(name_header) or "").strip() == name
        ):
            end += 1
        if end > start and code and name:
            code_cell = table.cell(data_start + start, code_col).merge(table.cell(data_start + end, code_col))
            code_cell.text = code
            name_cell = table.cell(data_start + start, name_col).merge(table.cell(data_start + end, name_col))
            name_cell.text = name
        start = end + 1


def add_monitor_points_table(doc: Document, headers: List[str], rows: List[Dict[str, Any]]) -> None:
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = TABLE_STYLE
    table.autofit = True
    header_cells = table.rows[0].cells
    for index, header in enumerate(headers):
        header_cells[index].text = str(header)
        if HEADER_FILL:
            shade_cell(header_cells[index], HEADER_FILL)
        set_cell_text_style(header_cells[index], bold=True)

    for row in rows:
        cells = table.add_row().cells
        for index, header in enumerate(headers):
            cells[index].text = str(row.get(header, "-"))
            set_cell_text_style(cells[index], bold=False)

    merge_repeated_data_columns(table, headers, rows, ["监测因子", "监测频次"])
    finalize_table(table, header_row_count=1)
    doc.add_paragraph()


def merge_repeated_data_columns(
    table: Any,
    headers: List[str],
    rows: List[Dict[str, Any]],
    merge_headers: List[str],
) -> None:
    if len(rows) <= 1:
        return
    for header in merge_headers:
        if header not in headers:
            continue
        values = [str(row.get(header, "")).strip() for row in rows]
        if not values or not values[0] or any(value != values[0] for value in values):
            continue
        column_index = headers.index(header)
        merged_cell = table.cell(1, column_index).merge(table.cell(len(rows), column_index))
        merged_cell.text = values[0]


def build_result_table_header(table: Any, include_flow: bool = True) -> None:
    # Left fixed columns: vertical merge across the three header rows.
    for col, text in [(0, "监测点编号"), (1, "监测点名称"), (2, "监测点位置")]:
        cell = table.cell(0, col).merge(table.cell(2, col))
        cell.text = text

    # Top-level groups.
    table.cell(0, 3).merge(table.cell(0, 6)).text = "LAeq"
    table.cell(0, 7).merge(table.cell(1, 8)).text = "平均值"
    table.cell(0, 9).merge(table.cell(1, 10)).text = "标准值"
    table.cell(0, 11).merge(table.cell(1, 12)).text = "超标量"

    # Second-level groups.
    table.cell(1, 3).merge(table.cell(1, 4)).text = "第一天"
    table.cell(1, 5).merge(table.cell(1, 6)).text = "第二天"

    day_night_labels = {
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
    }
    if include_flow:
        table.cell(0, 13).merge(table.cell(0, 16)).text = "车流量（辆/20min）"
        table.cell(1, 13).merge(table.cell(1, 14)).text = "第一天"
        table.cell(1, 15).merge(table.cell(1, 16)).text = "第二天"
        day_night_labels.update(
            {
                13: "昼",
                14: "夜",
                15: "昼",
                16: "夜",
            }
        )

    for col, text in day_night_labels.items():
        table.cell(2, col).text = text

    for row in table.rows[:3]:
        for cell in row.cells:
            if HEADER_FILL:
                shade_cell(cell, HEADER_FILL)


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


def read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def load_noise_section_texts(output_dir: Path) -> Dict[str, str]:
    debug_dir = output_dir / "debug_tables"
    texts_path = debug_dir / "noise_section_texts.json"
    if texts_path.exists():
        payload = read_json(texts_path)
        if isinstance(payload, dict) and payload:
            return {key: str(value) for key, value in payload.items()}
    output_path = debug_dir / "noise_llm_text_output.json"
    if output_path.exists():
        payload = read_json(output_path)
        if isinstance(payload, dict) and payload:
            return {key: str(value) for key, value in payload.items()}
    input_path = debug_dir / "noise_llm_text_input.json"
    if input_path.exists():
        payload = read_json(input_path)
        rule_texts = payload.get("rule_texts") if isinstance(payload, dict) else None
        if isinstance(rule_texts, dict) and rule_texts:
            return {key: str(value) for key, value in rule_texts.items()}
    raise FileNotFoundError("未找到可重建声环境章节的文本 JSON")


def rebuild_docx_from_output(
    output_dir: Path,
    numbering: DocxNumbering,
    label_mapping: Optional[Dict[str, str]] = None,
) -> Path:
    output_dir = Path(output_dir)
    debug_dir = output_dir / "debug_tables"
    table1 = read_json(debug_dir / "noise_monitor_points_table.json")
    table2 = read_json(debug_dir / "noise_sensitive_points_result_table.json")
    indoor_path = debug_dir / "noise_indoor_result_table.json"
    table_indoor = read_json(indoor_path) if indoor_path.exists() else empty_indoor_result_table()
    table3 = read_json(debug_dir / "traffic_noise_attenuation_table.json")
    summary = read_json(debug_dir / "noise_compliance_summary.json")
    ensure_noise_table_metadata(table1, table2, table3, table_indoor)

    numbering.begin_section("noise", "声环境现状调查与评价")
    numbering.reset_section_heading_counters("noise")
    register_noise_tables(numbering, table1, table2, table3, table_indoor)

    texts = load_noise_section_texts(output_dir)
    if label_mapping:
        texts = numbering.remap_text_fields(texts, label_mapping)

    doc = build_docx(table1, table2, table3, table_indoor, summary, texts, numbering)
    finalize_section_document(doc)
    doc_path = output_dir / "noise_section.docx"
    doc.save(doc_path)
    return doc_path


def ensure_noise_table_metadata(
    table1: Dict[str, Any],
    table2: Dict[str, Any],
    table3: Dict[str, Any],
    table_indoor: Optional[Dict[str, Any]] = None,
) -> None:
    defaults = [
        (table1, NOISE_TABLE_MONITOR_POINTS, "声环境质量现状监测点"),
        (table2, NOISE_TABLE_SENSITIVE, "项目沿线敏感点声环境现状监测平均值  单位：dB(A)"),
        (table_indoor, NOISE_TABLE_INDOOR, "室内噪声监测结果统计  单位：dB(A)"),
        (table3, NOISE_TABLE_ATTENUATION, "现状公路交通噪声衰减断面监测结果  单位：dB(A)"),
    ]
    for table, table_key, caption_suffix in defaults:
        if table is None:
            continue
        table.setdefault("table_key", table_key)
        table.setdefault("caption_suffix", caption_suffix)
        table.setdefault("title", caption_suffix)


def empty_indoor_result_table() -> Dict[str, Any]:
    return {
        "table_key": NOISE_TABLE_INDOOR,
        "caption_suffix": "室内噪声监测结果统计  单位：dB(A)",
        "title": "室内噪声监测结果统计  单位：dB(A)",
        "headers": result_headers(include_flow=False),
        "rows": [],
        "warnings": [],
    }


if __name__ == "__main__":
    main()
