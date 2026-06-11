import json
import math
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from word_processor import load_docx_chunks
from table_schema_mapper import apply_header_mapping, record_data_validation, resolve_table_schema

EXTRACTION_RECORDS_FILENAME = "extraction/records.json"


STANDARD_NAME = "地表水环境质量标准 GB3838-2002"
SURFACE_WATER_FACTORS = {
    "水温",
    "pH值",
    "pH",
    "悬浮物",
    "SS",
    "高锰酸盐指数",
    "石油类",
    "氨氮",
    "总磷",
    "TP",
    "溶解氧",
    "DO",
    "化学需氧量",
    "COD",
}

LIMITS: Dict[str, Dict[str, Any]] = {
    "Ⅰ类": {
        "pH值": (6.0, 9.0),
        "溶解氧": 7.5,
        "高锰酸盐指数": 2.0,
        "氨氮": 0.15,
        "总磷": 0.02,
        "石油类": 0.05,
        "化学需氧量": 15.0,
    },
    "Ⅱ类": {
        "pH值": (6.0, 9.0),
        "溶解氧": 6.0,
        "高锰酸盐指数": 4.0,
        "氨氮": 0.5,
        "总磷": 0.1,
        "石油类": 0.05,
        "化学需氧量": 15.0,
    },
    "Ⅲ类": {
        "pH值": (6.0, 9.0),
        "溶解氧": 5.0,
        "高锰酸盐指数": 6.0,
        "氨氮": 1.0,
        "总磷": 0.2,
        "石油类": 0.05,
        "化学需氧量": 20.0,
    },
    "Ⅳ类": {
        "pH值": (6.0, 9.0),
        "溶解氧": 3.0,
        "高锰酸盐指数": 10.0,
        "氨氮": 1.5,
        "总磷": 0.3,
        "石油类": 0.5,
        "化学需氧量": 30.0,
    },
    "Ⅴ类": {
        "pH值": (6.0, 9.0),
        "溶解氧": 2.0,
        "高锰酸盐指数": 15.0,
        "氨氮": 2.0,
        "总磷": 0.4,
        "石油类": 1.0,
        "化学需氧量": 40.0,
    },
}


def main() -> None:
    from record_adapter import (
        adapt_cli_surface_water_records,
        build_extraction_summary,
        merge_cli_and_rule_records,
    )

    input_dir = Path(os.getenv("EIA_INPUT_DIR", "input"))
    output_dir = Path(os.getenv("EIA_OUTPUT_DIR", "output"))
    output_dir.mkdir(parents=True, exist_ok=True)
    debug_dir = output_dir / "debug_tables"
    debug_dir.mkdir(parents=True, exist_ok=True)

    plan_path = find_docx(input_dir, "方案")
    report_path = find_docx(input_dir, "报告")

    standard_config = parse_standard_config(plan_path)
    extraction_records_path = output_dir / EXTRACTION_RECORDS_FILENAME
    extraction_available = extraction_records_path.exists()
    cli_raw_records = read_json(extraction_records_path) if extraction_available else []
    if not isinstance(cli_raw_records, list):
        cli_raw_records = []

    cli_records = adapt_cli_surface_water_records(cli_raw_records)
    rule_records = parse_surface_water_records(report_path)
    records = merge_cli_and_rule_records(cli_records, rule_records)
    align_surface_water_record_point_codes(records, standard_config)
    write_json(
        debug_dir / "extraction_summary.json",
        build_extraction_summary(cli_records, rule_records, records, extraction_available),
    )
    compliance_results = evaluate_compliance(records, standard_config)
    data_warnings = list(standard_config.get("warnings") or [])
    if records and not standard_config.get("points"):
        data_warnings.append("已识别地表水监测数据，但监测方案中缺少可用水质目标")
    missing_standard_points = sorted(
        {
            str(item.get("point_code") or "")
            for item in compliance_results
            if item.get("method") == "missing_standard" or not item.get("standard_class")
        }
    )
    if missing_standard_points:
        data_warnings.append("缺少水质目标: " + "、".join(item for item in missing_standard_points if item))
    record_data_validation(
        output_dir,
        "surface_water_data",
        valid=not data_warnings,
        warnings=data_warnings,
    )

    write_json(output_dir / "standard_config.json", standard_config)
    write_json(output_dir / "monitoring_records.json", records)
    write_json(output_dir / "compliance_results.json", compliance_results)

    print(f"standard_config: {len(standard_config['points'])} point(s)")
    print(f"cli_records: {len(cli_records)} record(s)")
    print(f"rule_records: {len(rule_records)} record(s)")
    print(f"monitoring_records: {len(records)} record(s)")
    print(f"compliance_results: {len(compliance_results)} result(s)")


def find_docx(input_dir: Path, keyword: str) -> Path:
    matches = sorted(path for path in input_dir.glob("*.docx") if keyword in path.name)
    if not matches:
        raise FileNotFoundError(f"No .docx file containing {keyword!r} under {input_dir}")
    return matches[0]


def parse_standard_config(plan_path: Path) -> Dict[str, Any]:
    chunks = load_docx_chunks(plan_path)
    points: Dict[str, Dict[str, Any]] = {}
    warnings: List[str] = []

    for chunk in chunks:
        if chunk.get("kind") != "table":
            continue
        table_text = chunk.get("text") or ""
        if "地表水" not in table_text and "WJ" not in table_text:
            continue
        rows = parse_table_rows(table_text)
        if not rows:
            continue
        headers = rows[0]
        schema_result = resolve_table_schema(
            "surface_water_plan",
            headers,
            sample_rows=rows[1:4],
            table_title=str((chunk.get("metadata") or {}).get("table_title") or ""),
            output_dir=Path(os.getenv("EIA_OUTPUT_DIR", "output")),
            source_table=str(chunk.get("chunk_id") or ""),
            job_id=Path(os.getenv("EIA_OUTPUT_DIR", "output")).parent.name,
        )
        if not schema_result["validation"]["valid"]:
            continue
        header_map = build_header_map(headers)
        for row in rows[1:]:
            if len(row) < 4:
                continue
            canonical = apply_header_mapping(headers, row, schema_result["mapping"])
            point_code = extract_point_code(canonical.get("point_code") or get_row_value(row, header_map, "序号", 0))
            standard_class = normalize_standard_class(
                canonical.get("standard_class")
                or get_row_value(row, header_map, "评价标准", None)
                or get_row_value(row, header_map, "水质目标", None)
                or get_row_value(row, header_map, "标准", 3)
            )
            if not point_code:
                continue
            river_name = (
                canonical.get("river_name")
                or get_row_value(row, header_map, "河流名称", None)
                or get_row_value(row, header_map, "河流", None)
                or get_row_value(row, header_map, "水体", None)
            )
            points[point_code] = {
                "point_code": point_code,
                "river_name": river_name,
                "standard_name": STANDARD_NAME,
                "standard_class": standard_class,
                "section_name": (
                    canonical.get("section_name")
                    or get_row_value(row, header_map, "取样断面", None)
                    or get_row_value(row, header_map, "监测断面", None)
                    or get_row_value(row, header_map, "断面", 4)
                ),
                "source_table": chunk.get("chunk_id"),
                "evidence": row_to_evidence(headers, row),
                "limits": LIMITS.get(standard_class, {}),
            }
            if not standard_class:
                warnings.append(f"{point_code} 缺少水质目标")

    return {
        "standard_name": STANDARD_NAME,
        "points": points,
        "warnings": warnings,
        "source_file": str(plan_path),
    }


def parse_surface_water_records(report_path: Path) -> List[Dict[str, Any]]:
    records: List[Dict[str, Any]] = []
    chunks = load_docx_chunks(report_path)
    for chunk in chunks:
        if chunk.get("kind") != "table":
            continue
        table_title = ((chunk.get("metadata") or {}).get("table_title") or "")
        table_text = chunk.get("text") or ""
        if not is_surface_water_result_table(table_title, table_text):
            continue
        rows = parse_table_rows(table_text)
        if not rows:
            continue
        headers = rows[0]
        schema_result = resolve_table_schema(
            "surface_water_result",
            headers,
            sample_rows=rows[1:4],
            table_title=table_title,
            output_dir=Path(os.getenv("EIA_OUTPUT_DIR", "output")),
            source_table=str(chunk.get("chunk_id") or ""),
            job_id=Path(os.getenv("EIA_OUTPUT_DIR", "output")).parent.name,
        )
        if not schema_result["validation"]["valid"]:
            continue
        for row in rows[1:]:
            if len(row) < 5:
                continue
            canonical = apply_header_mapping(headers, row, schema_result["mapping"])
            point = (canonical.get("point") or row[0]).strip()
            point_code = extract_point_code(point)
            factor = normalize_factor(canonical.get("factor") or row[2])
            if not point_code or factor not in SURFACE_WATER_FACTORS:
                continue
            raw_value = (canonical.get("value") or row[4]).strip()
            numeric_value, value_warning = parse_numeric_value(raw_value)
            record = {
                "source_type": "监测报告",
                "monitor_type": "surface_water",
                "point_code": point_code,
                "point": point,
                "sample_date": (canonical.get("sample_date") or row[1]).strip(),
                "factor": factor,
                "raw_factor": (canonical.get("factor") or row[2]).strip(),
                "value": raw_value,
                "numeric_value": numeric_value,
                "unit": (canonical.get("unit") or row[3]).strip() if len(row) > 3 else None,
                "sample_character": (canonical.get("sample_character") or row[5]).strip() if len(row) > 5 else None,
                "evidence": row_to_evidence(headers, row),
                "source_file": str(report_path),
                "source_table": chunk.get("chunk_id"),
                "needs_review": bool(value_warning),
                "warning": value_warning,
            }
            records.append(record)
    return records


def evaluate_compliance(
    records: List[Dict[str, Any]],
    standard_config: Dict[str, Any],
) -> List[Dict[str, Any]]:
    points = standard_config.get("points") or {}
    water_temps = {
        (record.get("point_code"), record.get("sample_date")): record.get("numeric_value")
        for record in records
        if record.get("factor") == "水温" and record.get("numeric_value") is not None
    }

    results: List[Dict[str, Any]] = []
    for record in records:
        point_code = record.get("point_code")
        config = points.get(point_code) or {}
        standard_class = config.get("standard_class")
        factor = record.get("factor")
        value = record.get("numeric_value")
        warnings: List[str] = []
        if record.get("warning"):
            warnings.append(str(record["warning"]))
        if not config:
            warnings.append(f"未在 standard_config 中找到 {point_code}")
        if not standard_class:
            warnings.append("未识别标准类别")

        standard_index: Optional[float] = None
        is_compliant: Optional[bool] = None
        limit_value: Any = None
        method = None

        if factor in {"水温", "悬浮物"}:
            method = "not_applicable"
            warnings.append(f"{factor} 暂不适用 GB3838 单项标准指数判定")
        elif value is None:
            method = "invalid_value"
            warnings.append("未识别有效数值")
        elif standard_class not in LIMITS:
            method = "missing_standard"
        elif factor in {"pH值", "pH"}:
            method = "ph_standard_index"
            low, high = LIMITS[standard_class]["pH值"]
            limit_value = {"lower": low, "upper": high}
            standard_index = calc_ph_index(value, low, high)
            is_compliant = low <= value <= high
        elif factor == "溶解氧":
            method = "do_standard_index"
            limit_value = LIMITS[standard_class]["溶解氧"]
            temperature = water_temps.get((point_code, record.get("sample_date")))
            if temperature is None:
                warnings.append("未找到同点位同日期水温，无法计算 DO 标准指数")
            else:
                standard_index = calc_do_index(value, float(limit_value), float(temperature))
                is_compliant = standard_index <= 1
        elif factor in {"高锰酸盐指数", "氨氮", "总磷", "石油类", "化学需氧量"}:
            method = "value_div_limit"
            limit_value = LIMITS[standard_class][factor]
            standard_index = value / float(limit_value)
            is_compliant = standard_index <= 1
        else:
            method = "not_applicable"
            warnings.append(f"{factor} 暂不判定")

        result = {
            **record,
            "river_name": config.get("river_name"),
            "standard_name": config.get("standard_name"),
            "standard_class": standard_class,
            "limit_value": limit_value,
            "standard_index": round(standard_index, 4) if standard_index is not None else None,
            "is_compliant": is_compliant,
            "method": method,
            "needs_review": bool(warnings) or bool(record.get("needs_review")),
            "warning": "; ".join(warnings),
        }
        results.append(result)
    return results


def parse_table_rows(table_text: str) -> List[List[str]]:
    rows: List[List[str]] = []
    for line in table_text.splitlines():
        body = re.sub(r"^row\s+\d+:\s*", "", line.strip(), flags=re.IGNORECASE)
        cells = [
            match.group(2).strip()
            for match in re.finditer(
                r"col\s+(\d+):\s*(.*?)(?=\s*\|\s*col\s+\d+:|$)",
                body,
                flags=re.IGNORECASE,
            )
        ]
        if any(cells):
            rows.append(cells)
    return rows


def build_header_map(headers: List[str]) -> Dict[str, int]:
    return {str(header or "").strip(): index for index, header in enumerate(headers)}


def get_row_value(
    row: List[str],
    header_map: Dict[str, int],
    header_keyword: str,
    fallback_index: Optional[int],
) -> Optional[str]:
    for header, index in header_map.items():
        if header_keyword in header and index < len(row):
            value = row[index].strip()
            if value:
                return value
    if fallback_index is not None and fallback_index < len(row):
        value = row[fallback_index].strip()
        return value or None
    return None


def row_to_evidence(headers: List[str], row: List[str]) -> Dict[str, Any]:
    evidence: Dict[str, Any] = {}
    for index, value in enumerate(row):
        key = headers[index] if index < len(headers) and headers[index] else f"col_{index + 1}"
        evidence[key] = value
    return evidence


def extract_point_code(text: str) -> Optional[str]:
    match = re.search(r"(?<![A-Za-z0-9])(?:WJ|D)\s*\d+", text or "", flags=re.IGNORECASE)
    if match:
        return re.sub(r"\s+", "", match.group(0)).upper()
    return match.group(0) if match else None


def align_surface_water_record_point_codes(
    records: List[Dict[str, Any]],
    standard_config: Dict[str, Any],
) -> None:
    points = standard_config.get("points") or {}
    if not points:
        return
    for record in records:
        code = str(record.get("point_code") or "").strip().upper()
        if not code or code in points:
            continue
        match = re.search(r"(\d+)$", code)
        if not match:
            continue
        candidates = [
            point_code
            for point_code in points
            if re.search(rf"{re.escape(match.group(1))}$", str(point_code or ""))
        ]
        if len(candidates) == 1:
            record["raw_point_code"] = code
            record["point_code"] = candidates[0]


def normalize_standard_class(text: str) -> Optional[str]:
    text = str(text or "").strip()
    for item in ("Ⅰ类", "Ⅱ类", "Ⅲ类", "Ⅳ类", "Ⅴ类"):
        if item in text:
            return item
    chinese_mapping = {
        "一类": "Ⅰ类",
        "二类": "Ⅱ类",
        "三类": "Ⅲ类",
        "四类": "Ⅳ类",
        "五类": "Ⅴ类",
        "1类": "Ⅰ类",
        "2类": "Ⅱ类",
        "3类": "Ⅲ类",
        "4类": "Ⅳ类",
        "5类": "Ⅴ类",
    }
    compact_text = re.sub(r"\s+", "", text)
    for key, value in chinese_mapping.items():
        if key in compact_text:
            return value
    mapping = {
        "I类": "Ⅰ类",
        "II类": "Ⅱ类",
        "III类": "Ⅲ类",
        "IV类": "Ⅳ类",
        "V类": "Ⅴ类",
    }
    for key, value in mapping.items():
        if key in text.upper().replace("Ⅲ", "III").replace("Ⅱ", "II").replace("Ⅳ", "IV"):
            return value
    return None


def is_surface_water_result_table(table_title: str, table_text: str) -> bool:
    text = f"{table_title}\n{table_text}"
    if any(marker in text for marker in ("地表水检测结果", "地表水监测结果", "水质检测结果", "水质监测结果")):
        return True
    if "采样位置" in text or "监测断面" in text or "断面位置" in text:
        if any(factor in text for factor in ("pH", "DO", "溶解氧", "高锰酸盐指数", "氨氮", "总磷", "石油类")):
            return True
    if re.search(r"\bWJ\s*\d+\b", text, flags=re.IGNORECASE) and any(
        factor in text for factor in ("pH", "DO", "溶解氧", "高锰酸盐指数", "氨氮", "总磷", "石油类")
    ):
        return True
    return False


def normalize_factor(text: str) -> str:
    text = str(text or "").strip()
    aliases = {
        "pH": "pH值",
        "PH": "pH值",
        "TP": "总磷",
        "NH3-N": "氨氮",
        "NH₃-N": "氨氮",
        "DO": "溶解氧",
        "COD": "化学需氧量",
        "SS": "悬浮物",
    }
    return aliases.get(text, text)


def parse_numeric_value(text: str) -> Tuple[Optional[float], str]:
    text = str(text or "").strip()
    match = re.search(r"\d+(?:\.\d+)?", text)
    if not match:
        return None, f"未识别数值: {text}"
    value = float(match.group(0))
    if text.startswith("<"):
        return value / 2, f"低于检出限，按 1/2 检出限参与计算: {text}"
    return value, ""


def calc_ph_index(value: float, low: float, high: float) -> float:
    if value <= 7.0:
        return (7.0 - value) / (7.0 - low)
    return (value - 7.0) / (high - 7.0)


def calc_do_index(value: float, limit: float, temperature: float) -> float:
    do_saturation = 468.0 / (31.6 + temperature)
    if value >= limit:
        denominator = do_saturation - limit
        if math.isclose(denominator, 0.0):
            return 0.0 if value >= limit else 10.0
        return abs(do_saturation - value) / abs(denominator)
    return 10.0 - 9.0 * value / limit


def write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


if __name__ == "__main__":
    main()
