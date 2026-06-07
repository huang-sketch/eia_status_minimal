from typing import Any, Dict, List

from surface_water_pipeline import (
    SURFACE_WATER_FACTORS,
    extract_point_code,
    normalize_factor,
    parse_numeric_value,
)


def adapt_cli_surface_water_records(records: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    adapted: List[Dict[str, Any]] = []
    for record in records:
        if str(record.get("monitor_type") or "").strip() != "surface_water":
            continue

        point = str(record.get("point") or "").strip()
        point_code = extract_point_code(point)
        raw_factor = str(record.get("factor") or "").strip()
        factor = normalize_factor(raw_factor)
        if not point_code or factor not in SURFACE_WATER_FACTORS:
            continue

        raw_value = str(record.get("value") or "").strip()
        numeric_value, value_warning = parse_numeric_value(raw_value)
        needs_review = bool(record.get("needs_review")) or bool(value_warning)
        warning = str(record.get("warning") or value_warning or "").strip()

        adapted.append(
            {
                "source_type": record.get("source_type") or "监测报告",
                "monitor_type": "surface_water",
                "point_code": point_code,
                "point": point or None,
                "sample_date": record.get("sample_date"),
                "factor": factor,
                "raw_factor": raw_factor or factor,
                "value": raw_value,
                "numeric_value": numeric_value,
                "unit": record.get("unit"),
                "sample_character": record.get("sample_character"),
                "evidence": record.get("evidence"),
                "source_file": record.get("source_file"),
                "source_table": record.get("source_table") or record.get("chunk_id"),
                "needs_review": needs_review,
                "warning": warning,
                "extraction_method": record.get("extraction_method") or "cli",
                "confidence": record.get("confidence"),
            }
        )
    return adapted


def merge_cli_and_rule_records(
    cli_records: List[Dict[str, Any]],
    rule_records: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    merged: List[Dict[str, Any]] = []
    seen: set[str] = set()

    for record in cli_records:
        key = record_merge_key(record)
        if key in seen:
            continue
        seen.add(key)
        merged.append(record)

    fallback_added = 0
    for record in rule_records:
        key = record_merge_key(record)
        if key in seen:
            continue
        seen.add(key)
        merged.append(record)
        fallback_added += 1

    return merged


def record_merge_key(record: Dict[str, Any]) -> str:
    return "|".join(
        [
            str(record.get("point_code") or "").strip(),
            str(record.get("sample_date") or "").strip(),
            normalize_factor(record.get("factor")),
        ]
    )


def build_extraction_summary(
    cli_records: List[Dict[str, Any]],
    rule_records: List[Dict[str, Any]],
    merged_records: List[Dict[str, Any]],
    extraction_available: bool,
) -> Dict[str, Any]:
    fallback_added = max(0, len(merged_records) - len(cli_records))
    return {
        "extraction_available": extraction_available,
        "cli_count": len(cli_records),
        "rule_count": len(rule_records),
        "merged_count": len(merged_records),
        "fallback_added": fallback_added,
        "fallback_used": fallback_added > 0,
    }
