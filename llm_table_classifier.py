from typing import Any, Dict, List

from eia_document_router import route_table_with_llm


TABLE_TYPES = {
    "surface_water_plan",
    "surface_water_result",
    "noise_plan",
    "noise_result",
    "method_table",
    "standard_table",
    "irrelevant_table",
    "unknown",
}
TARGET_TABLE_TYPES = {
    "surface_water_plan",
    "surface_water_result",
    "noise_plan",
    "noise_result",
}


def classify_table_with_llm(table_payload: Dict[str, Any]) -> Dict[str, Any]:
    route = route_table_with_llm(table_payload)
    return classification_from_route(route)


def classification_from_route(route: Dict[str, Any]) -> Dict[str, Any]:
    candidate_rule = str(route.get("candidate_rule") or "").strip()
    table_role = str(route.get("table_role") or "").strip()
    if candidate_rule in TARGET_TABLE_TYPES:
        table_type = candidate_rule
    elif table_role == "method":
        table_type = "method_table"
    elif table_role == "standard":
        table_type = "standard_table"
    elif route.get("section_type") == "unknown" and route.get("table_role") == "unknown":
        table_type = "unknown"
    else:
        table_type = "irrelevant_table" if not route.get("is_supported_rule") else "unknown"
    return {
        "status": route.get("status", "success"),
        "table_type": table_type if table_type in TABLE_TYPES else "unknown",
        "confidence": coerce_confidence(route.get("confidence")),
        "is_target_table": table_type in TARGET_TABLE_TYPES,
        "reasons": coerce_string_list(route.get("reason")),
        "warnings": coerce_string_list(route.get("warnings")),
        "router": route,
    }


def normalize_classification(raw: Dict[str, Any]) -> Dict[str, Any]:
    table_type = str(raw.get("table_type") or "unknown").strip()
    if table_type not in TABLE_TYPES:
        table_type = "unknown"
    confidence = coerce_confidence(raw.get("confidence"))
    is_target = raw.get("is_target_table")
    if not isinstance(is_target, bool):
        is_target = table_type in TARGET_TABLE_TYPES
    return {
        "status": "success",
        "table_type": table_type,
        "confidence": confidence,
        "is_target_table": bool(is_target),
        "reasons": coerce_string_list(raw.get("reasons")),
        "warnings": coerce_string_list(raw.get("warnings")),
    }


def skipped_result(reason: str) -> Dict[str, Any]:
    return {
        "status": "skipped",
        "skip_reason": reason,
        "table_type": "unknown",
        "confidence": 0.0,
        "is_target_table": False,
        "reasons": [],
        "warnings": [reason],
    }


def failed_result(error: BaseException) -> Dict[str, Any]:
    return {
        "status": "failed",
        "error": str(error),
        "table_type": "unknown",
        "confidence": 0.0,
        "is_target_table": False,
        "reasons": [],
        "warnings": [str(error)],
    }


def coerce_confidence(value: Any) -> float:
    try:
        confidence = float(value)
    except (TypeError, ValueError):
        return 0.0
    return max(0.0, min(1.0, confidence))


def coerce_string_list(value: Any) -> List[str]:
    if isinstance(value, list):
        return [str(item).strip() for item in value if str(item).strip()]
    text = str(value or "").strip()
    return [text] if text else []
