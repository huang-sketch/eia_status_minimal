import json
from typing import Any, Dict, List

from llm_client import LlmProfile, chat_completion_json_object, llm_enabled


PROJECT_TYPES = {"highway", "waterway", "railway", "airport", "unknown"}
DOCUMENT_ROLES = {"plan", "report", "reference", "unknown"}
SECTION_TYPES = {
    "noise",
    "surface_water",
    "regional_status",
    "ecology",
    "air",
    "vibration",
    "sediment",
    "unknown",
}
TABLE_ROLES = {"plan_points", "monitor_results", "standard", "frequency", "method", "metadata", "unknown"}
SUPPORTED_RULES = {"noise_plan", "noise_result", "surface_water_plan", "surface_water_result"}

RULE_BY_SECTION_ROLE = {
    ("noise", "plan_points"): "noise_plan",
    ("noise", "monitor_results"): "noise_result",
    ("surface_water", "plan_points"): "surface_water_plan",
    ("surface_water", "monitor_results"): "surface_water_result",
}


def route_table_with_llm(table_payload: Dict[str, Any]) -> Dict[str, Any]:
    """Classify a document table for rule routing without extracting final records."""
    if not llm_enabled():
        return skipped_route("llm_disabled_or_missing_api_key", table_payload)

    raw = chat_completion_json_object(
        [
            {
                "role": "system",
                "content": (
                    "Return strict JSON only. You are an EIA document router. "
                    "Classify document chunks and choose an existing rule if appropriate. "
                    "Do not extract monitoring records, calculate standards, or invent values."
                ),
            },
            {"role": "user", "content": build_route_prompt(table_payload)},
        ],
        profile=LlmProfile.fast,
        max_retries=1,
        label="eia_document_router",
    )
    return normalize_route(raw, table_payload)


def diagnose_rule_failure_with_llm(failure_payload: Dict[str, Any]) -> Dict[str, Any]:
    """Diagnose why a rule failed and propose review-safe hints."""
    if not llm_enabled():
        return {
            "status": "skipped",
            "skip_reason": "llm_disabled_or_missing_api_key",
            "failure_type": "unknown",
            "candidate_header_mapping": {},
            "candidate_aliases": [],
            "manual_review_items": ["LLM diagnosis skipped: missing API key or disabled"],
            "confidence": 0.0,
        }

    raw = chat_completion_json_object(
        [
            {
                "role": "system",
                "content": (
                    "Return strict JSON only. Diagnose EIA table routing or schema failures. "
                    "Never invent headers, point codes, values, standards, or conclusions."
                ),
            },
            {"role": "user", "content": build_failure_prompt(failure_payload)},
        ],
        profile=LlmProfile.fast,
        max_retries=1,
        label="eia_rule_failure_diagnosis",
    )
    return normalize_diagnosis(raw, failure_payload)


def heuristic_route(table_payload: Dict[str, Any]) -> Dict[str, Any]:
    text = compact_table_text(table_payload)
    project_type = infer_project_type(text)
    document_role = infer_document_role(text, table_payload)
    section_type = infer_section_type(text)
    table_role = infer_table_role(text, section_type)
    candidate_rule = RULE_BY_SECTION_ROLE.get((section_type, table_role), "")
    if project_type in {"waterway", "railway", "airport"}:
        candidate_rule = ""
    confidence = heuristic_confidence(section_type, table_role, candidate_rule, text)
    needs_review = confidence < 0.75 or not candidate_rule
    return {
        "status": "heuristic",
        "project_type": project_type,
        "document_role": document_role,
        "section_type": section_type,
        "table_role": table_role,
        "candidate_rule": candidate_rule,
        "confidence": confidence,
        "is_supported_rule": candidate_rule in SUPPORTED_RULES,
        "needs_review": needs_review,
        "reason": "heuristic routing from title, context, headers, and sample rows",
        "warnings": [] if candidate_rule else ["No implemented rule matched this table route."],
    }


def skipped_route(reason: str, table_payload: Dict[str, Any]) -> Dict[str, Any]:
    route = heuristic_route(table_payload)
    return {
        **route,
        "status": "skipped",
        "skip_reason": reason,
        "warnings": [reason, *route.get("warnings", [])],
    }


def build_route_prompt(table_payload: Dict[str, Any]) -> str:
    compact_payload = compact_payload_for_llm(table_payload)
    return (
        "Classify this EIA Word table for routing to an existing rule.\n"
        f"Allowed project_type: {sorted(PROJECT_TYPES)}.\n"
        f"Allowed document_role: {sorted(DOCUMENT_ROLES)}.\n"
        f"Allowed section_type: {sorted(SECTION_TYPES)}.\n"
        f"Allowed table_role: {sorted(TABLE_ROLES)}.\n"
        f"Implemented candidate_rule values: {sorted(SUPPORTED_RULES)}; use empty string when unsupported.\n"
        "Return JSON keys: project_type, document_role, section_type, table_role, "
        "candidate_rule, confidence, reason, needs_review, warnings.\n"
        "Rules:\n"
        "- Only route/classify. Do not extract records or conclusions.\n"
        "- Prefer supported highway noise/surface-water rules when clearly applicable.\n"
        "- For waterway, railway, or airport content without implemented rules, set candidate_rule to empty string.\n"
        "- Low confidence or unsupported content must set needs_review=true.\n"
        f"Input table:\n{json.dumps(compact_payload, ensure_ascii=False)}"
    )


def build_failure_prompt(failure_payload: Dict[str, Any]) -> str:
    compact_payload = {
        "candidate_rule": failure_payload.get("candidate_rule"),
        "schema": failure_payload.get("schema"),
        "missing_fields": failure_payload.get("missing_fields") or failure_payload.get("missing_after_llm") or [],
        "headers": failure_payload.get("headers") or failure_payload.get("raw_headers") or [],
        "sample_rows": failure_payload.get("sample_rows") or failure_payload.get("normalized_rows") or [],
        "table_title": failure_payload.get("table_title"),
        "context": failure_payload.get("context") or failure_payload.get("context_before"),
        "rule_error": failure_payload.get("error") or failure_payload.get("rule_error"),
    }
    return (
        "Diagnose why the existing EIA rule could not parse this table.\n"
        "Return JSON keys: failure_type, candidate_header_mapping, candidate_aliases, "
        "manual_review_items, confidence, warnings.\n"
        "Constraints:\n"
        "- candidate_header_mapping values must be exact strings from headers.\n"
        "- Do not invent point codes, measurement values, standards, or conclusions.\n"
        "- If the table appears unsupported, say so in manual_review_items.\n"
        f"Failure payload:\n{json.dumps(compact_payload, ensure_ascii=False)}"
    )


def normalize_route(raw: Dict[str, Any], table_payload: Dict[str, Any]) -> Dict[str, Any]:
    heuristic = heuristic_route(table_payload)
    if not isinstance(raw, dict):
        return {**heuristic, "status": "failed", "warnings": ["LLM route output was not an object"]}

    project_type = coerce_choice(raw.get("project_type"), PROJECT_TYPES, heuristic["project_type"])
    document_role = coerce_choice(raw.get("document_role"), DOCUMENT_ROLES, heuristic["document_role"])
    section_type = coerce_choice(raw.get("section_type"), SECTION_TYPES, heuristic["section_type"])
    table_role = coerce_choice(raw.get("table_role"), TABLE_ROLES, heuristic["table_role"])
    candidate_rule = str(raw.get("candidate_rule") or "").strip()
    inferred_rule = RULE_BY_SECTION_ROLE.get((section_type, table_role), "")
    if candidate_rule not in SUPPORTED_RULES:
        candidate_rule = inferred_rule if inferred_rule in SUPPORTED_RULES else ""
    if project_type in {"waterway", "railway", "airport"}:
        candidate_rule = ""
    confidence = coerce_confidence(raw.get("confidence"), heuristic["confidence"])
    is_supported = candidate_rule in SUPPORTED_RULES
    needs_review = raw.get("needs_review")
    if not isinstance(needs_review, bool):
        needs_review = confidence < 0.75 or not is_supported

    return {
        "status": "success",
        "project_type": project_type,
        "document_role": document_role,
        "section_type": section_type,
        "table_role": table_role,
        "candidate_rule": candidate_rule,
        "confidence": confidence,
        "is_supported_rule": is_supported,
        "needs_review": needs_review,
        "reason": str(raw.get("reason") or heuristic["reason"]),
        "warnings": coerce_string_list(raw.get("warnings")),
        "heuristic": heuristic,
    }


def normalize_diagnosis(raw: Dict[str, Any], failure_payload: Dict[str, Any]) -> Dict[str, Any]:
    headers = {str(item or "").strip() for item in failure_payload.get("headers") or failure_payload.get("raw_headers") or []}
    mapping = raw.get("candidate_header_mapping") if isinstance(raw, dict) else {}
    accepted_mapping: Dict[str, str] = {}
    if isinstance(mapping, dict):
        for field, header in mapping.items():
            field_text = str(field or "").strip()
            header_text = str(header or "").strip()
            if field_text and header_text in headers:
                accepted_mapping[field_text] = header_text

    return {
        "status": "success" if isinstance(raw, dict) else "failed",
        "failure_type": str((raw or {}).get("failure_type") or "unknown") if isinstance(raw, dict) else "unknown",
        "candidate_header_mapping": accepted_mapping,
        "candidate_aliases": coerce_string_list((raw or {}).get("candidate_aliases") if isinstance(raw, dict) else []),
        "manual_review_items": coerce_string_list((raw or {}).get("manual_review_items") if isinstance(raw, dict) else []),
        "confidence": coerce_confidence((raw or {}).get("confidence") if isinstance(raw, dict) else None, 0.0),
        "warnings": coerce_string_list((raw or {}).get("warnings") if isinstance(raw, dict) else []),
    }


def compact_payload_for_llm(table_payload: Dict[str, Any]) -> Dict[str, Any]:
    return {
        "chunk_id": table_payload.get("chunk_id"),
        "source_file": table_payload.get("source_file"),
        "table_title": table_payload.get("table_title"),
        "context_before": table_payload.get("context_before"),
        "normalized_table_title": table_payload.get("normalized_table_title"),
        "normalized_context_before": table_payload.get("normalized_context_before"),
        "row_count": table_payload.get("row_count"),
        "col_counts": table_payload.get("col_counts"),
        "empty_cell_count": table_payload.get("empty_cell_count"),
        "raw_headers": table_payload.get("raw_headers"),
        "normalized_headers": table_payload.get("normalized_headers"),
        "raw_rows": table_payload.get("raw_rows"),
        "normalized_rows": table_payload.get("normalized_rows"),
    }


def compact_table_text(table_payload: Dict[str, Any]) -> str:
    compact = compact_payload_for_llm(table_payload)
    return json.dumps(compact, ensure_ascii=False).lower()


def infer_project_type(text: str) -> str:
    if has_any(text, ["航道", "码头", "港区", "疏浚", "船闸"]):
        return "waterway"
    if has_any(text, ["铁路", "轨道", "声屏障", "振动"]):
        return "railway"
    if has_any(text, ["机场", "飞行", "跑道", "航空", "等值线"]):
        return "airport"
    if has_any(text, ["高速", "公路", "道路", "路线", "桩号"]):
        return "highway"
    return "unknown"


def infer_document_role(text: str, table_payload: Dict[str, Any]) -> str:
    source = str(table_payload.get("source_file") or "").lower()
    if "方案" in source or has_any(text, ["方案", "布点", "监测点", "执行标准", "频次"]):
        return "plan"
    if "报告" in source or has_any(text, ["报告", "检测结果", "监测结果", "采样日期", "laeq"]):
        return "report"
    if has_any(text, ["公报", "百度百科", "统计年鉴", "公开资料"]):
        return "reference"
    return "unknown"


def infer_section_type(text: str) -> str:
    title_text = title_context_text(text)
    if has_any(title_text, ["声环境", "噪声", "声现状"]):
        return "noise"
    if has_any(title_text, ["地表水", "水环境", "水质"]):
        return "surface_water"
    if has_any(text, ["laeq", "l10", "l50", "l90", "噪声", "声环境", "声现状", "等效声级", "等效连续a声级"]):
        return "noise"
    if has_any(text, ["地表水", "水环境", "水质", "断面", "溶解氧", "高锰酸盐", "氨氮", "总磷", "采样位置"]):
        return "surface_water"
    if has_any(text, ["区域环境", "自然环境", "社会环境", "行政区划"]):
        return "regional_status"
    if has_any(text, ["生态", "水生生物", "底栖", "浮游"]):
        return "ecology"
    if has_any(text, ["环境空气", "大气", "pm10", "pm2.5", "二氧化硫", "氮氧化物"]):
        return "air"
    if has_any(text, ["振动", "铅垂向", "vlz"]):
        return "vibration"
    if has_any(text, ["底泥", "沉积物"]):
        return "sediment"
    return "unknown"


def infer_table_role(text: str, section_type: str) -> str:
    title_text = title_context_text(text)
    if has_any(text, ["检测方法", "监测方法", "分析方法", "仪器", "检出限"]):
        return "method"
    if has_any(title_text, ["方法", "仪器", "检出限"]):
        return "method"
    if has_any(text, ["标准限值", "执行标准", "评价标准", "标准类别", "水质目标"]):
        if has_any(text, ["监测点", "点位", "断面", "敏感点"]):
            return "plan_points"
        return "standard"
    if has_any(title_text, ["标准", "限值"]) and not has_any(title_text, ["监测点", "点位", "监测结果", "检测结果"]):
        return "standard"
    if has_any(title_text, ["监测结果", "检测结果", "测量结果"]):
        return "monitor_results"
    if section_type == "noise" and has_any(text, ["laeq", "l10", "l50", "l90", "昼间", "夜间", "车流量", "等效声级", "等效连续a声级"]):
        return "monitor_results"
    if section_type == "surface_water" and has_any(text, ["检测结果", "监测结果", "检测值", "采样日期", "单位"]):
        return "monitor_results"
    if has_any(text, ["监测点", "点位", "布点", "断面", "敏感点", "监测位置"]):
        return "plan_points"
    if has_any(text, ["频次", "监测时间", "连续", "昼", "夜"]):
        return "frequency"
    return "unknown"


def title_context_text(text: str) -> str:
    # compact_table_text serializes title/context before headers and rows, so this prefix carries the strongest labels.
    return str(text or "")[:800]


def heuristic_confidence(section_type: str, table_role: str, candidate_rule: str, text: str) -> float:
    score = 0.25
    if section_type != "unknown":
        score += 0.25
    if table_role != "unknown":
        score += 0.2
    if candidate_rule:
        score += 0.2
    if has_any(text, ["laeq", "溶解氧", "采样日期", "监测点编号", "检测结果"]):
        score += 0.1
    return min(0.95, score)


def has_any(text: str, needles: List[str]) -> bool:
    return any(needle.lower() in text for needle in needles)


def coerce_choice(value: Any, allowed: set, default: str) -> str:
    text = str(value or "").strip()
    return text if text in allowed else default


def coerce_confidence(value: Any, default: float = 0.0) -> float:
    try:
        confidence = float(value)
    except (TypeError, ValueError):
        return default
    return max(0.0, min(1.0, confidence))


def coerce_string_list(value: Any) -> List[str]:
    if isinstance(value, list):
        return [str(item).strip() for item in value if str(item).strip()]
    text = str(value or "").strip()
    return [text] if text else []
