import json
import os
import re
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional

from llm_client import LlmProfile, chat_completion_json_object


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_CONFIG_PATH = BASE_DIR / "config" / "table_schema_mappings.json"
DEBUG_FILENAMES = {
    "detection": "table_schema_detection.json",
    "llm_input": "table_schema_llm_input.json",
    "llm_output": "table_schema_llm_output.json",
    "validation": "table_schema_validation.json",
}


def resolve_table_schema(
    schema_name: str,
    headers: Iterable[str],
    *,
    sample_rows: Optional[List[List[str]]] = None,
    table_title: str = "",
    context: str = "",
    output_dir: Optional[Path] = None,
    source_table: str = "",
    job_id: str = "",
    enable_llm: Optional[bool] = None,
) -> Dict[str, Any]:
    started = time.perf_counter()
    header_list = [str(header or "").strip() for header in headers]
    config = load_mapping_config()
    schema = (config.get("schemas") or {}).get(schema_name) or {}
    aliases = schema.get("aliases") or {}
    required_fields = list(schema.get("required_fields") or [])
    mapping: Dict[str, str] = {}
    sources: Dict[str, str] = {}

    for field, field_aliases in aliases.items():
        matched = match_header(header_list, field_aliases)
        if matched:
            mapping[field] = matched
            sources[field] = "formal"

    for candidate in config.get("candidates") or []:
        if candidate.get("schema") != schema_name:
            continue
        field = str(candidate.get("field") or "")
        header = str(candidate.get("header") or "")
        if field and field not in mapping and header in header_list:
            mapping[field] = header
            sources[field] = "candidate"

    missing_before_llm = [field for field in required_fields if field not in mapping]
    llm_status = "not_needed"
    llm_error = None
    llm_payload: Optional[Dict[str, Any]] = None
    llm_result: Optional[Dict[str, Any]] = None

    should_use_llm = (
        bool(missing_before_llm)
        and schema
        and (enable_llm if enable_llm is not None else schema_fallback_enabled())
        and bool(os.getenv("EIA_LLM_API_KEY"))
    )
    if should_use_llm:
        llm_status = "attempted"
        llm_payload = build_llm_payload(
            schema_name,
            header_list,
            aliases,
            required_fields,
            missing_before_llm,
            sample_rows or [],
            table_title,
            context,
        )
        try:
            llm_result = call_schema_llm(llm_payload)
            accepted = validate_llm_mapping(llm_result, header_list, aliases)
            for field, header in accepted.items():
                if field not in mapping:
                    mapping[field] = header
                    sources[field] = "llm"
            llm_status = "success" if accepted else "failed"
        except Exception as exc:
            llm_status = "failed"
            llm_error = str(exc)
    elif missing_before_llm and not bool(os.getenv("EIA_LLM_API_KEY")):
        llm_status = "skipped_no_api_key"
    elif missing_before_llm:
        llm_status = "disabled"

    missing_after_llm = [field for field in required_fields if field not in mapping]
    valid = not missing_after_llm and len(set(mapping.values())) == len(mapping.values())
    validation = {
        "schema": schema_name,
        "source_table": source_table,
        "valid": valid,
        "required_fields": required_fields,
        "missing_before_llm": missing_before_llm,
        "missing_after_llm": missing_after_llm,
        "duplicate_header_assignments": len(set(mapping.values())) != len(mapping.values()),
        "llm_status": llm_status,
        "llm_error": llm_error,
    }
    if valid:
        promote_llm_candidates(config, schema_name, mapping, sources, job_id, source_table)

    result = {
        "schema": schema_name,
        "source_table": source_table,
        "table_title": table_title,
        "headers": header_list,
        "mapping": mapping,
        "mapping_sources": sources,
        "validation": validation,
        "elapsed_ms": round((time.perf_counter() - started) * 1000, 1),
    }
    write_debug_event(output_dir, "detection", result)
    if llm_payload is not None:
        write_debug_event(output_dir, "llm_input", {**llm_payload, "source_table": source_table})
    if llm_result is not None:
        write_debug_event(output_dir, "llm_output", {"source_table": source_table, "result": llm_result})
    write_debug_event(output_dir, "validation", validation)
    return result


def apply_header_mapping(
    headers: List[str],
    row: List[str],
    mapping: Dict[str, str],
) -> Dict[str, str]:
    header_index = {str(header or "").strip(): index for index, header in enumerate(headers)}
    result: Dict[str, str] = {}
    for field, header in mapping.items():
        index = header_index.get(header)
        result[field] = row[index].strip() if index is not None and index < len(row) else ""
    return result


def schema_fallback_enabled() -> bool:
    value = os.getenv("ENABLE_SCHEMA_FALLBACK", os.getenv("ENABLE_LLM_EXTRACTION", "true"))
    return str(value).lower() == "true"


def load_mapping_config(path: Path = DEFAULT_CONFIG_PATH) -> Dict[str, Any]:
    if not path.exists():
        return {"version": 1, "schemas": {}, "candidates": []}
    return json.loads(path.read_text(encoding="utf-8"))


def match_header(headers: List[str], aliases: Iterable[str]) -> Optional[str]:
    normalized = {normalize_header(header): header for header in headers}
    for alias in aliases:
        exact = normalized.get(normalize_header(alias))
        if exact:
            return exact
    for alias in aliases:
        alias_norm = normalize_header(alias)
        for header in headers:
            header_norm = normalize_header(header)
            if alias_norm and (alias_norm in header_norm or header_norm in alias_norm):
                return header
    return None


def normalize_header(value: str) -> str:
    return re.sub(r"[\s：:（）()_\-/]+", "", str(value or "")).lower()


def build_llm_payload(
    schema_name: str,
    headers: List[str],
    aliases: Dict[str, List[str]],
    required_fields: List[str],
    missing_fields: List[str],
    sample_rows: List[List[str]],
    table_title: str,
    context: str,
) -> Dict[str, Any]:
    return {
        "schema": schema_name,
        "table_title": table_title,
        "context": context[:500],
        "headers": headers,
        "sample_rows": sample_rows[:3],
        "canonical_fields": list(aliases),
        "required_fields": required_fields,
        "missing_fields": missing_fields,
    }


def call_schema_llm(payload: Dict[str, Any]) -> Dict[str, Any]:
    prompt = (
        "你是环境监测Word表格结构识别器。只解释表头含义，不提取、修改或生成任何监测数值。"
        "返回严格JSON对象，格式为 {\"mapping\": {\"canonical_field\": \"原始表头\"}}。"
        "只能使用输入中真实存在的原始表头，只映射有把握的字段。\n"
        + json.dumps(payload, ensure_ascii=False)
    )
    return chat_completion_json_object(
        [
            {"role": "system", "content": "Return strict JSON only. Never invent headers or values."},
            {"role": "user", "content": prompt},
        ],
        profile=LlmProfile.fast,
        max_retries=1,
        label="table_schema_mapping",
    )


def validate_llm_mapping(
    result: Dict[str, Any],
    headers: List[str],
    aliases: Dict[str, List[str]],
) -> Dict[str, str]:
    raw_mapping = result.get("mapping") if isinstance(result, dict) else {}
    if not isinstance(raw_mapping, dict):
        return {}
    accepted: Dict[str, str] = {}
    used_headers = set()
    for field, header in raw_mapping.items():
        field = str(field or "")
        header = str(header or "").strip()
        if field not in aliases or header not in headers or header in used_headers:
            continue
        accepted[field] = header
        used_headers.add(header)
    return accepted


def promote_llm_candidates(
    config: Dict[str, Any],
    schema_name: str,
    mapping: Dict[str, str],
    sources: Dict[str, str],
    job_id: str,
    source_table: str,
) -> None:
    changed = False
    candidates = config.setdefault("candidates", [])
    schemas = config.setdefault("schemas", {})
    aliases = schemas.setdefault(schema_name, {}).setdefault("aliases", {})
    for field, header in mapping.items():
        if sources.get(field) not in {"llm", "candidate"}:
            continue
        if header in aliases.setdefault(field, []):
            continue
        candidate = next(
            (
                item
                for item in candidates
                if item.get("schema") == schema_name
                and item.get("field") == field
                and item.get("header") == header
            ),
            None,
        )
        if candidate is None:
            candidate = {
                "schema": schema_name,
                "field": field,
                "header": header,
                "validated_jobs": [],
                "first_seen_at": datetime.now().isoformat(timespec="seconds"),
                "first_source_table": source_table,
            }
            candidates.append(candidate)
            changed = True
        validated_jobs = candidate.setdefault("validated_jobs", [])
        identity = job_id or source_table or "unknown"
        if identity not in validated_jobs:
            validated_jobs.append(identity)
            changed = True
        if len(validated_jobs) >= 2:
            aliases[field].append(header)
            candidates.remove(candidate)
            changed = True
    if changed:
        DEFAULT_CONFIG_PATH.write_text(
            json.dumps(config, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )


def write_debug_event(output_dir: Optional[Path], kind: str, payload: Dict[str, Any]) -> None:
    if output_dir is None:
        return
    debug_dir = Path(output_dir) / "debug_tables"
    debug_dir.mkdir(parents=True, exist_ok=True)
    path = debug_dir / DEBUG_FILENAMES[kind]
    items: List[Dict[str, Any]] = []
    if path.exists():
        try:
            loaded = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(loaded, list):
                items = loaded
        except Exception:
            items = []
    if kind in {"detection", "validation"}:
        schema = payload.get("schema")
        source_table = payload.get("source_table")
        items = [
            item
            for item in items
            if not (
                isinstance(item, dict)
                and item.get("schema") == schema
                and item.get("source_table") == source_table
            )
        ]
    items.append(payload)
    path.write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")


def summarize_schema_status(output_dir: Path, enabled: bool = True) -> Dict[str, Any]:
    path = Path(output_dir) / "debug_tables" / DEBUG_FILENAMES["validation"]
    detection_path = Path(output_dir) / "debug_tables" / DEBUG_FILENAMES["detection"]
    detections: List[Dict[str, Any]] = []
    if detection_path.exists():
        try:
            loaded_detections = json.loads(detection_path.read_text(encoding="utf-8"))
            if isinstance(loaded_detections, list):
                detections = loaded_detections
        except Exception:
            detections = []
    total_elapsed_ms = round(
        sum(float(item.get("elapsed_ms") or 0) for item in detections if isinstance(item, dict)),
        1,
    )
    if not path.exists():
        return {
            "enabled": enabled,
            "state": "pending",
            "label": "等待结构解析",
            "details": [],
            "review_reasons": [],
            "elapsed_ms": total_elapsed_ms,
        }
    try:
        items = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        items = []
    if not isinstance(items, list):
        items = []
    if any(not item.get("valid") for item in items):
        state = "needs_review"
        label = "需人工复核"
    elif any(item.get("llm_status") == "success" for item in items):
        state = "llm_success"
        label = "LLM结构兜底成功"
    elif any(item.get("llm_status") == "failed" for item in items):
        state = "llm_failed"
        label = "LLM结构兜底失败"
    else:
        state = "rule_success"
        label = "规则解析成功"
    return {
        "enabled": enabled,
        "state": state,
        "label": label,
        "details": items,
        "review_reasons": collect_review_reasons(items),
        "elapsed_ms": total_elapsed_ms,
    }


def collect_review_reasons(items: List[Dict[str, Any]]) -> List[str]:
    reasons: List[str] = []
    seen = set()
    for item in items:
        if not isinstance(item, dict):
            continue
        for reason in schema_item_review_reasons(item):
            if reason and reason not in seen:
                reasons.append(reason)
                seen.add(reason)
    return reasons


def schema_item_review_reasons(item: Dict[str, Any]) -> List[str]:
    schema = str(item.get("schema") or "结构解析")
    reasons: List[str] = []
    warnings = item.get("warnings")
    if isinstance(warnings, list):
        reasons.extend(str(warning) for warning in warnings if str(warning or "").strip())
    missing = item.get("missing_after_llm")
    if isinstance(missing, list) and missing:
        reasons.append(f"{schema} 缺少字段: " + "、".join(str(field) for field in missing))
    if item.get("duplicate_header_assignments"):
        reasons.append(f"{schema} 存在重复表头映射")
    llm_error = str(item.get("llm_error") or "").strip()
    if llm_error:
        reasons.append(f"{schema} LLM结构解析失败: {llm_error}")
    if item.get("valid") is False and not reasons:
        reasons.append(f"{schema} 校验未通过")
    return reasons


def record_data_validation(
    output_dir: Path,
    schema_name: str,
    *,
    valid: bool,
    warnings: Optional[List[str]] = None,
    source_table: str = "data_validation",
) -> None:
    write_debug_event(
        output_dir,
        "validation",
        {
            "schema": schema_name,
            "source_table": source_table,
            "valid": valid,
            "required_fields": [],
            "missing_before_llm": [],
            "missing_after_llm": [],
            "duplicate_header_assignments": False,
            "llm_status": "not_needed",
            "llm_error": None,
            "warnings": [str(item) for item in (warnings or []) if item],
            "validation_type": "data",
        },
    )
