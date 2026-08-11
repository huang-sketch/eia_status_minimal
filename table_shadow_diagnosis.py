import json
import os
import re
import time
from pathlib import Path
from typing import Any, Dict, List

from dotenv import load_dotenv

from llm_table_classifier import classify_table_with_llm, failed_result
from table_text_normalizer import normalize_table_chunk
from word_processor import load_docx_chunks


load_dotenv()

ENABLE_SHADOW = os.getenv("ENABLE_LLM_TABLE_CLASSIFICATION_SHADOW", "false").lower() == "true"
LOW_CONFIDENCE_TARGET_THRESHOLD = 0.72
HIGH_CONFIDENCE_TARGET_THRESHOLD = 0.85
TARGET_TABLE_TYPES = {"surface_water_plan", "surface_water_result", "noise_plan", "noise_result"}


def main() -> None:
    input_dir = Path(os.getenv("EIA_INPUT_DIR", "input"))
    output_dir = Path(os.getenv("EIA_OUTPUT_DIR", "output"))
    debug_dir = output_dir / "debug_tables"
    debug_dir.mkdir(parents=True, exist_ok=True)

    started = time.perf_counter()
    table_events: List[Dict[str, Any]] = []
    candidate_events: List[Dict[str, Any]] = []

    docx_files = sorted(path for path in input_dir.glob("*.docx") if not path.name.startswith("~$"))
    if not ENABLE_SHADOW:
        write_json(
            debug_dir / "table_llm_classification.json",
            {
                "enabled": False,
                "status": "skipped",
                "skip_reason": "ENABLE_LLM_TABLE_CLASSIFICATION_SHADOW=false",
                "input_dir": str(input_dir),
                "docx_files": [str(path) for path in docx_files],
                "tables": [],
                "elapsed_ms": elapsed_ms(started),
            },
        )
        write_json(debug_dir / "unclassified_candidate_tables.json", [])
        print("table shadow diagnosis skipped: ENABLE_LLM_TABLE_CLASSIFICATION_SHADOW=false")
        return

    if not docx_files:
        write_json(
            debug_dir / "table_llm_classification.json",
            {
                "enabled": True,
                "status": "no_input_files",
                "input_dir": str(input_dir),
                "docx_files": [],
                "tables": [],
                "elapsed_ms": elapsed_ms(started),
            },
        )
        write_json(debug_dir / "unclassified_candidate_tables.json", [])
        print(f"table shadow diagnosis found no .docx files under {input_dir}")
        return

    for path in docx_files:
        try:
            chunks = load_docx_chunks(path)
        except Exception as exc:
            table_events.append(
                {
                    "status": "failed",
                    "source_file": str(path),
                    "error": f"load_docx_chunks failed: {exc}",
                }
            )
            continue

        for chunk in chunks:
            if chunk.get("kind") != "table":
                continue
            event = classify_chunk(chunk)
            table_events.append(event)
            if should_flag_candidate(event):
                candidate_events.append(build_candidate_event(event))

    write_json(
        debug_dir / "table_llm_classification.json",
        {
            "enabled": True,
            "status": "completed",
            "input_dir": str(input_dir),
            "docx_files": [str(path) for path in docx_files],
            "tables": table_events,
            "table_count": len(table_events),
            "candidate_count": len(candidate_events),
            "elapsed_ms": elapsed_ms(started),
        },
    )
    write_json(debug_dir / "unclassified_candidate_tables.json", candidate_events)
    print(
        "table shadow diagnosis completed: "
        f"{len(table_events)} table(s), {len(candidate_events)} candidate(s)"
    )


def classify_chunk(chunk: Dict[str, Any]) -> Dict[str, Any]:
    table_payload = normalize_table_chunk(chunk)
    event: Dict[str, Any] = {
        "chunk_id": table_payload.get("chunk_id"),
        "source_file": table_payload.get("source_file"),
        "table_title": table_payload.get("table_title"),
        "normalized_table_title": table_payload.get("normalized_table_title"),
        "row_count": table_payload.get("row_count"),
        "col_counts": table_payload.get("col_counts"),
        "empty_cell_count": table_payload.get("empty_cell_count"),
        "raw_headers": table_payload.get("raw_headers"),
        "normalized_headers": table_payload.get("normalized_headers"),
        "sample_rows": table_payload.get("raw_rows"),
        "normalized_sample_rows": table_payload.get("normalized_rows"),
        "heuristics": heuristic_table_signals(table_payload),
    }
    try:
        classification = classify_table_with_llm(table_payload)
    except Exception as exc:
        classification = failed_result(exc)
    event["classification"] = classification
    event["risk"] = classify_risk(event)
    return event


def heuristic_table_signals(table_payload: Dict[str, Any]) -> Dict[str, Any]:
    text = "\n".join(
        [
            str(table_payload.get("normalized_table_title") or ""),
            str(table_payload.get("normalized_context_before") or ""),
            json.dumps(table_payload.get("normalized_headers") or [], ensure_ascii=False),
            json.dumps(table_payload.get("normalized_rows") or [], ensure_ascii=False),
        ]
    )
    signals = {
        "has_monitor_word": has_any(text, ["监测", "检测", "采样", "测量", "现状"]),
        "has_result_word": has_any(text, ["结果", "监测值", "检测值", "LAeq", "L10", "L50", "L90"]),
        "has_plan_word": has_any(text, ["方案", "布点", "点位", "断面", "执行标准", "评价标准"]),
        "has_unit": bool(re.search(r"mg/L|ug/m3|dB\(A\)|dB|%", text, re.IGNORECASE)),
        "has_date": bool(re.search(r"20\d{2}[-/.年]\d{1,2}|采样日期|监测日期|监测时间", text)),
        "has_numeric_value": bool(re.search(r"\d+(?:\.\d+)?", text)),
    }
    signals["looks_like_monitoring_table"] = (
        signals["has_monitor_word"]
        and signals["has_numeric_value"]
        and (signals["has_result_word"] or signals["has_plan_word"] or signals["has_unit"])
    )
    return signals


def classify_risk(event: Dict[str, Any]) -> str:
    classification = event.get("classification") or {}
    heuristics = event.get("heuristics") or {}
    table_type = classification.get("table_type")
    confidence = float(classification.get("confidence") or 0)
    if classification.get("status") != "success":
        return "llm_not_available_or_failed"
    if table_type in TARGET_TABLE_TYPES and confidence < LOW_CONFIDENCE_TARGET_THRESHOLD:
        return "target_table_low_confidence"
    if table_type == "unknown" and heuristics.get("looks_like_monitoring_table"):
        return "unclassified_monitoring_candidate"
    if table_type in TARGET_TABLE_TYPES and confidence >= HIGH_CONFIDENCE_TARGET_THRESHOLD:
        return "high_confidence_target_shadow_only"
    return "shadow_only"


def should_flag_candidate(event: Dict[str, Any]) -> bool:
    return event.get("risk") in {
        "target_table_low_confidence",
        "unclassified_monitoring_candidate",
        "high_confidence_target_shadow_only",
    }


def build_candidate_event(event: Dict[str, Any]) -> Dict[str, Any]:
    classification = event.get("classification") or {}
    return {
        "chunk_id": event.get("chunk_id"),
        "source_file": event.get("source_file"),
        "table_title": event.get("table_title"),
        "risk": event.get("risk"),
        "table_type": classification.get("table_type"),
        "confidence": classification.get("confidence"),
        "is_target_table": classification.get("is_target_table"),
        "reasons": classification.get("reasons") or [],
        "warnings": classification.get("warnings") or [],
        "raw_headers": event.get("raw_headers") or [],
        "sample_rows": event.get("sample_rows") or [],
        "heuristics": event.get("heuristics") or {},
    }


def has_any(text: str, needles: List[str]) -> bool:
    return any(needle in text for needle in needles)


def elapsed_ms(started: float) -> int:
    return int((time.perf_counter() - started) * 1000)


def write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
