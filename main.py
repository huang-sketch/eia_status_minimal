import argparse
import json
import os
import re
import time
from pathlib import Path
from typing import Any, Dict, List, Optional

from dotenv import load_dotenv

load_dotenv()

from llm_extractor import (
    call_llm,
    fallback_extract,
    normalize_and_validate,
)
from noise_table_preprocessor import (
    export_noise_table_debug,
    is_noise_table,
    preprocess_noise_table,
)
from observability import RunLogger
from word_processor import load_docx_chunks

DEFAULT_MAX_CHUNKS_PER_RUN = 20
MAX_LLM_PART_CHARS = 3400
ENABLE_LLM_EXTRACTION = os.getenv("ENABLE_LLM_EXTRACTION", "false").lower() == "true"


def max_chunks_per_run() -> int:
    return int(os.getenv("EIA_MAX_CHUNKS_PER_RUN", str(DEFAULT_MAX_CHUNKS_PER_RUN)))

HEADER_KEYWORD_RE = re.compile(
    r"点位|断面|位置|测点|采样|日期|时间|因子|项目|指标|检测|监测|结果|数值|浓度|单位|"
    r"LAeq|L10|L50|L90|Lmax|Lmin|COD|BOD5|pH|氨氮|总磷|总氮"
)
VALUE_KEYWORD_RE = re.compile(
    r"mg/L|mg/m3|mg/m³|ug/m3|μg/m3|dB\(A\)|dB|无量纲|COD|BOD5|pH|氨氮|"
    r"总磷|总氮|溶解氧|高锰酸盐指数|石油类|PM10|PM2\.5|SO2|NO2|LAeq|L10|L50|L90|Lmax|Lmin|"
    r"监测点|监测结果|检测结果|采样日期|标准限值"
)
MONITOR_TYPE_RE = re.compile(r"地表水|地下水|环境空气|声环境|噪声|土壤|废气|废水|交通噪声|区域环境噪声")
ROW_RE = re.compile(r"^row\s+\d+:\s*", re.IGNORECASE)


def analyze_docx_files(input_paths: List[Path]) -> Dict[str, Any]:
    all_records: List[Dict[str, Any]] = []
    all_chunks: List[Dict[str, Any]] = []
    extraction_chunks = 0
    processed_chunks = 0
    run_log_root = Path(os.getenv("EIA_RUN_LOG_DIR", os.getenv("EIA_OUTPUT_DIR", "output"))) / "extraction_logs"
    logger = RunLogger(root=run_log_root)
    debug_table_index = 0

    for path in input_paths:
        chunks = load_docx_chunks(path)

        for chunk in chunks:
            if processed_chunks >= max_chunks_per_run():
                print(f"Reached chunk limit: {max_chunks_per_run()}")
                break

            processed_chunks += 1
            all_chunks.append(chunk)

            chunk_type = classify_chunk(chunk)
            export_debug_chunk(chunk, processed_chunks)
            if chunk.get("kind") == "table":
                exported = export_noise_table_debug(chunk, debug_table_index)
                if exported:
                    debug_table_index += 1
                    logger.log_warning(
                        chunk,
                        "noise_table_preprocessor_debug_exported",
                        {
                            "original": str(exported["original"]),
                            "flattened": str(exported["flattened"]),
                        },
                    )
            print(
                f"Processing chunk {processed_chunks}/{max_chunks_per_run()}: "
                f"{chunk['chunk_id']} [{chunk_type}]",
                flush=True,
            )

            chunk_records: List[Dict[str, Any]] = []
            if chunk_type == "simple_table":
                logger.log_chunk(
                    chunk,
                    {
                        "chunk_type": chunk_type,
                        "method": "fallback",
                        "split": False,
                        "part_count": 1,
                    },
                )
                started = time.time()
                chunk_records = extract_with_fallback(chunk)
                logger.log_chunk(
                    chunk,
                    {
                        "chunk_type": chunk_type,
                        "method": "fallback",
                        "records": len(chunk_records),
                        "elapsed_ms": elapsed_ms(started),
                    },
                )
                print(f"Chunk result: {len(chunk_records)} record(s)", flush=True)

            elif chunk_type == "complex_table":
                if is_noise_table(chunk.get("text") or ""):
                    try:
                        chunk_records, part_count = extract_complex_noise_table_new_flow(
                            chunk,
                            logger,
                        )
                        extraction_chunks += part_count
                    except Exception as exc:
                        logger.log_warning(
                            chunk,
                            "complex_noise_table_new_flow_failed_use_old_flow",
                            {
                                "error": str(exc),
                                "llm_enabled": ENABLE_LLM_EXTRACTION,
                            },
                        )
                        if ENABLE_LLM_EXTRACTION:
                            chunk_records, part_count = extract_complex_table_old_flow(
                                chunk,
                                logger,
                            )
                            extraction_chunks += part_count
                        else:
                            chunk_records = extract_with_fallback(chunk)
                            extraction_chunks += 1
                else:
                    if ENABLE_LLM_EXTRACTION:
                        chunk_records, part_count = extract_complex_table_old_flow(
                            chunk,
                            logger,
                        )
                        extraction_chunks += part_count
                    else:
                        logger.log_chunk(
                            chunk,
                            {
                                "chunk_type": chunk_type,
                                "method": "fallback",
                                "llm_enabled": False,
                                "message": "table processed by rules/flattened pipeline",
                            },
                        )
                        print("LLM extraction disabled; table processed by rules", flush=True)
                        chunk_records = extract_with_fallback(chunk)
                        extraction_chunks += 1

            else:
                logger.log_chunk(
                    chunk,
                    {
                        "chunk_type": chunk_type,
                        "method": "skip",
                        "split": False,
                        "part_count": 0,
                        "records": 0,
                        "message": "normal_text skipped: not a data source",
                    },
                )
                print("normal_text skipped: not a data source", flush=True)

            if chunk_type == "simple_table":
                extraction_chunks += 1
            print(f"Chunk total: {len(chunk_records)} record(s)", flush=True)
            all_records.extend(chunk_records)

        if processed_chunks >= max_chunks_per_run():
            break

    deduped_records = normalize_evidence_objects(dedupe_output_records(all_records))
    return {
        "records": deduped_records,
        "meta": {
            "input_files": [str(path) for path in input_paths],
            "chunk_count": len(all_chunks),
            "extraction_chunk_count": extraction_chunks,
            "record_count": len(deduped_records),
            "run_dir": str(logger.run_dir.resolve()),
        },
    }


def classify_chunk(chunk: Dict[str, Any]) -> str:
    if chunk.get("kind") != "table":
        return "normal_text"

    text = chunk.get("text") or ""
    rows = table_rows(text)
    if not rows:
        return "normal_text"

    col_counts = [len(split_table_line(row)) for row in rows]
    stable_columns = len(set(col_counts)) <= 1
    max_cols = max(col_counts) if col_counts else 0
    header_text = "\n".join(rows[:3])
    header_hits = len(HEADER_KEYWORD_RE.findall(header_text))
    value_hits = len(VALUE_KEYWORD_RE.findall(text))
    monitor_types = set(MONITOR_TYPE_RE.findall(text))
    empty_ratio = table_empty_ratio(rows)
    multi_header_sign = has_multi_header(rows)
    mixed_monitor_types = len(monitor_types) >= 2

    if looks_like_non_result_table(header_text):
        return "normal_text"

    if value_hits == 0 and header_hits < 3:
        return "normal_text"

    if (
        stable_columns
        and max_cols <= 7
        and header_hits >= 3
        and not multi_header_sign
        and not mixed_monitor_types
    ):
        return "simple_table"

    if (
        max_cols >= 8
        or empty_ratio > 0.12
        or multi_header_sign
        or mixed_monitor_types
        or len(text) > MAX_LLM_PART_CHARS
    ):
        return "complex_table"

    if header_hits >= 2 and stable_columns:
        return "simple_table"
    return "complex_table"


def extract_with_fallback(chunk: Dict[str, Any]) -> List[Dict[str, Any]]:
    raw_records = fallback_extract(chunk.get("text") or "")
    for record in raw_records:
        record["extraction_method"] = "rule"
    return normalize_and_validate(raw_records, chunk, default_method="rule")


def extract_part_with_llm(
    part: Dict[str, Any],
    logger: RunLogger,
    fallback_on_error: bool,
) -> List[Dict[str, Any]]:
    started = time.time()
    logger.log_part(
        part,
        {
            "chunk_type": part.get("chunk_type"),
            "method": "llm",
            "part_index": part.get("part_index"),
            "part_count": part.get("part_count"),
            "part_chars": len(part.get("text") or ""),
            "carry": part.get("carry"),
        },
    )

    try:
        raw_records = call_llm(part.get("text") or "", chunk=part, logger=logger)
        for record in raw_records:
            record["extraction_method"] = "llm"
        records = normalize_and_validate(raw_records, part, default_method="llm")
        logger.log_part(
            part,
            {
                "chunk_type": part.get("chunk_type"),
                "method": "llm",
                "part_chars": len(part.get("text") or ""),
                "elapsed_ms": elapsed_ms(started),
                "records": len(records),
            },
        )
        return records
    except Exception as exc:
        part["_llm_failed"] = True
        logger.log_failed_chunk(part, "llm_extract", exc, fallback_used=fallback_on_error)
        if not fallback_on_error:
            return []

        fallback_records = extract_with_fallback(part)
        logger.log_warning(
            part,
            "llm_failed_used_fallback",
            {
                "error": str(exc),
                "records": len(fallback_records),
                "elapsed_ms": elapsed_ms(started),
            },
        )
        return fallback_records


def extract_complex_noise_table_new_flow(
    chunk: Dict[str, Any],
    logger: RunLogger,
) -> tuple[List[Dict[str, Any]], int]:
    parsed = preprocess_noise_table(chunk)
    flattened = parsed["flattened"]
    rows = flattened.get("records") or []
    headers = flattened.get("flattened_headers") or []
    if not rows or not headers:
        raise ValueError("flattened noise table has no rows or headers")

    parts = build_flattened_noise_parts(chunk, flattened)
    print(f"Use flattened noise-table flow: {len(parts)} part(s)", flush=True)

    all_records: List[Dict[str, Any]] = []
    for part_index, part in enumerate(parts, start=1):
        print(
            f"  Processing flattened part {part_index}/{len(parts)}: "
            f"{part['chunk_id']} ({len(part['text'])} chars)",
            flush=True,
        )
        started = time.time()
        records = extract_flattened_noise_part(part)
        logger.log_part(
            part,
            {
                "chunk_type": "complex_noise_table",
                "method": "flattened_rule",
                "part_index": part.get("part_index"),
                "part_count": part.get("part_count"),
                "part_chars": len(part.get("text") or ""),
                "elapsed_ms": elapsed_ms(started),
                "records": len(records),
            },
        )
        print(f"  Flattened part result: {len(records)} record(s)", flush=True)
        all_records.extend(records)

    logger.log_chunk(
        chunk,
        {
            "chunk_type": "complex_noise_table",
            "method": "flattened_rule",
            "split": True,
            "part_count": len(parts),
            "records": len(all_records),
            "message": "table processed by rules/flattened pipeline",
        },
    )
    return all_records, len(parts)


def extract_complex_table_old_flow(
    chunk: Dict[str, Any],
    logger: RunLogger,
) -> tuple[List[Dict[str, Any]], int]:
    parts = split_complex_table(chunk)
    print(f"Split complex table into {len(parts)} part(s)", flush=True)
    chunk_records: List[Dict[str, Any]] = []
    llm_failures = 0

    for part_index, part in enumerate(parts, start=1):
        print(
            f"  Processing part {part_index}/{len(parts)}: "
            f"{part['chunk_id']} ({len(part['text'])} chars)",
            flush=True,
        )

        if llm_failures:
            logger.log_warning(
                part,
                "complex_table_llm_circuit_breaker_used_fallback",
                {"previous_llm_failures": llm_failures},
            )
            records = extract_with_fallback(part)
        else:
            records = extract_part_with_llm(part, logger, fallback_on_error=True)
            if part.get("_llm_failed"):
                llm_failures += 1

        print(f"  Part result: {len(records)} record(s)", flush=True)
        chunk_records.extend(records)

    logger.log_chunk(
        chunk,
        {
            "chunk_type": "complex_table",
            "method": "old_llm_parts",
            "split": True,
            "part_count": len(parts),
            "records": len(chunk_records),
        },
    )
    return chunk_records, len(parts)


def build_flattened_noise_parts(
    chunk: Dict[str, Any],
    flattened: Dict[str, Any],
    max_rows: int = 3,
) -> List[Dict[str, Any]]:
    headers = flattened.get("flattened_headers") or []
    rows = flattened.get("records") or []
    title = flattened.get("table_title") or ""

    parts: List[Dict[str, Any]] = []
    current_rows: List[Dict[str, Any]] = []
    current_carry: Dict[str, Any] = {}

    for row in rows:
        row = fill_flattened_noise_row(row, current_carry)
        candidate = render_flattened_noise_part(title, headers, current_rows + [row])
        if current_rows and (len(current_rows) >= max_rows or len(candidate) > MAX_LLM_PART_CHARS):
            parts.append(build_flattened_noise_part_chunk(chunk, title, headers, current_rows, len(parts) + 1))
            current_rows = []
        current_rows.append(row)

    if current_rows:
        parts.append(build_flattened_noise_part_chunk(chunk, title, headers, current_rows, len(parts) + 1))

    part_count = len(parts)
    for part in parts:
        part["part_count"] = part_count
    return parts


def fill_flattened_noise_row(row: Dict[str, Any], carry: Dict[str, Any]) -> Dict[str, Any]:
    filled = dict(row)
    for field in ("点位", "监测时间"):
        if filled.get(field):
            carry[field] = filled[field]
        elif carry.get(field):
            filled[field] = carry[field]
    return filled


def build_flattened_noise_part_chunk(
    chunk: Dict[str, Any],
    title: str,
    headers: List[str],
    rows: List[Dict[str, Any]],
    part_index: int,
) -> Dict[str, Any]:
    part = dict(chunk)
    part["chunk_id"] = f"{chunk['chunk_id']}:flattened_part:{part_index}"
    part["parent_chunk_id"] = chunk["chunk_id"]
    part["part_index"] = part_index
    part["chunk_type"] = "complex_noise_table"
    part["llm_model"] = os.getenv("EIA_LLM_FAST_MODEL", "qwen-flash")
    part["llm_max_tokens"] = int(os.getenv("EIA_LLM_FAST_MAX_TOKENS", "4096"))
    part["table_title"] = title
    part["flattened_headers"] = headers
    part["flattened_rows"] = rows
    part["text"] = render_flattened_noise_part(title, headers, rows)
    return part


def render_flattened_noise_part(
    title: str,
    headers: List[str],
    rows: List[Dict[str, Any]],
) -> str:
    clean_rows = [
        {key: row.get(key) for key in headers if key in row}
        for row in rows
    ]
    return "\n".join(
        [
            "FLATTENED_NOISE_TABLE",
            f"title: {title}",
            "headers:",
            json.dumps(headers, ensure_ascii=False),
            "rows:",
            json.dumps(clean_rows, ensure_ascii=False),
            "",
            "Extract one record for each numeric noise metric value. Use the flattened header as factor.",
            "For example, 区域环境噪声_LAeq means monitor_type=acoustic, factor=LAeq.",
            "Use 点位 as point and 监测时间 as sample_date. Return [] if no measurement values.",
        ]
    )


def extract_flattened_noise_part(part: Dict[str, Any]) -> List[Dict[str, Any]]:
    headers = part.get("flattened_headers") or []
    rows = part.get("flattened_rows") or []
    table_title = str(part.get("table_title") or "")
    raw_records: List[Dict[str, Any]] = []

    for row in rows:
        point = row.get("点位")
        sample_date = row.get("监测时间")
        evidence = {key: row.get(key) for key in headers if key in row}
        for header in headers:
            factor = noise_factor_from_header(header)
            if not factor:
                continue
            noise_type, noise_type_label = noise_type_from_header(header, table_title)
            value = str(row.get(header) or "").strip()
            if not value or not re.search(r"\d", value):
                continue
            raw_records.append(
                {
                    "source_type": "监测报告",
                    "monitor_type": "acoustic",
                    "noise_type": noise_type,
                    "noise_type_label": noise_type_label,
                    "point": point,
                    "sample_date": sample_date,
                    "factor": factor,
                    "value": value,
                    "unit": "dB(A)",
                    "standard_class": None,
                    "evidence": evidence,
                    "confidence": 0.95,
                    "extraction_method": "rule",
                }
            )

    return normalize_and_validate(raw_records, part, default_method="rule")


def noise_factor_from_header(header: str) -> Optional[str]:
    match = re.search(r"(LAeq|L10|L50|L90|Lmax|Lmin)", header, re.IGNORECASE)
    return match.group(1) if match else None


def noise_type_from_header(header: str, table_title: str = "") -> tuple[Optional[str], Optional[str]]:
    text = "\n".join([str(header or ""), str(table_title or "")])
    if "交通噪声" in text:
        return "traffic_noise", "交通噪声"
    if "区域环境噪声" in text:
        return "area_environment_noise", "区域环境噪声"
    return None, None


def split_complex_table(chunk: Dict[str, Any]) -> List[Dict[str, Any]]:
    rows = table_rows(chunk.get("text") or "")
    if not rows:
        return [chunk]

    context = build_table_context(chunk, rows)
    data_rows = rows[context["header_row_count"] :]
    if not data_rows:
        data_rows = rows

    parts: List[Dict[str, Any]] = []
    current_rows: List[str] = []
    carry: Dict[str, Optional[str]] = {}
    current_carry = ""

    for row in data_rows:
        row_carry = infer_carry_from_row(row, carry)
        rendered_candidate = render_table_part(context, current_rows + [row], current_carry)
        if current_rows and len(rendered_candidate) > MAX_LLM_PART_CHARS:
            parts.append(build_part_chunk(chunk, context, current_rows, len(parts) + 1, current_carry))
            current_rows = []
            current_carry = render_carry(carry)

        current_rows.append(row)
        carry.update({key: value for key, value in row_carry.items() if value})
        if not current_carry:
            current_carry = render_carry(carry)

    if current_rows:
        parts.append(build_part_chunk(chunk, context, current_rows, len(parts) + 1, current_carry))

    part_count = len(parts)
    for part in parts:
        part["part_count"] = part_count
    return parts or [chunk]


def build_table_context(chunk: Dict[str, Any], rows: List[str]) -> Dict[str, Any]:
    metadata = chunk.get("metadata") or {}
    header_row_count = guess_header_row_count(rows)
    header_rows = rows[:header_row_count]
    context_before = metadata.get("context_before") or ""
    title = metadata.get("table_title") or infer_title(context_before)
    monitor_type = infer_monitor_type("\n".join([context_before, "\n".join(rows[:5])]))

    return {
        "title": title,
        "context_before": context_before,
        "header_rows": header_rows,
        "header_row_count": header_row_count,
        "monitor_type": monitor_type,
        "unit_note": infer_unit_note(rows[:5]),
    }


def build_part_chunk(
    chunk: Dict[str, Any],
    context: Dict[str, Any],
    rows: List[str],
    part_index: int,
    carry_text: str,
) -> Dict[str, Any]:
    part = dict(chunk)
    part["chunk_id"] = f"{chunk['chunk_id']}:part:{part_index}"
    part["parent_chunk_id"] = chunk["chunk_id"]
    part["part_index"] = part_index
    part["chunk_type"] = "complex_table"
    part["carry"] = carry_text
    part["text"] = render_table_part(context, rows, carry_text)
    return part


def render_table_part(context: Dict[str, Any], rows: List[str], carry_text: str) -> str:
    blocks = [
        "TABLE_CONTEXT",
        f"title: {context.get('title') or ''}",
        f"monitor_type_hint: {context.get('monitor_type') or ''}",
        f"unit_note: {context.get('unit_note') or ''}",
        "header:",
        "\n".join(context.get("header_rows") or []),
    ]
    if carry_text:
        blocks.extend(["carry_info:", carry_text])
    blocks.extend(["data_rows:", "\n".join(rows)])
    return "\n".join(block for block in blocks if block is not None).strip()


def table_rows(text: str) -> List[str]:
    return [line.strip() for line in text.splitlines() if ROW_RE.search(line.strip())]


def split_table_line(line: str) -> List[str]:
    line = ROW_RE.sub("", line).strip()
    return [re.sub(r"^col\s+\d+:\s*", "", cell.strip(), flags=re.IGNORECASE) for cell in line.split("|")]


def table_empty_ratio(rows: List[str]) -> float:
    cells = [cell for row in rows for cell in split_table_line(row)]
    if not cells:
        return 0.0
    return sum(1 for cell in cells if not cell) / len(cells)


def has_multi_header(rows: List[str]) -> bool:
    first_rows = rows[:3]
    if len(first_rows) < 2:
        return False
    empty_counts = [sum(1 for cell in split_table_line(row) if not cell) for row in first_rows]
    return max(empty_counts or [0]) >= 2


def looks_like_non_result_table(header_text: str) -> bool:
    return bool(
        re.search(r"仪器名称|仪器型号|设备编号|检测方法|分析方法|检出限|方法名称|标准号|依据标准", header_text)
    )


def guess_header_row_count(rows: List[str]) -> int:
    count = 0
    for row in rows[:4]:
        if HEADER_KEYWORD_RE.search(row) or sum(1 for cell in split_table_line(row) if not cell) >= 2:
            count += 1
            continue
        break
    return max(1, count)


def infer_title(context_before: str) -> Optional[str]:
    lines = [line.strip() for line in context_before.splitlines() if line.strip()]
    return lines[-1] if lines else None


def infer_monitor_type(text: str) -> Optional[str]:
    match = MONITOR_TYPE_RE.search(text)
    return match.group(0) if match else None


def infer_unit_note(rows: List[str]) -> Optional[str]:
    text = "\n".join(rows)
    units = sorted(set(re.findall(r"mg/L|mg/m3|mg/m³|ug/m3|μg/m3|dB\(A\)|dB|无量纲|%", text)))
    return ", ".join(units) if units else None


def infer_carry_from_row(row: str, carry: Dict[str, Optional[str]]) -> Dict[str, Optional[str]]:
    cells = split_table_line(row)
    result: Dict[str, Optional[str]] = {}
    if cells:
        first = cells[0].strip()
        if first:
            result["point"] = first
    if len(cells) > 1:
        second = cells[1].strip()
        if re.search(r"20\d{2}[-/.年]\d{1,2}|^\d{1,2}月\d{1,2}", second):
            result["sample_date"] = second
    return result or carry


def render_carry(carry: Dict[str, Optional[str]]) -> str:
    return "; ".join(f"{key}: {value}" for key, value in carry.items() if value)


def normal_text_needs_llm(text: str) -> bool:
    return bool(VALUE_KEYWORD_RE.search(text) and re.search(r"\d", text))


def dedupe_output_records(records: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    seen: set[str] = set()
    result: List[Dict[str, Any]] = []
    for record in records:
        key = "|".join(
            str(record.get(field) or "").strip()
            for field in ("source_file", "point", "sample_date", "factor", "value", "unit", "evidence")
        )
        if key in seen:
            continue
        seen.add(key)
        result.append(record)
    return result


def normalize_evidence_objects(records: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    normalized: List[Dict[str, Any]] = []
    for record in records:
        item = dict(record)
        item["evidence"] = normalize_evidence_value(item.get("evidence"))
        normalized.append(item)
    return normalized


def normalize_evidence_value(evidence: Any) -> Any:
    if isinstance(evidence, (dict, list)):
        return evidence
    text = str(evidence or "").strip()
    if not text:
        return {}
    if text[:1] in {"{", "["}:
        try:
            parsed = json.loads(text)
            if isinstance(parsed, (dict, list)):
                return parsed
        except json.JSONDecodeError:
            pass
    return {"text": text}


def elapsed_ms(started: float) -> int:
    return int((time.time() - started) * 1000)


def export_debug_chunk(chunk: Dict[str, Any], index: int) -> None:
    debug_root = Path(os.getenv("EIA_OUTPUT_DIR", "output")) / "debug_chunks"
    debug_root.mkdir(parents=True, exist_ok=True)
    safe_id = re.sub(r"[^0-9A-Za-z_.-]+", "_", chunk.get("chunk_id") or f"chunk_{index}")
    payload = {
        "chunk_id": chunk.get("chunk_id"),
        "source_file": chunk.get("source_file"),
        "kind": chunk.get("kind"),
        "char_count": len(chunk.get("text") or ""),
        "metadata": chunk.get("metadata") or {},
        "text": chunk.get("text") or "",
    }
    path = debug_root / f"{index:03d}_{safe_id}.json"
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Minimal EIA monitoring status extractor for .docx files."
    )
    parser.add_argument(
        "inputs",
        nargs="+",
        help="One or more .docx files, for example: report.docx plan.docx",
    )
    parser.add_argument(
        "-o",
        "--output",
        default="output/eia_result_test.json",
        help="Output JSON path. Default: output/eia_result_test.json",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if not ENABLE_LLM_EXTRACTION:
        os.environ["EIA_LLM_DISABLE"] = "1"
    else:
        os.environ.pop("EIA_LLM_DISABLE", None)
        if not os.getenv("EIA_LLM_API_KEY"):
            raise RuntimeError("EIA_LLM_API_KEY is required when ENABLE_LLM_EXTRACTION=true.")

    print(f"ENABLE_LLM_EXTRACTION: {ENABLE_LLM_EXTRACTION}", flush=True)
    if not ENABLE_LLM_EXTRACTION:
        print("LLM extraction disabled", flush=True)
    else:
        print("LLM extraction enabled", flush=True)

    input_paths = [Path(item).resolve() for item in args.inputs]

    missing = [str(path) for path in input_paths if not path.exists()]
    if missing:
        raise FileNotFoundError(f"Input file(s) not found: {', '.join(missing)}")

    result = analyze_docx_files(input_paths)
    output_path = Path(args.output).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(
        json.dumps(result, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    monitoring_records_path = output_path.parent / "monitoring_records.json"
    monitoring_records_path.write_text(
        json.dumps(result["records"], ensure_ascii=False),
        encoding="utf-8",
    )
    print(f"Wrote {result['meta']['record_count']} records to {output_path}")
    print(f"Wrote monitoring records to {monitoring_records_path}")


if __name__ == "__main__":
    main()
