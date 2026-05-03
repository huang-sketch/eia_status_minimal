import json
import re
from pathlib import Path
from typing import Any, Dict, List, Optional


NOISE_HINT_RE = re.compile(r"LAeq|L10|L50|L90|Lmax|Lmin|交通噪声|区域环境噪声", re.IGNORECASE)
DATA_HINT_RE = re.compile(r"\d{4}[-/.年]\d{1,2}|\d{1,2}:\d{2}|LAeq|L10|L50|L90|Lmax|Lmin|dB", re.IGNORECASE)
DROP_ROW_RE = re.compile(r"备注|注[:：]|标准限值|限值|单位[:：]|执行标准|评价标准")
UNIT_ONLY_RE = re.compile(r"^(dB|dB\(A\)|辆/20min|次|%|mg/L|μg/m3|ug/m3|\s*)$", re.IGNORECASE)
PARENT_HEADER_CHILDREN = {
    "区域环境噪声": {"LAeq", "L10", "L50", "L90", "Lmax", "Lmin"},
}


def export_noise_table_debug(
    chunk: Dict[str, Any],
    table_index: int,
    debug_root: Path = Path("output/debug_tables"),
) -> Optional[Dict[str, Path]]:
    """
    Stage-1 debug preprocessor only.

    It exports original and flattened noise-table JSON for human inspection.
    The current extraction pipeline does not consume these files.
    """
    if chunk.get("kind") != "table":
        return None
    if not is_noise_table(chunk.get("text") or ""):
        return None

    debug_root.mkdir(parents=True, exist_ok=True)
    parsed = preprocess_noise_table(chunk)
    original_path = debug_root / f"original_table_{table_index}.json"
    flattened_path = debug_root / f"flattened_table_{table_index}.json"

    write_json(original_path, parsed["original"])
    write_json(flattened_path, parsed["flattened"])
    return {"original": original_path, "flattened": flattened_path}


def preprocess_noise_table(chunk: Dict[str, Any]) -> Dict[str, Any]:
    rows = parse_rendered_table(chunk.get("text") or "")
    normalized_rows = normalize_width(rows)
    filled_rows = fill_merged_cells(normalized_rows)
    header_count = detect_header_region(filled_rows)
    header_rows = filled_rows[:header_count]
    data_rows = filled_rows[header_count:]
    flattened_headers = flatten_headers(header_rows)
    cleaned_data_rows = remove_non_data_rows(data_rows)
    flattened_headers, cleaned_data_rows = remove_parent_placeholder_columns(
        header_rows,
        flattened_headers,
        cleaned_data_rows,
    )

    flattened_records = [
        row_to_record(flattened_headers, row, row_index + header_count + 1)
        for row_index, row in enumerate(cleaned_data_rows)
    ]

    return {
        "original": {
            "chunk_id": chunk.get("chunk_id"),
            "source_file": chunk.get("source_file"),
            "table_title": (chunk.get("metadata") or {}).get("table_title"),
            "context_before": (chunk.get("metadata") or {}).get("context_before"),
            "row_count": len(rows),
            "column_count": max((len(row) for row in rows), default=0),
            "rows": rows,
        },
        "flattened": {
            "chunk_id": chunk.get("chunk_id"),
            "source_file": chunk.get("source_file"),
            "table_title": (chunk.get("metadata") or {}).get("table_title"),
            "header_row_count": header_count,
            "header_rows": header_rows,
            "flattened_headers": flattened_headers,
            "data_row_count": len(flattened_records),
            "records": flattened_records,
        },
    }


def is_noise_table(text: str) -> bool:
    if not NOISE_HINT_RE.search(text):
        return False
    first_lines = "\n".join(text.splitlines()[:4])
    if re.search(r"委托单位|受测单位|样品类型|报告日期|检测方法|检出限|方法名称|仪器名称|监测点名称|监测频次", first_lines):
        return False
    return bool(re.search(r"检测位置|监测位置|测量时间|监测时间|车流量|昼间|夜间|LAeq", text, re.IGNORECASE))


def parse_rendered_table(text: str) -> List[List[str]]:
    rows: List[List[str]] = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        line = re.sub(r"^row\s+\d+:\s*", "", line, flags=re.IGNORECASE)
        cells = [
            re.sub(r"^col\s+\d+:\s*", "", cell.strip(), flags=re.IGNORECASE)
            for cell in line.split("|")
        ]
        rows.append(cells)
    return rows


def normalize_width(rows: List[List[str]]) -> List[List[str]]:
    width = max((len(row) for row in rows), default=0)
    return [row + [""] * (width - len(row)) for row in rows]


def fill_merged_cells(rows: List[List[str]]) -> List[List[str]]:
    filled: List[List[str]] = []
    for row_index, row in enumerate(rows):
        new_row: List[str] = []
        for col_index, cell in enumerate(row):
            value = cell.strip()
            if not value and row_index > 0 and col_index < len(filled[row_index - 1]):
                value = filled[row_index - 1][col_index]
            if not value and col_index > 0:
                value = new_row[col_index - 1]
            new_row.append(value)
        filled.append(new_row)
    return filled


def detect_header_region(rows: List[List[str]]) -> int:
    if not rows:
        return 0

    max_header_rows = min(5, len(rows))
    for index in range(1, max_header_rows):
        if looks_like_data_row(rows[index]):
            return max(1, index)
    return min(max_header_rows, 2 if len(rows) >= 2 else 1)


def looks_like_data_row(row: List[str]) -> bool:
    text = " ".join(row)
    if re.search(r"LAeq|L10|L50|L90|Lmax|Lmin", text, re.IGNORECASE):
        return False
    numeric_cells = sum(1 for cell in row if re.search(r"\d", cell))
    has_location_or_time = bool(re.search(r"K\d+\+\d+|\d{4}[-/.年]\d{1,2}|\d{1,2}:\d{2}|N[J]?\d+", text))
    has_noise_value = bool(re.search(r"\d+(?:\.\d+)?", text) and not re.search(r"LAeq|L10|L50|L90|Lmax|Lmin", text))
    return numeric_cells >= 2 and (has_location_or_time or has_noise_value)


def flatten_headers(header_rows: List[List[str]]) -> List[str]:
    if not header_rows:
        return []
    width = max(len(row) for row in header_rows)
    headers: List[str] = []
    for col_index in range(width):
        parts: List[str] = []
        for row in header_rows:
            value = row[col_index].strip() if col_index < len(row) else ""
            value = normalize_header_cell(value)
            if value and value not in parts:
                parts.append(value)
        headers.append("_".join(parts) if parts else f"col_{col_index + 1}")
    return headers


def normalize_header_cell(value: str) -> str:
    value = re.sub(r"\s+", "", value)
    value = value.replace("检测位置", "点位").replace("监测位置", "点位")
    value = value.replace("测量时间", "监测时间")
    value = value.replace("dB（A）", "dB(A)").replace("dB(A)", "")
    return value.strip("_")


def remove_non_data_rows(rows: List[List[str]]) -> List[List[str]]:
    cleaned: List[List[str]] = []
    for row in rows:
        stripped = [cell.strip() for cell in row]
        text = " ".join(stripped)
        if not any(stripped):
            continue
        if DROP_ROW_RE.search(text):
            continue
        if all(UNIT_ONLY_RE.fullmatch(cell or "") for cell in stripped):
            continue
        if not DATA_HINT_RE.search(text) and sum(1 for cell in stripped if re.search(r"\d", cell)) < 2:
            continue
        cleaned.append(stripped)
    return cleaned


def remove_parent_placeholder_columns(
    header_rows: List[List[str]],
    headers: List[str],
    rows: List[List[str]],
) -> tuple[List[str], List[List[str]]]:
    drop_indexes = find_parent_placeholder_indexes(header_rows, headers)
    if not drop_indexes:
        return headers, rows

    kept_headers = [
        header for index, header in enumerate(headers) if index not in drop_indexes
    ]
    kept_rows = [
        [cell for index, cell in enumerate(row) if index not in drop_indexes]
        for row in rows
    ]
    return kept_headers, kept_rows


def find_parent_placeholder_indexes(
    header_rows: List[List[str]],
    headers: List[str],
) -> set[int]:
    drop_indexes: set[int] = set()
    for index, header in enumerate(headers):
        parts = [
            normalize_header_cell(row[index]).strip()
            for row in header_rows
            if index < len(row)
        ]
        parts = [part for part in parts if part]

        for parent, children in PARENT_HEADER_CHILDREN.items():
            if header != parent:
                continue
            has_parent = parent in parts
            has_child = any(child in parts for child in children)
            if has_parent and not has_child:
                drop_indexes.add(index)
    return drop_indexes


def row_to_record(headers: List[str], row: List[str], source_row_index: int) -> Dict[str, Any]:
    record = {headers[index] if index < len(headers) else f"col_{index + 1}": value for index, value in enumerate(row)}
    record["_source_row_index"] = source_row_index
    return record


def write_json(path: Path, payload: Dict[str, Any]) -> None:
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
