import re
import unicodedata
from typing import Any, Dict, List


ROW_RE = re.compile(r"^row\s+\d+:\s*", re.IGNORECASE)
CELL_RE = re.compile(r"^col\s+\d+:\s*", re.IGNORECASE)


def normalize_text(value: Any) -> str:
    text = unicodedata.normalize("NFKC", str(value or ""))
    text = text.replace("（", "(").replace("）", ")")
    text = text.replace("／", "/").replace("\\", "/")
    text = text.replace("μ", "u").replace("µ", "u")
    text = re.sub(r"\s+", " ", text).strip()
    text = normalize_units(text)
    return text


def normalize_units(text: str) -> str:
    replacements = {
        "mg / L": "mg/L",
        "mg/ L": "mg/L",
        "mg /L": "mg/L",
        "ug / m3": "ug/m3",
        "ug/ m3": "ug/m3",
        "ug /m3": "ug/m3",
        "dB ( A )": "dB(A)",
        "dB(A )": "dB(A)",
        "dB( A)": "dB(A)",
    }
    for source, target in replacements.items():
        text = text.replace(source, target)
    return text


def parse_rendered_table(text: str) -> List[List[str]]:
    rows: List[List[str]] = []
    for line in str(text or "").splitlines():
        line = line.strip()
        if not line or not ROW_RE.search(line):
            continue
        row_text = ROW_RE.sub("", line).strip()
        cells = [CELL_RE.sub("", cell.strip()) for cell in row_text.split("|")]
        rows.append(cells)
    return rows


def normalize_table_chunk(chunk: Dict[str, Any], sample_size: int = 8) -> Dict[str, Any]:
    raw_rows = parse_rendered_table(chunk.get("text") or "")
    normalized_rows = [[normalize_text(cell) for cell in row] for row in raw_rows]
    raw_headers = infer_headers(raw_rows)
    normalized_headers = infer_headers(normalized_rows)
    metadata = chunk.get("metadata") or {}

    return {
        "chunk_id": chunk.get("chunk_id"),
        "source_file": chunk.get("source_file"),
        "table_title": metadata.get("table_title"),
        "context_before": metadata.get("context_before"),
        "normalized_table_title": normalize_text(metadata.get("table_title")),
        "normalized_context_before": normalize_text(metadata.get("context_before")),
        "row_count": metadata.get("row_count", len(raw_rows)),
        "col_counts": metadata.get("col_counts", [len(row) for row in raw_rows]),
        "empty_cell_count": metadata.get("empty_cell_count", count_empty_cells(raw_rows)),
        "raw_headers": raw_headers,
        "normalized_headers": normalized_headers,
        "raw_rows": raw_rows[:sample_size],
        "normalized_rows": normalized_rows[:sample_size],
        "raw_row_count": len(raw_rows),
    }


def infer_headers(rows: List[List[str]], max_header_rows: int = 1) -> List[str]:
    if not rows:
        return []
    header_rows = rows[: max(1, min(max_header_rows, len(rows)))]
    max_cols = max((len(row) for row in header_rows), default=0)
    headers: List[str] = []
    for col_index in range(max_cols):
        parts = []
        for row in header_rows:
            value = row[col_index].strip() if col_index < len(row) else ""
            if value and value not in parts:
                parts.append(value)
        headers.append("_".join(parts).strip("_"))
    return headers


def count_empty_cells(rows: List[List[str]]) -> int:
    return sum(1 for row in rows for cell in row if not str(cell or "").strip())
