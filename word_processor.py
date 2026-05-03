from pathlib import Path
from typing import Any, Dict, Iterator, List, Union

from docx import Document
from docx.document import Document as DocumentObject
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph


MAX_PARAGRAPH_CHARS = 1800
Block = Union[Paragraph, Table]


def load_docx_chunks(path: Path) -> List[Dict[str, Any]]:
    """
    Read a .docx file and produce independent chunks.

    Blocks are read in Word document order. Paragraphs are grouped by size,
    while each table is kept as a standalone chunk with row/column labels.
    """
    doc = Document(str(path))
    chunks: List[Dict[str, Any]] = []

    paragraph_buffer: List[str] = []
    recent_paragraphs: List[str] = []
    paragraph_start_index = 0
    block_index = 0
    table_index = 0

    def flush_paragraphs() -> None:
        nonlocal paragraph_buffer, paragraph_start_index
        text = "\n".join(paragraph_buffer).strip()
        if not text:
            paragraph_buffer = []
            return

        chunks.append(
            {
                "chunk_id": f"{path.name}:paragraphs:{paragraph_start_index}",
                "source_file": str(path),
                "kind": "paragraphs",
                "text": text,
            }
        )
        paragraph_buffer = []

    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            if not text:
                block_index += 1
                continue

            if not paragraph_buffer:
                paragraph_start_index = block_index

            next_size = len("\n".join(paragraph_buffer + [text]))
            if paragraph_buffer and next_size > MAX_PARAGRAPH_CHARS:
                flush_paragraphs()
                paragraph_start_index = block_index

            paragraph_buffer.append(text)
            recent_paragraphs.append(text)
            recent_paragraphs[:] = recent_paragraphs[-3:]
            block_index += 1
            continue

        flush_paragraphs()

        table_text = table_to_text(block)
        if table_text.strip():
            row_count, col_counts, empty_cell_count = table_shape(block)
            chunks.append(
                {
                    "chunk_id": f"{path.name}:table:{table_index}",
                    "source_file": str(path),
                    "kind": "table",
                    "text": table_text,
                    "metadata": {
                        "table_title": recent_paragraphs[-1] if recent_paragraphs else None,
                        "context_before": "\n".join(recent_paragraphs),
                        "row_count": row_count,
                        "col_counts": col_counts,
                        "empty_cell_count": empty_cell_count,
                    },
                }
            )

        table_index += 1
        block_index += 1

    flush_paragraphs()
    return chunks


def iter_block_items(doc: DocumentObject) -> Iterator[Block]:
    for child in doc.element.body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield Table(child, doc)


def table_to_text(table: Any) -> str:
    rows: List[str] = []
    for row_index, row in enumerate(table.rows):
        cells = row_to_cells(row)
        if any(cells):
            rendered = [
                f"col {cell_index + 1}: {cell}" if cell else f"col {cell_index + 1}:"
                for cell_index, cell in enumerate(cells)
            ]
            rows.append(f"row {row_index + 1}: " + " | ".join(rendered))
    return "\n".join(rows)


def table_shape(table: Any) -> tuple[int, List[int], int]:
    col_counts: List[int] = []
    empty_cell_count = 0
    for row in table.rows:
        cells = row_to_cells(row)
        col_counts.append(len(cells))
        empty_cell_count += sum(1 for cell in cells if not cell)
    return len(table.rows), col_counts, empty_cell_count


def row_to_cells(row: Any) -> List[str]:
    cells: List[str] = []
    seen_cell_ids: set[int] = set()

    for cell in row.cells:
        cell_id = id(cell._tc)
        if cell_id in seen_cell_ids:
            cells.append("")
            continue

        seen_cell_ids.add(cell_id)
        cells.append(clean_cell(cell.text))

    return cells


def clean_cell(text: str) -> str:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    return " ".join(lines)
