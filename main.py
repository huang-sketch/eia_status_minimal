import argparse
import json
from pathlib import Path
from typing import Any, Dict, List

from llm_extractor import extract_records_from_chunk
from word_processor import load_docx_chunks


def analyze_docx_files(input_paths: List[Path]) -> Dict[str, Any]:
    all_records: List[Dict[str, Any]] = []
    all_chunks: List[Dict[str, Any]] = []

    for path in input_paths:
        chunks = load_docx_chunks(path)
        all_chunks.extend(chunks)

        for chunk in chunks:
            records = extract_records_from_chunk(chunk)
            all_records.extend(records)

    return {
        "records": all_records,
        "meta": {
            "input_files": [str(path) for path in input_paths],
            "chunk_count": len(all_chunks),
            "record_count": len(all_records),
        },
    }


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
        default="eia_result.json",
        help="Output JSON path. Default: eia_result.json",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_paths = [Path(item).resolve() for item in args.inputs]

    missing = [str(path) for path in input_paths if not path.exists()]
    if missing:
        raise FileNotFoundError(f"Input file(s) not found: {', '.join(missing)}")

    result = analyze_docx_files(input_paths)
    output_path = Path(args.output).resolve()
    output_path.write_text(
        json.dumps(result, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"Wrote {result['meta']['record_count']} records to {output_path}")


if __name__ == "__main__":
    main()
