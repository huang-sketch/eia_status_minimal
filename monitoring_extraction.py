import os
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

if not os.getenv("EIA_MAX_CHUNKS_PER_RUN"):
    os.environ["EIA_MAX_CHUNKS_PER_RUN"] = "100"

SCHEMA_FALLBACK_REQUESTED = os.getenv("ENABLE_SCHEMA_FALLBACK", os.getenv("ENABLE_LLM_EXTRACTION", "true")).lower() == "true"
# main.py reads this flag at import time. Keep its generic extractor rule-only;
# schema fallback is handled later by the dedicated table-schema mapper.
os.environ["ENABLE_LLM_EXTRACTION"] = "false"
os.environ["EIA_LLM_DISABLE"] = "1"

from main import analyze_docx_files
from surface_water_pipeline import find_docx, write_json

EXTRACTION_DIRNAME = "extraction"


def main() -> None:
    input_dir = Path(os.getenv("EIA_INPUT_DIR", "input"))
    output_dir = Path(os.getenv("EIA_OUTPUT_DIR", "output"))
    extraction_dir = output_dir / EXTRACTION_DIRNAME
    extraction_dir.mkdir(parents=True, exist_ok=True)

    configure_llm_extraction()

    plan_path = find_docx(input_dir, "方案")
    report_path = find_docx(input_dir, "报告")
    result = analyze_docx_files([plan_path, report_path])

    write_json(extraction_dir / "eia_result.json", result)
    write_json(extraction_dir / "records.json", result.get("records") or [])
    write_json(
        extraction_dir / "meta.json",
        {
            "record_count": result.get("meta", {}).get("record_count", 0),
            "chunk_count": result.get("meta", {}).get("chunk_count", 0),
            "extraction_chunk_count": result.get("meta", {}).get("extraction_chunk_count", 0),
            "run_dir": result.get("meta", {}).get("run_dir"),
            "input_files": result.get("meta", {}).get("input_files", []),
        },
    )

    print(f"extraction records: {len(result.get('records') or [])}")
    print(f"extraction meta: {extraction_dir / 'meta.json'}")


def configure_llm_extraction() -> None:
    enabled = SCHEMA_FALLBACK_REQUESTED
    os.environ["ENABLE_SCHEMA_FALLBACK"] = "true" if enabled else "false"
    # Generic extraction remains deterministic. LLM is reserved for the
    # table-schema mapper, which only interprets headers and never copies values.
    os.environ["EIA_LLM_DISABLE"] = "1"
    if enabled:
        if not os.getenv("EIA_LLM_API_KEY"):
            print("Smart schema fallback enabled, but no API key; rule parsing only", flush=True)
            return
        print("Smart schema fallback enabled; generic extraction remains rule-only", flush=True)
        return

    print("Smart schema fallback disabled; rule extraction only", flush=True)


if __name__ == "__main__":
    main()
