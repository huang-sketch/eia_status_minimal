import json
import os
from pathlib import Path
from typing import Any, Dict, Iterable, List


BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = Path(os.getenv("EIA_OUTPUT_DIR", "output"))
DEBUG_DIR = OUTPUT_DIR / "debug_tables"

MOJIBAKE_MARKERS = ["锛", "鐩", "缂", "妫", "绉", "鏍", "鑸", "閾", "杞", "澹", "鎸", "楂", "鐐"]

DEFAULT_SCAN_PATHS = [
    "config/table_schema_mappings.json",
    "table_schema_mapper.py",
    "table_text_normalizer.py",
    "eia_document_router.py",
    "llm_table_classifier.py",
    "table_shadow_diagnosis.py",
    ".codex/skills/highway-eia-status-writing/SKILL.md",
    ".codex/skills/highway-eia-status-writing/references/writing-rules.md",
    ".codex/skills/eia-document-router/SKILL.md",
    ".codex/skills/eia-document-router/references/router-rules.md",
]


def main() -> None:
    result = build_encoding_health_report(BASE_DIR, DEFAULT_SCAN_PATHS)
    DEBUG_DIR.mkdir(parents=True, exist_ok=True)
    write_json(DEBUG_DIR / "encoding_health_check.json", result)
    print(f"encoding health: {result['status']}; issues={result['issue_count']}")


def build_encoding_health_report(base_dir: Path, relative_paths: Iterable[str]) -> Dict[str, Any]:
    issues: List[Dict[str, Any]] = []
    checked_files: List[str] = []
    for relative_path in relative_paths:
        path = base_dir / relative_path
        if not path.exists():
            continue
        checked_files.append(relative_path)
        text = path.read_text(encoding="utf-8", errors="replace")
        for line_number, line in enumerate(text.splitlines(), start=1):
            markers = [marker for marker in MOJIBAKE_MARKERS if marker in line]
            if not markers:
                continue
            issues.append(
                {
                    "file": relative_path,
                    "line": line_number,
                    "markers": markers,
                    "text": line[:200],
                }
            )
    return {
        "status": "ok" if not issues else "needs_review",
        "encoding": "utf-8",
        "marker_set": MOJIBAKE_MARKERS,
        "checked_files": checked_files,
        "issue_count": len(issues),
        "issues": issues,
    }


def write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
