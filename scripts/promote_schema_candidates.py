import argparse
import json
import os
import re
from copy import deepcopy
from pathlib import Path
from typing import Any, Dict, List, Tuple
from uuid import uuid4


BASE_DIR = Path(__file__).resolve().parents[1]
DEFAULT_CONFIG_PATH = BASE_DIR / "config" / "table_schema_mappings.json"


def normalize_header(value: Any) -> str:
    return re.sub(r"[\s\uFF0C\u3002\uFF08\uFF09()_\-/]+", "", str(value or "")).lower()


def load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def review_candidates(
    config: Dict[str, Any],
    candidates: List[Dict[str, Any]],
) -> Tuple[Dict[str, Any], Dict[str, Any]]:
    updated = deepcopy(config)
    schemas = updated.get("schemas")
    if not isinstance(schemas, dict):
        raise ValueError("mapping config must contain a schemas object")

    accepted: List[Dict[str, str]] = []
    skipped_existing: List[Dict[str, str]] = []
    rejected: List[Dict[str, str]] = []
    seen = set()

    for raw in candidates:
        if not isinstance(raw, dict):
            rejected.append({"reason": "candidate must be an object"})
            continue
        schema_name = str(raw.get("schema") or "").strip()
        field = str(raw.get("field") or "").strip()
        header = str(raw.get("header") or "").strip()
        identity = (schema_name, field, normalize_header(header))
        item = {"schema": schema_name, "field": field, "header": header}

        if raw.get("validation_status") != "accepted":
            rejected.append({**item, "reason": "candidate was not accepted at runtime"})
            continue
        if identity in seen:
            rejected.append({**item, "reason": "duplicate candidate"})
            continue
        seen.add(identity)

        schema = schemas.get(schema_name)
        if not isinstance(schema, dict):
            rejected.append({**item, "reason": "unknown schema"})
            continue
        aliases = schema.get("aliases")
        if not isinstance(aliases, dict) or field not in aliases:
            rejected.append({**item, "reason": "unknown canonical field"})
            continue
        if not header:
            rejected.append({**item, "reason": "empty header"})
            continue
        source_headers = raw.get("source_headers")
        if not isinstance(source_headers, list) or header not in source_headers:
            rejected.append({**item, "reason": "header is absent from source_headers"})
            continue

        normalized = normalize_header(header)
        conflict_field = ""
        for other_field, field_aliases in aliases.items():
            if other_field == field or not isinstance(field_aliases, list):
                continue
            if normalized in {normalize_header(alias) for alias in field_aliases}:
                conflict_field = str(other_field)
                break
        if conflict_field:
            rejected.append(
                {**item, "reason": f"header already belongs to field {conflict_field}"}
            )
            continue

        field_aliases = aliases[field]
        if not isinstance(field_aliases, list):
            rejected.append({**item, "reason": "field aliases must be a list"})
            continue
        if normalized in {normalize_header(alias) for alias in field_aliases}:
            skipped_existing.append(item)
            continue
        field_aliases.append(header)
        accepted.append(item)

    summary = {
        "accepted": accepted,
        "skipped_existing": skipped_existing,
        "rejected": rejected,
        "can_apply": not rejected,
    }
    return updated, summary


def write_json_atomic(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_name(f".{path.name}.{uuid4().hex}.tmp")
    try:
        temp_path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        os.replace(temp_path, path)
    finally:
        if temp_path.exists():
            temp_path.unlink()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Review and explicitly promote validated table schema candidates."
    )
    parser.add_argument("--input", required=True, type=Path, help="Candidate JSON file")
    parser.add_argument(
        "--config",
        type=Path,
        default=DEFAULT_CONFIG_PATH,
        help="Mapping configuration to review or update",
    )
    parser.add_argument(
        "--apply",
        action="store_true",
        help="Write accepted candidates to the mapping configuration",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    config = load_json(args.config)
    candidates = load_json(args.input)
    if not isinstance(config, dict):
        raise ValueError("mapping config must be a JSON object")
    if not isinstance(candidates, list):
        raise ValueError("candidate file must be a JSON array")

    updated, summary = review_candidates(config, candidates)
    summary["applied"] = False
    if args.apply:
        if not summary["can_apply"]:
            print(json.dumps(summary, ensure_ascii=False, indent=2))
            return 2
        write_json_atomic(args.config, updated)
        summary["applied"] = True
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
