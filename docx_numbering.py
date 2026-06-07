"""Sequential chapter and table numbering for EIA section docx files."""

import json
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Optional


STATE_FILENAME = "docx_numbering_state.json"
DEFAULT_NUMBERING_CONFIG = {
    "numbered_sections": ["overview", "noise", "surface_water"],
    "generation_order": ["overview", "noise", "surface_water"],
}

SECTION_DOCX_FILENAMES = {
    "overview": "project_area_overview.docx",
    "noise": "noise_section.docx",
    "surface_water": "surface_water_section.docx",
}


def generation_order() -> List[str]:
    config = load_numbering_config()
    order = config.get("generation_order")
    if isinstance(order, list) and order:
        return [str(item) for item in order]
    return list(DEFAULT_NUMBERING_CONFIG["generation_order"])


def section_docx_path(output_dir: Path, section_key: str) -> Path:
    filename = SECTION_DOCX_FILENAMES.get(section_key)
    if not filename:
        raise ValueError(f"unknown section key: {section_key}")
    return Path(output_dir) / filename


def numbering_config_path() -> Path:
    return Path(__file__).resolve().parent / "config" / "reference_layout.json"


def load_numbering_config() -> Dict[str, Any]:
    path = numbering_config_path()
    if not path.exists():
        return dict(DEFAULT_NUMBERING_CONFIG)
    payload = json.loads(path.read_text(encoding="utf-8"))
    numbering = payload.get("numbering") if isinstance(payload, dict) else None
    if not isinstance(numbering, dict):
        return dict(DEFAULT_NUMBERING_CONFIG)
    merged = dict(DEFAULT_NUMBERING_CONFIG)
    merged.update(numbering)
    return merged


def numbering_state_path(output_dir: Optional[Path] = None) -> Path:
    base = Path(output_dir or os.getenv("EIA_OUTPUT_DIR", "output"))
    return base / "debug_tables" / STATE_FILENAME


def reset_numbering(output_dir: Optional[Path] = None) -> "DocxNumbering":
    path = numbering_state_path(output_dir)
    path.parent.mkdir(parents=True, exist_ok=True)
    numbering = DocxNumbering(path)
    numbering.reset()
    return numbering


def load_numbering(output_dir: Optional[Path] = None) -> "DocxNumbering":
    path = numbering_state_path(output_dir)
    path.parent.mkdir(parents=True, exist_ok=True)
    return DocxNumbering(path)


class DocxNumbering:
    def __init__(self, state_path: Path) -> None:
        self.config = load_numbering_config()
        self.numbered_sections = set(self.config.get("numbered_sections") or [])
        self.state_path = state_path
        self.state = self._load()

    def reset(self) -> None:
        self.state = {
            "chapter_index": 0,
            "global_table_order": 0,
            "current_section": None,
            "numbering_plan": {},
            "sections": {},
            "tables": {},
        }
        self._save()

    def capture_table_labels(self) -> Dict[str, str]:
        tables = self.state.get("tables") or {}
        return {
            str(table_key): str(entry.get("label") or "")
            for table_key, entry in tables.items()
            if entry.get("label")
        }

    def finalize_plan(self, ordered_section_keys: List[str]) -> Dict[str, str]:
        plan = {
            str(section_key): str(index + 1)
            for index, section_key in enumerate(ordered_section_keys)
        }
        self.state["numbering_plan"] = plan
        self.state["chapter_index"] = len(plan)
        self.state["global_table_order"] = 0
        self.state["current_section"] = None
        self.state["sections"] = {}
        self.state["tables"] = {}
        self._save()
        return plan

    @staticmethod
    def remap_table_refs(text: str, label_mapping: Dict[str, str]) -> str:
        updated = str(text or "")
        for old_label, new_label in label_mapping.items():
            if old_label and new_label and old_label != new_label:
                updated = updated.replace(old_label, new_label)
        return updated

    def build_table_label_mapping(self, old_labels: Dict[str, str]) -> Dict[str, str]:
        mapping: Dict[str, str] = {}
        for table_key, old_label in old_labels.items():
            new_label = self.table_label(table_key)
            if old_label and new_label and old_label != new_label:
                mapping[old_label] = new_label
        return mapping

    def remap_text_fields(self, texts: Dict[str, str], label_mapping: Dict[str, str]) -> Dict[str, str]:
        if not label_mapping:
            return dict(texts)
        return {
            key: self.remap_table_refs(value, label_mapping)
            for key, value in texts.items()
        }

    def begin_section(self, section_key: str, title: str) -> str:
        sections = self.state.setdefault("sections", {})
        if section_key in sections:
            self.state["current_section"] = section_key
            self._save()
            return self.section_title(section_key, title)

        numbered = section_key in self.numbered_sections
        plan = self.state.get("numbering_plan") or {}
        chapter_number = None
        if numbered:
            if section_key in plan:
                chapter_number = str(plan[section_key])
            else:
                self.state["chapter_index"] = int(self.state.get("chapter_index") or 0) + 1
                chapter_number = str(self.state["chapter_index"])

        sections[section_key] = {
            "title": str(title).strip(),
            "numbered": numbered,
            "chapter_number": chapter_number,
            "table_index": 0,
            "level2_index": 0,
            "level3_index": 0,
        }
        self.state["current_section"] = section_key
        self._save()
        return self.section_title(section_key, title)

    def reset_section_heading_counters(self, section_key: str) -> None:
        section = self.state.get("sections", {}).get(section_key)
        if not section:
            return
        section["level2_index"] = 0
        section["level3_index"] = 0
        self._save()

    def next_chapter_title(self, title: str, section_key: Optional[str] = None) -> str:
        if not section_key:
            raise ValueError("next_chapter_title requires section_key; use begin_section() instead")
        return self.begin_section(section_key, title)

    def section_title(self, section_key: str, title: Optional[str] = None) -> str:
        entry = self.state.get("sections", {}).get(section_key) or {}
        base = str(title or entry.get("title") or "").strip()
        chapter_number = entry.get("chapter_number")
        if chapter_number:
            return f"{chapter_number} {base}"
        return base

    def next_level2_heading(self, section_key: str, title: str) -> str:
        section = self.state.setdefault("sections", {}).setdefault(
            section_key,
            {
                "title": section_key,
                "numbered": False,
                "chapter_number": None,
                "table_index": 0,
                "level2_index": 0,
                "level3_index": 0,
            },
        )
        section["level2_index"] = int(section.get("level2_index") or 0) + 1
        section["level3_index"] = 0
        base = re.sub(r"^\d+(?:\.\d+)*\s*", "", str(title or "").strip())
        self._save()
        chapter_number = section.get("chapter_number")
        if chapter_number:
            return f"{chapter_number}.{section['level2_index']} {base}"
        return base

    def next_level3_heading(self, section_key: str, title: str) -> str:
        section = self.state.setdefault("sections", {}).setdefault(
            section_key,
            {
                "title": section_key,
                "numbered": False,
                "chapter_number": None,
                "table_index": 0,
                "level2_index": 0,
                "level3_index": 0,
            },
        )
        section["level3_index"] = int(section.get("level3_index") or 0) + 1
        base = re.sub(r"^（\d+）", "", str(title or "").strip())
        self._save()
        return f"（{section['level3_index']}）{base}"

    def register_table(self, table_key: str, caption_suffix: str, section_key: Optional[str] = None) -> str:
        tables = self.state.setdefault("tables", {})
        if table_key in tables:
            return self.format_table_caption(tables[table_key]["label"], caption_suffix)

        active_section = section_key or self.state.get("current_section")
        if not active_section:
            raise ValueError(f"register_table({table_key}) requires an active section")

        section = self.state.setdefault("sections", {}).setdefault(
            active_section,
            {
                "title": active_section,
                "numbered": False,
                "chapter_number": None,
                "table_index": 0,
                "level2_index": 0,
                "level3_index": 0,
            },
        )
        section["table_index"] = int(section.get("table_index") or 0) + 1
        local_index = section["table_index"]
        chapter_number = section.get("chapter_number")
        self.state["global_table_order"] = int(self.state.get("global_table_order") or 0) + 1
        if chapter_number:
            label = f"表{chapter_number}.{local_index}"
        else:
            label = f"表{self.state['global_table_order']}"
        tables[table_key] = {
            "label": label,
            "suffix": str(caption_suffix or "").strip(),
            "section_key": active_section,
            "order": self.state["global_table_order"],
        }
        self._save()
        return self.format_table_caption(label, caption_suffix)

    def table_label(self, table_key: str) -> str:
        entry = self.state.get("tables", {}).get(table_key) or {}
        return str(entry.get("label") or "表?")

    def table_caption(self, table_key: str) -> str:
        entry = self.state.get("tables", {}).get(table_key) or {}
        if not entry:
            return f"表?  {table_key}"
        return self.format_table_caption(entry["label"], entry.get("suffix") or "")

    def all_table_labels(self) -> List[str]:
        tables = self.state.get("tables") or {}
        return [
            str(entry.get("label") or "")
            for _, entry in sorted(tables.items(), key=lambda item: item[1].get("order", 0))
            if entry.get("label")
        ]

    def replace_legacy_table_refs(self, text: str, mapping: Dict[str, str]) -> str:
        updated = str(text or "")
        for old_label, table_key in mapping.items():
            new_label = self.table_label(table_key)
            updated = updated.replace(old_label, new_label)
        return updated

    @staticmethod
    def format_table_caption(label: str, caption_suffix: str) -> str:
        suffix = str(caption_suffix or "").strip()
        return f"{label}  {suffix}" if suffix else label

    def _load(self) -> Dict[str, Any]:
        if not self.state_path.exists():
            return self._empty_state()
        payload = json.loads(self.state_path.read_text(encoding="utf-8"))
        if not isinstance(payload, dict):
            return self._empty_state()
        if "sections" not in payload:
            return self._migrate_legacy_state(payload)
        payload.setdefault("sections", {})
        payload.setdefault("tables", {})
        payload.setdefault("global_table_order", 0)
        payload.setdefault("chapter_index", payload.pop("level2_index", payload.get("chapter_index", 0)))
        payload.setdefault("numbering_plan", {})
        for section in payload.get("sections", {}).values():
            section.setdefault("level2_index", 0)
            section.setdefault("level3_index", 0)
        return payload

    def _empty_state(self) -> Dict[str, Any]:
        return {
            "chapter_index": 0,
            "global_table_order": 0,
            "current_section": None,
            "numbering_plan": {},
            "sections": {},
            "tables": {},
        }

    def _migrate_legacy_state(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        state = self._empty_state()
        legacy_chapters = payload.get("chapters") or {}
        for section_key, chapter_number in legacy_chapters.items():
            state["sections"][section_key] = {
                "title": section_key,
                "numbered": True,
                "chapter_number": chapter_number,
                "table_index": 0,
                "level2_index": 0,
                "level3_index": 0,
            }
        state["chapter_index"] = int(payload.get("chapter_index") or payload.get("level2_index") or 0)
        state["tables"] = payload.get("tables") or {}
        state["global_table_order"] = int(payload.get("table_index") or payload.get("global_table_order") or 0)
        return state

    def _save(self) -> None:
        self.state_path.parent.mkdir(parents=True, exist_ok=True)
        self.state_path.write_text(json.dumps(self.state, ensure_ascii=False, indent=2), encoding="utf-8")


def replace_table_label(text: str, old_label: str, new_label: str) -> str:
    return str(text or "").replace(old_label, new_label)
