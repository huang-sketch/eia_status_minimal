"""Finalize section numbering and rebuild chapter docx files before merge."""

from pathlib import Path
from typing import Dict, List

from docx_numbering import (
    generation_order,
    load_numbering,
    section_docx_path,
)
import noise_section_generator as noise_generator
import project_area_overview_generator as overview_generator
import surface_water_section_generator as surface_generator

NOISE_TABLE_KEYS = {
    noise_generator.NOISE_TABLE_MONITOR_POINTS,
    noise_generator.NOISE_TABLE_SENSITIVE,
    noise_generator.NOISE_TABLE_INDOOR,
    noise_generator.NOISE_TABLE_ATTENUATION,
}
SURFACE_TABLE_KEYS = {
    surface_generator.SURFACE_TABLE_MONITOR_POINTS,
    surface_generator.SURFACE_TABLE_MONITOR_RESULTS,
    surface_generator.SURFACE_TABLE_COMPLIANCE,
}


def list_present_sections(output_dir: Path) -> List[str]:
    output_dir = Path(output_dir)
    return [
        section_key
        for section_key in generation_order()
        if section_docx_path(output_dir, section_key).exists()
    ]


def finalize_and_rebuild_section_docx(output_dir: Path) -> Dict[str, str]:
    output_dir = Path(output_dir)
    present_sections = list_present_sections(output_dir)
    if not present_sections:
        return {}

    numbering = load_numbering(output_dir)
    old_table_labels = numbering.capture_table_labels()
    plan = numbering.finalize_plan(present_sections)

    for section_key in present_sections:
        label_mapping = _label_mapping_for_section(section_key, old_table_labels, numbering)
        _rebuild_section_docx(section_key, output_dir, numbering, label_mapping)

    return plan


def _label_mapping_for_section(
    section_key: str,
    old_table_labels: Dict[str, str],
    numbering,
) -> Dict[str, str]:
    if section_key == "noise":
        section_old = {
            key: label
            for key, label in old_table_labels.items()
            if key in NOISE_TABLE_KEYS
        }
    elif section_key == "surface_water":
        section_old = {
            key: label
            for key, label in old_table_labels.items()
            if key in SURFACE_TABLE_KEYS
        }
    else:
        return {}

    if not section_old:
        return {}

    if section_key == "noise":
        debug_dir = numbering.state_path.parent
        table1 = noise_generator.read_json(debug_dir / "noise_monitor_points_table.json")
        table2 = noise_generator.read_json(debug_dir / "noise_sensitive_points_result_table.json")
        indoor_path = debug_dir / "noise_indoor_result_table.json"
        table_indoor = (
            noise_generator.read_json(indoor_path)
            if indoor_path.exists()
            else noise_generator.empty_indoor_result_table()
        )
        table3 = noise_generator.read_json(debug_dir / "traffic_noise_attenuation_table.json")
        numbering.begin_section("noise", "声环境现状调查与评价")
        noise_generator.register_noise_tables(numbering, table1, table2, table3, table_indoor)
    elif section_key == "surface_water":
        debug_dir = numbering.state_path.parent
        table1 = surface_generator.read_json(debug_dir / "surface_water_monitor_points_table.json")
        table2 = surface_generator.read_json(debug_dir / "surface_water_monitor_results_table.json")
        table3 = surface_generator.read_json(debug_dir / "surface_water_compliance_table.json")
        numbering.begin_section("surface_water", "地表水环境现状调查与评价")
        surface_generator.register_surface_tables(numbering, table1, table2, table3)

    return numbering.build_table_label_mapping(section_old)


def _rebuild_section_docx(
    section_key: str,
    output_dir: Path,
    numbering,
    label_mapping: Dict[str, str],
) -> Path:
    if section_key == "overview":
        return overview_generator.rebuild_docx_from_output(output_dir, numbering)
    if section_key == "noise":
        return noise_generator.rebuild_docx_from_output(output_dir, numbering, label_mapping)
    if section_key == "surface_water":
        return surface_generator.rebuild_docx_from_output(output_dir, numbering, label_mapping)
    raise ValueError(f"unknown section key: {section_key}")
