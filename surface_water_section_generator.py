import json
import os
import subprocess
import sys
from collections import OrderedDict, defaultdict
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


OUTPUT_DIR = Path(os.getenv("EIA_OUTPUT_DIR", "output"))
DEBUG_DIR = OUTPUT_DIR / "debug_tables"
SECTION_PATH = OUTPUT_DIR / "surface_water_section.docx"

FACTOR_ALIASES = {
    "pH": "pH值",
    "PH": "pH值",
    "pH值": "pH值",
    "DO": "溶解氧",
    "溶解氧": "溶解氧",
    "TP": "总磷",
    "总磷": "总磷",
    "SS": "悬浮物",
    "悬浮物": "悬浮物",
    "CODMn": "高锰酸盐指数",
    "高锰酸盐指数": "高锰酸盐指数",
    "NH3-N": "氨氮",
    "氨氮": "氨氮",
}

NOT_APPLICABLE_FACTORS = {"水温", "悬浮物"}


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    DEBUG_DIR.mkdir(parents=True, exist_ok=True)
    ensure_inputs()

    monitoring_records = read_json(OUTPUT_DIR / "monitoring_records.json")
    compliance_results = read_json(OUTPUT_DIR / "compliance_results.json")
    standard_config = read_json(OUTPUT_DIR / "standard_config.json")

    surface_records = [
        record for record in monitoring_records
        if record.get("monitor_type") == "surface_water"
    ]
    surface_results = [
        result for result in compliance_results
        if result.get("monitor_type") == "surface_water"
    ]

    factor_list = detect_factors(surface_records)
    detected_payload = {
        "factor_list": factor_list,
        "source": "output/monitoring_records.json",
        "normalization": FACTOR_ALIASES,
    }
    write_json(DEBUG_DIR / "detected_factors.json", detected_payload)

    table1 = build_monitor_points_table(surface_records, standard_config, factor_list)
    table2 = build_monitor_results_table(surface_records, standard_config, factor_list)
    evaluated_factors = detect_evaluated_factors(surface_results)
    table3 = build_compliance_table(surface_results, standard_config, evaluated_factors)

    write_json(DEBUG_DIR / "surface_water_monitor_points_table.json", table1)
    write_json(DEBUG_DIR / "surface_water_monitor_results_table.json", table2)
    write_json(DEBUG_DIR / "surface_water_compliance_table.json", table3)

    conclusion = build_conclusion(surface_results)
    doc = build_docx(table1, table2, table3, factor_list, evaluated_factors, conclusion)
    doc.save(SECTION_PATH)
    print(f"generated: {SECTION_PATH}")
    print(f"detected_factors: {len(factor_list)}")
    print(f"table1 rows: {len(table1['rows'])}")
    print(f"table2 rows: {len(table2['rows'])}")
    print(f"table3 rows: {len(table3['rows'])}")


def ensure_inputs() -> None:
    required = [
        OUTPUT_DIR / "monitoring_records.json",
        OUTPUT_DIR / "standard_config.json",
    ]
    compliance_path = OUTPUT_DIR / "compliance_results.json"
    if all(path.exists() for path in required) and compliance_path.exists():
        return

    engine = Path("compliance_engine.py")
    if engine.exists():
        subprocess.run([sys.executable, str(engine)], check=True)
    else:
        subprocess.run([sys.executable, "surface_water_pipeline.py"], check=True)


def read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path: Path, payload: Any) -> None:
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def detect_factors(records: List[Dict[str, Any]]) -> List[str]:
    factors: List[str] = []
    seen: set[str] = set()
    for record in records:
        factor = normalize_factor(record.get("factor"))
        if not factor or factor in seen:
            continue
        seen.add(factor)
        factors.append(factor)
    return factors


def detect_evaluated_factors(results: List[Dict[str, Any]]) -> List[str]:
    factors: List[str] = []
    seen: set[str] = set()
    for result in results:
        factor = normalize_factor(result.get("factor"))
        method = result.get("method")
        not_applicable = bool(result.get("not_applicable")) or method == "not_applicable"
        if not factor or factor in seen or not_applicable or factor in NOT_APPLICABLE_FACTORS:
            continue
        seen.add(factor)
        factors.append(factor)
    return factors


def normalize_factor(value: Any) -> str:
    text = str(value or "").strip()
    return FACTOR_ALIASES.get(text, text)


def points_config(standard_config: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    return standard_config.get("points") or {}


def build_monitor_points_table(
    records: List[Dict[str, Any]],
    standard_config: Dict[str, Any],
    factor_list: List[str],
) -> Dict[str, Any]:
    configs = points_config(standard_config)
    rows: List[Dict[str, Any]] = []
    for index, point_code in enumerate(sorted_point_codes(records, configs), start=1):
        config = configs.get(point_code, {})
        first_record = first_record_for_point(records, point_code)
        point_text = (first_record or {}).get("point") or ""
        parsed = parse_point_text(point_text)
        evidence = config.get("evidence") or {}
        dates = unique_ordered(
            record.get("sample_date")
            for record in records
            if record.get("point_code") == point_code
        )
        rows.append(
            {
                "序号": point_code,
                "河流名称": config.get("river_name") or parsed.get("river_name") or "-",
                "中心桩号": evidence.get("中心桩号") or parsed.get("center_station") or "-",
                "取样断面": evidence.get("取样断面") or parsed.get("sampling_section") or "-",
                "取样频次": evidence.get("取样频次") or infer_frequency(dates),
                "监测因子": "、".join(factor_list),
                "point_code": point_code,
                "_point_order": index,
            }
        )
    return {
        "title": "表4.2-4 水质监测断面布置",
        "headers": ["序号", "河流名称", "中心桩号", "取样断面", "取样频次", "监测因子"],
        "rows": rows,
    }


def build_monitor_results_table(
    records: List[Dict[str, Any]],
    standard_config: Dict[str, Any],
    factor_list: List[str],
) -> Dict[str, Any]:
    configs = points_config(standard_config)
    point_orders = {code: index for index, code in enumerate(sorted_point_codes(records, configs), start=1)}
    grouped: Dict[Tuple[str, str], Dict[str, Any]] = {}
    for record in records:
        point_code = record.get("point_code")
        sample_date = record.get("sample_date")
        if not point_code or not sample_date:
            continue
        key = (point_code, sample_date)
        if key not in grouped:
            config = configs.get(point_code, {})
            river_name = config.get("river_name") or parse_point_text(record.get("point") or "").get("river_name") or "-"
            grouped[key] = {
                "序号": point_orders.get(point_code, len(point_orders) + 1),
                "河流": f"{river_name}（{point_code}）",
                "监测时间": sample_date,
                "point_code": point_code,
            }
            for factor in factor_list:
                grouped[key][factor] = "-"
        factor = normalize_factor(record.get("factor"))
        if factor in factor_list:
            grouped[key][factor] = record.get("value") or "-"
    headers = ["序号", "河流", "监测时间", *factor_list]
    return {
        "title": "表4.2-5 现状监测结果表",
        "headers": headers,
        "rows": [grouped[key] for key in sorted(grouped, key=point_date_sort_key)],
    }


def build_compliance_table(
    results: List[Dict[str, Any]],
    standard_config: Dict[str, Any],
    evaluated_factors: List[str],
) -> Dict[str, Any]:
    configs = points_config(standard_config)
    grouped: Dict[Tuple[str, str], Dict[str, Any]] = {}
    review_cells = defaultdict(set)
    for result in results:
        point_code = result.get("point_code")
        sample_date = result.get("sample_date")
        if not point_code or not sample_date:
            continue
        key = (point_code, sample_date)
        if key not in grouped:
            config = configs.get(point_code, {})
            river_name = result.get("river_name") or config.get("river_name") or "-"
            grouped[key] = {
                "编号": f"{river_name}（{point_code}）",
                "监测时间": sample_date,
                "point_code": point_code,
            }
            for factor in evaluated_factors:
                grouped[key][factor] = "-"
        factor = normalize_factor(result.get("factor"))
        if factor not in evaluated_factors:
            continue
        method = result.get("method")
        not_applicable = bool(result.get("not_applicable")) or method == "not_applicable"
        if not_applicable:
            grouped[key][factor] = "-"
        elif result.get("needs_review"):
            grouped[key][factor] = "需复核"
            review_cells[key].add(factor)
        else:
            index = result.get("standard_index")
            grouped[key][factor] = format_index(index)
    headers = ["编号", "监测时间", *evaluated_factors]
    return {
        "title": "表4.2-6 地表水环境现状评价结果",
        "headers": headers,
        "rows": [grouped[key] for key in sorted(grouped, key=point_date_sort_key)],
        "evaluated_factors": evaluated_factors,
        "review_cells": [
            {"point_code": key[0], "sample_date": key[1], "factor": factor}
            for key, factors in review_cells.items()
            for factor in sorted(factors)
        ],
    }


def build_conclusion(results: List[Dict[str, Any]]) -> str:
    exceed_items = [
        result for result in results
        if result.get("is_compliant") is False
    ]
    if not exceed_items:
        rivers = unique_ordered(result.get("river_name") for result in results if result.get("river_name"))
        factors = unique_ordered(
            normalize_factor(result.get("factor"))
            for result in results
            if result.get("is_compliant") is True
        )
        river_text = "、".join(rivers) if rivers else "各监测断面"
        factor_text = "、".join(factors) if factors else "各评价因子"
        return (
            f"根据监测及评价结果，{river_text}的{factor_text}等评价因子"
            "均满足《地表水环境质量标准》（GB3838-2002）相应水质类别标准要求。"
        )

    details = []
    for item in exceed_items:
        details.append(
            f"{item.get('river_name') or '-'}（{item.get('point_code') or '-'}）"
            f"{item.get('sample_date') or '-'} {item.get('factor') or '-'}"
        )
    return "根据现状监测及评价结果，存在超标情况，超标断面、因子及监测日期为：" + "；".join(details) + "。"


def build_docx(
    table1: Dict[str, Any],
    table2: Dict[str, Any],
    table3: Dict[str, Any],
    factor_list: List[str],
    evaluated_factors: List[str],
    conclusion: str,
) -> Document:
    doc = Document()
    setup_document(doc)

    doc.add_heading("1.1.1 地表水环境现状调查与评价", level=1)
    doc.add_heading("1.1.1.1 水环境质量现状调查", level=2)

    add_numbered_title(doc, "（1）现状监测点布置")
    doc.add_paragraph(
        f"根据项目所在区域的水文特征、河流水体规模，共计在评价范围设置了"
        f"{len(table1['rows'])}个监测断面进行水质监测。监测断面概况详见表4.2-4。"
    )
    add_caption(doc, table1["title"])
    add_table(doc, table1["headers"], table1["rows"])

    add_numbered_title(doc, "（2）监测时间、频率和方法")
    doc.add_paragraph(build_monitoring_description(table1, factor_list))

    add_numbered_title(doc, "（3）现状监测结果")
    doc.add_paragraph("本项目地表水监测结果详见表4.2-5。")
    add_caption(doc, table2["title"])
    add_table(doc, table2["headers"], table2["rows"])

    add_numbered_title(doc, "（4）现状评价结果")
    add_subtitle(doc, "①评价方法")
    add_evaluation_method_text(doc)
    add_subtitle(doc, "②评价结果")
    doc.add_paragraph("地表水监测点位环境现状评价结果见表4.2-6。")
    add_caption(doc, table3["title"])
    add_table(doc, table3["headers"], table3["rows"])

    doc.add_paragraph(conclusion)
    return doc


def setup_document(doc: Document) -> None:
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.8)
    section.left_margin = Cm(1.6)
    section.right_margin = Cm(1.6)

    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "宋体"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = RGBColor(0, 0, 0)

    for style_name, size in [("Heading 1", 14), ("Heading 2", 12), ("Heading 3", 11)]:
        style = styles[style_name]
        style.font.name = "黑体"
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "黑体")
        style.font.size = Pt(size)
        style.font.color.rgb = RGBColor(0, 0, 0)


def add_numbered_title(doc: Document, text: str) -> None:
    paragraph = doc.add_paragraph()
    run = paragraph.add_run(text)
    run.bold = True


def add_subtitle(doc: Document, text: str) -> None:
    paragraph = doc.add_paragraph()
    run = paragraph.add_run(text)
    run.bold = True


def add_caption(doc: Document, text: str) -> None:
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run(text)
    run.bold = True


def add_table(doc: Document, headers: List[str], rows: List[Dict[str, Any]]) -> None:
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.autofit = True
    header_cells = table.rows[0].cells
    for index, header in enumerate(headers):
        header_cells[index].text = str(header)
        shade_cell(header_cells[index], "EDEDED")
        set_cell_text_style(header_cells[index], bold=True)

    for row in rows:
        cells = table.add_row().cells
        for index, header in enumerate(headers):
            cells[index].text = str(row.get(header, "-"))
            set_cell_text_style(cells[index], bold=False)

    for row in table.rows:
        for cell in row.cells:
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_margin(cell)
    doc.add_paragraph()


def add_evaluation_method_text(doc: Document) -> None:
    doc.add_paragraph("现状监测结果按水质指数法进行单项水质参数评价，计算公式如下：")
    add_formula(doc, "Sᵢ,ⱼ = Cᵢ,ⱼ / Cₛᵢ")
    doc.add_paragraph("式中：Sᵢ,ⱼ——评价因子i的水质指数，大于1表明该水质因子超标；")
    doc.add_paragraph("Cᵢ,ⱼ——评价因子i在j点的实测统计代表值，mg/L；")
    doc.add_paragraph("Cₛᵢ——评价因子i的水质评价标准限值，mg/L。")
    doc.add_paragraph("其中，pH的标准指数为：")
    add_formula(doc, "SₚH,ⱼ = (7.0 - pHⱼ) / (7.0 - pHsd)        （pHⱼ≤7.0）")
    add_formula(doc, "SₚH,ⱼ = (pHⱼ - 7.0) / (pHsu - 7.0)        （pHⱼ＞7.0）")
    doc.add_paragraph("DO的标准指数为：")
    add_formula(doc, "DOsat = 468 / (31.6 + T)")
    add_formula(doc, "SᴅO,ⱼ = |DOf - DOⱼ| / (DOf - DOs)        （DOⱼ≥DOs）")
    add_formula(doc, "SᴅO,ⱼ = 10 - 9DOⱼ / DOs        （DOⱼ＜DOs）")
    doc.add_paragraph("式中：SₚH,ⱼ——pH值的指数，大于1表明该水质因子超标；")
    doc.add_paragraph("pHⱼ——pH值实测统计代表值；")
    doc.add_paragraph("pHsd——评价标准中pH值的下限值；")
    doc.add_paragraph("pHsu——评价标准中pH值的上限值；")
    doc.add_paragraph("SᴅO,ⱼ——溶解氧的标准指数，大于1表明该水质因子超标；")
    doc.add_paragraph("DOⱼ——溶解氧在j点的实测统计代表值，mg/L；")
    doc.add_paragraph("DOs——溶解氧的水质评价标准限值，mg/L；")
    doc.add_paragraph("DOf——饱和溶解氧浓度，mg/L；")
    doc.add_paragraph("T——水温，℃。")
    doc.add_paragraph("水温、悬浮物等无相应限值的指标仅列出现状监测值，不参与标准指数评价。")


def add_formula(doc: Document, text: str) -> None:
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run(text)
    run.font.name = "Times New Roman"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    run.font.size = Pt(10.5)
    run.font.color.rgb = RGBColor(0, 0, 0)


def set_cell_text_style(cell: Any, bold: bool = False) -> None:
    for paragraph in cell.paragraphs:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in paragraph.runs:
            run.font.name = "宋体"
            run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
            run.font.size = Pt(8)
            run.font.color.rgb = RGBColor(0, 0, 0)
            run.bold = bold


def shade_cell(cell: Any, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), fill)
    tc_pr.append(shading)


def set_cell_margin(cell: Any, margin: int = 80) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for side in ("top", "left", "bottom", "right"):
        node = tc_mar.find(qn(f"w:{side}"))
        if node is None:
            node = OxmlElement(f"w:{side}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(margin))
        node.set(qn("w:type"), "dxa")


def build_monitoring_description(table1: Dict[str, Any], factor_list: List[str]) -> str:
    frequencies = unique_ordered(row.get("取样频次") for row in table1["rows"] if row.get("取样频次"))
    frequency_text = "；".join(frequencies) if frequencies else "-"
    factors = "、".join(factor_list)
    return (
        f"本次地表水现状监测断面按监测方案布设，取样频次为{frequency_text}。"
        f"监测因子包括{factors}。各监测因子检测方法按监测报告及相关水质检测标准执行。"
    )


def sorted_point_codes(
    records: List[Dict[str, Any]],
    configs: Dict[str, Dict[str, Any]],
) -> List[str]:
    codes = unique_ordered([*configs.keys(), *(record.get("point_code") for record in records)])
    return sorted(codes, key=point_sort_key)


def point_sort_key(code: str) -> Tuple[int, str]:
    digits = "".join(ch for ch in str(code) if ch.isdigit())
    return (int(digits) if digits else 9999, str(code))


def point_date_sort_key(key: Tuple[str, str]) -> Tuple[int, str, Tuple[int, ...], str]:
    point_code, sample_date = key
    return (*point_sort_key(point_code), date_sort_key(sample_date), str(sample_date))


def date_sort_key(value: Any) -> Tuple[int, ...]:
    numbers = [int(item) for item in re_find_numbers(str(value or ""))]
    return tuple(numbers) if numbers else (9999,)


def re_find_numbers(text: str) -> List[str]:
    import re

    return re.findall(r"\d+", text)


def first_record_for_point(records: List[Dict[str, Any]], point_code: str) -> Optional[Dict[str, Any]]:
    return next((record for record in records if record.get("point_code") == point_code), None)


def parse_point_text(text: str) -> Dict[str, Optional[str]]:
    parts = str(text or "").split()
    result: Dict[str, Optional[str]] = {
        "center_station": None,
        "river_name": None,
        "sampling_section": None,
    }
    if len(parts) >= 2:
        result["center_station"] = parts[1]
    if len(parts) >= 3:
        result["river_name"] = parts[2]
    if len(parts) >= 4:
        section_parts = []
        for part in parts[3:]:
            if part.startswith("点位编号"):
                break
            section_parts.append(part)
        result["sampling_section"] = "".join(section_parts) if section_parts else None
    return result


def infer_frequency(dates: List[str]) -> str:
    if len(dates) == 3:
        return "连续取样三天，每天一次"
    if len(dates) == 1:
        return "取样一天，一天一次"
    if dates:
        return f"共取样{len(dates)}天"
    return "-"


def unique_ordered(items: Iterable[Any]) -> List[str]:
    result: List[str] = []
    seen: set[str] = set()
    for item in items:
        text = str(item or "").strip()
        if not text or text in seen:
            continue
        seen.add(text)
        result.append(text)
    return result


def format_index(value: Any) -> str:
    if value is None:
        return "-"
    try:
        return f"{float(value):.2f}"
    except (TypeError, ValueError):
        return "需复核"


if __name__ == "__main__":
    main()
