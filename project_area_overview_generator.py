import json
import os
import re
from datetime import datetime
from pathlib import Path
from statistics import mean
from typing import Any, Dict, Iterable, List, Optional

from dotenv import load_dotenv
from docx import Document

from docx_layout import add_body_paragraph, add_chapter_title, add_section_heading, create_section_document, finalize_section_document, setup_document
from docx_numbering import DocxNumbering, load_numbering
from llm_client import LlmProfile, chat_completion_json_object


load_dotenv()

INPUT_DIR = Path(os.getenv("EIA_INPUT_DIR", "input"))
OUTPUT_DIR = Path(os.getenv("EIA_OUTPUT_DIR", "output"))
DEBUG_DIR = OUTPUT_DIR / "debug_tables"
OUTPUT_DOCX = OUTPUT_DIR / "project_area_overview.docx"
REFERENCE_DOCX = Path(
    os.getenv(
        "EIA_PROJECT_OVERVIEW_REFERENCE_DOCX",
        r"D:\华设\智能体\环评报告\参考报告\项目区域环境概况.docx",
    )
)

SECTION_TITLES = ["地理位置", "地形、地貌", "气候", "水文", "地质、地震"]
STATUS_FILE = DEBUG_DIR / "project_area_overview_status.json"
LLM_INPUT_FILE = DEBUG_DIR / "project_area_overview_llm_input.json"
LLM_OUTPUT_FILE = DEBUG_DIR / "project_area_overview_llm_output.json"
TEXT_FILE = DEBUG_DIR / "project_area_overview_text.json"
SOURCES_FILE = DEBUG_DIR / "project_area_overview_sources.json"


class MissingLlmApiKeyError(RuntimeError):
    pass


class LlmCallError(RuntimeError):
    pass


class LlmValidationError(RuntimeError):
    def __init__(self, errors: List[str], warnings: Optional[List[str]], debug_output: Dict[str, Any]) -> None:
        super().__init__("LLM 区域概况输出未通过校验：" + "；".join(errors))
        self.errors = errors
        self.warnings = warnings or []
        self.debug_output = debug_output


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    DEBUG_DIR.mkdir(parents=True, exist_ok=True)
    remove_stale_docx()

    meta = read_project_meta()
    report_name = str(meta.get("report_name") or "").strip()
    admin_division = str(meta.get("admin_division") or "").strip()
    length_profile = read_reference_length_profile()
    llm_input = {
        "report_name": report_name,
        "admin_division": admin_division,
        "section_titles": SECTION_TITLES,
        "reference_length_profile": length_profile,
        "reference_docx": str(REFERENCE_DOCX) if REFERENCE_DOCX.exists() else None,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "generation_mode": "llm_direct_baike_style",
    }
    write_json(LLM_INPUT_FILE, llm_input)
    write_json(
        SOURCES_FILE,
        {
            "generation_mode": "llm_direct_baike_style",
            "note": "区域概况由当前配置的大模型直接生成；篇幅参考本地参考报告，表达采用百科式公开资料风格，不自动大段复制网页原文。",
            "admin_division": admin_division,
            "report_name": report_name,
            "reference_docx": str(REFERENCE_DOCX) if REFERENCE_DOCX.exists() else None,
            "sources": [],
        },
    )

    try:
        if not admin_division:
            raise RuntimeError("project_meta.json 中缺少 admin_division，跳过项目区域环境概况。")
        if not os.getenv("EIA_LLM_API_KEY"):
            raise MissingLlmApiKeyError("EIA_LLM_API_KEY 未配置，跳过项目区域环境概况。")

        generation = generate_sections_with_retry(llm_input)
        write_json(LLM_OUTPUT_FILE, generation["debug_output"])
        sections = generation["sections"]

        payload = {
            "admin_division": admin_division,
            "report_name": report_name,
            "generation_mode": "llm_direct_baike_style",
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "sections": [
                {"title": title, "paragraphs": sections[title]}
                for title in SECTION_TITLES
            ],
        }
        write_json(TEXT_FILE, {
            "admin_division": admin_division,
            "report_name": report_name,
            "sections": sections,
        })
        doc = build_docx(payload)
        finalize_section_document(doc)
        doc.save(OUTPUT_DOCX)
        write_status(generation["status"], None, admin_division, report_name, length_profile, generation.get("warnings", []))
        print(f"generated: {OUTPUT_DOCX}")
    except MissingLlmApiKeyError as exc:
        remove_stale_docx()
        write_json(LLM_OUTPUT_FILE, {"error": str(exc), "attempts": []})
        write_json(TEXT_FILE, {
            "admin_division": admin_division,
            "report_name": report_name,
            "sections": {},
        })
        write_status("skipped_no_api_key", str(exc), admin_division, report_name, length_profile)
        print(f"project area overview skipped: {exc}")
    except LlmCallError as exc:
        remove_stale_docx()
        write_json(LLM_OUTPUT_FILE, {"error": str(exc), "attempts": []})
        write_json(TEXT_FILE, {
            "admin_division": admin_division,
            "report_name": report_name,
            "sections": {},
        })
        write_status("llm_call_failed", str(exc), admin_division, report_name, length_profile)
        print(f"project area overview skipped: {exc}")
    except LlmValidationError as exc:
        remove_stale_docx()
        write_json(LLM_OUTPUT_FILE, exc.debug_output)
        write_json(TEXT_FILE, {
            "admin_division": admin_division,
            "report_name": report_name,
            "sections": {},
        })
        write_status("llm_validation_failed", str(exc), admin_division, report_name, length_profile, exc.warnings)
        print(f"project area overview skipped: {exc}")
    except Exception as exc:
        remove_stale_docx()
        write_json(LLM_OUTPUT_FILE, {"error": str(exc)})
        write_json(TEXT_FILE, {
            "admin_division": admin_division,
            "report_name": report_name,
            "sections": {},
        })
        write_status("skipped", str(exc), admin_division, report_name, length_profile)
        print(f"project area overview skipped: {exc}")


def read_project_meta() -> Dict[str, Any]:
    path = INPUT_DIR / "project_meta.json"
    if not path.exists():
        return {}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}
    return payload if isinstance(payload, dict) else {}


def read_reference_length_profile() -> Dict[str, Any]:
    default_profile = {
        "paragraphs_per_section": "2-4",
        "chars_per_paragraph": "160-300",
        "target_chars_per_section": "400-900",
        "source": "default",
    }
    if not REFERENCE_DOCX.exists():
        return default_profile
    try:
        doc = Document(str(REFERENCE_DOCX))
    except Exception as exc:
        return {**default_profile, "source": "default", "error": str(exc)}

    section_lengths: Dict[str, List[int]] = {title: [] for title in SECTION_TITLES}
    current_title = ""
    for paragraph in doc.paragraphs:
        text = clean_paragraph(paragraph.text)
        if not text:
            continue
        matched_title = match_section_title(text)
        if matched_title:
            current_title = matched_title
            continue
        if current_title in section_lengths and len(text) >= 30:
            section_lengths[current_title].append(len(text))

    populated = {title: lengths for title, lengths in section_lengths.items() if lengths}
    all_lengths = [length for lengths in populated.values() for length in lengths]
    if not all_lengths:
        return {**default_profile, "source": str(REFERENCE_DOCX), "warning": "no_matching_paragraphs"}

    average = int(mean(all_lengths))
    low = max(120, int(average * 0.75))
    high = min(420, max(low + 40, int(average * 1.35)))
    per_section = {
        title: {
            "paragraph_count": len(lengths),
            "avg_chars": int(mean(lengths)),
            "total_chars": sum(lengths),
        }
        for title, lengths in populated.items()
    }
    avg_section_total = int(mean([item["total_chars"] for item in per_section.values()]))
    return {
        "paragraphs_per_section": "2-4",
        "chars_per_paragraph": f"{low}-{high}",
        "target_chars_per_section": f"{max(350, int(avg_section_total * 0.8))}-{max(500, int(avg_section_total * 1.2))}",
        "source": str(REFERENCE_DOCX),
        "per_section": per_section,
    }


def match_section_title(text: str) -> str:
    normalized = re.sub(r"^[\d.、\s]+", "", text)
    for title in SECTION_TITLES:
        if normalized == title or normalized.endswith(title):
            return title
    return ""


def write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def write_status(
    status: str,
    error: Optional[str],
    admin_division: str,
    report_name: str,
    length_profile: Dict[str, Any],
    warnings: Optional[List[str]] = None,
) -> None:
    write_json(
        STATUS_FILE,
        {
            "status": status,
            "used_llm": status in {"success", "llm_retry_success"},
            "generation_mode": "llm_direct_baike_style",
            "admin_division": admin_division,
            "report_name": report_name,
            "reference_length_profile": length_profile,
            "warnings": warnings or [],
            "error": error,
            "output_docx": str(OUTPUT_DOCX) if OUTPUT_DOCX.exists() else None,
            "updated_at": datetime.now().isoformat(timespec="seconds"),
        },
    )


def remove_stale_docx() -> None:
    if OUTPUT_DOCX.exists():
        OUTPUT_DOCX.unlink()


def generate_sections_with_retry(llm_input: Dict[str, Any]) -> Dict[str, Any]:
    attempts: List[Dict[str, Any]] = []
    first_prompt = build_prompt(llm_input)
    raw_output = request_overview_json(first_prompt)
    sections = normalize_llm_sections(raw_output)
    errors, warnings = validate_sections(sections)
    attempts.append(
        {
            "attempt": 1,
            "prompt_type": "initial",
            "raw_output": raw_output,
            "normalized_sections": sections,
            "validation_errors": errors,
            "validation_warnings": warnings,
        }
    )
    if not errors:
        return {
            "status": "success",
            "sections": sections,
            "warnings": warnings,
            "debug_output": {"attempts": attempts, "selected_attempt": 1},
        }

    retry_prompt = build_retry_prompt(llm_input, sections, errors, warnings)
    retry_output = request_overview_json(retry_prompt)
    retry_sections = normalize_llm_sections(retry_output)
    retry_errors, retry_warnings = validate_sections(retry_sections)
    attempts.append(
        {
            "attempt": 2,
            "prompt_type": "validation_repair",
            "raw_output": retry_output,
            "normalized_sections": retry_sections,
            "validation_errors": retry_errors,
            "validation_warnings": retry_warnings,
        }
    )
    if not retry_errors:
        return {
            "status": "llm_retry_success",
            "sections": retry_sections,
            "warnings": retry_warnings,
            "debug_output": {"attempts": attempts, "selected_attempt": 2},
        }

    raise LlmValidationError(
        retry_errors,
        retry_warnings,
        {"attempts": attempts, "selected_attempt": None},
    )


def build_prompt(payload: Dict[str, Any]) -> str:
    length_profile = payload.get("reference_length_profile") or {}
    chars_per_paragraph = length_profile.get("chars_per_paragraph") or "160-300"
    target_chars_per_section = length_profile.get("target_chars_per_section") or "400-900"
    return f"""
你是环评报告“项目区域环境概况”章节撰写助手。请根据输入的项目名称和行政区划，直接生成区域概况正文。

输入信息：
- 项目名称：{payload.get("report_name") or "未提供"}
- 行政区划：{payload.get("admin_division") or "未提供"}

篇幅要求：
- 参考报告段落长度画像：{json.dumps(length_profile, ensure_ascii=False)}
- 每个小节尽量写 2 到 4 段。
- 每段尽量控制在 {chars_per_paragraph} 个中文字符。
- 每个小节总长度尽量控制在 {target_chars_per_section} 个中文字符，使整体篇幅接近参考报告。

写作要求：
1. 固定输出 5 个小节：地理位置、地形、地貌、气候、水文、地质、地震。
2. 内容主体采用百度百科、地方概况、公开常识类资料的客观表述风格，句式可以接近百科原文，但不要输出网页菜单、脚注、版权信息，也不要整段连续复制受版权保护网页正文。
3. 如果行政区划到县级市、区、县，优先写该县级行政区；如果只到地级市，则写地级市区域概况。
4. 可以写该行政区常见的区位、地貌、气候、水系、地质地震概况；不确定的精确经纬度、极端气温、年降水量、地震动参数不要写。
5. 不要编造项目路线、具体厂址、具体监测点位、具体断裂带、具体河流或水库等无法从输入确认的信息。
6. 语言风格应接近环评报告正文，客观、平实、信息密度高，不要使用宣传性语气。
7. 只输出合法 JSON，不要输出 Markdown、代码块或解释。

JSON 格式：
{{
  "sections": {{
    "地理位置": ["段落1", "段落2"],
    "地形、地貌": ["段落1", "段落2"],
    "气候": ["段落1", "段落2"],
    "水文": ["段落1", "段落2"],
    "地质、地震": ["段落1", "段落2"]
  }}
}}
""".strip()


def build_retry_prompt(
    llm_input: Dict[str, Any],
    sections: Dict[str, List[str]],
    errors: List[str],
    warnings: List[str],
) -> str:
    return (
        build_prompt(llm_input)
        + "\n\n上一次输出未通过本地校验，请只修正问题并重新输出完整 JSON。"
        + f"\n校验错误：{json.dumps(errors, ensure_ascii=False)}"
        + f"\n校验警告：{json.dumps(warnings, ensure_ascii=False)}"
        + f"\n上一次已解析文本：{json.dumps(sections, ensure_ascii=False)}"
        + "\n修正要求：补足空小节或明显过短小节；删除网页噪声、异常英文长串和乱码；涉及气象或地震参数时只保留概述性表述，不写精确数值。"
    )


def request_overview_json(prompt: str) -> Dict[str, Any]:
    try:
        return chat_completion_json_object(
            [
                {
                    "role": "system",
                    "content": "你是严谨的环评报告写作助手，只输出符合要求的 JSON。",
                },
                {"role": "user", "content": prompt},
            ],
            profile=LlmProfile.project_overview,
            label="project_area_overview",
        )
    except RuntimeError as exc:
        raise LlmCallError(str(exc)) from exc


def normalize_llm_sections(payload: Dict[str, Any]) -> Dict[str, List[str]]:
    raw_sections = payload.get("sections", payload)
    if not isinstance(raw_sections, dict):
        raise ValueError("LLM 输出缺少 sections 对象")
    sections: Dict[str, List[str]] = {}
    for title in SECTION_TITLES:
        value = raw_sections.get(title)
        if isinstance(value, str):
            paragraphs = [value]
        elif isinstance(value, list):
            paragraphs = [str(item).strip() for item in value if str(item).strip()]
        else:
            paragraphs = []
        sections[title] = unique_ordered(clean_paragraph(paragraph) for paragraph in paragraphs)[:4]
    return sections


def validate_sections(sections: Dict[str, List[str]]) -> tuple[List[str], List[str]]:
    errors: List[str] = []
    warnings: List[str] = []
    for title in SECTION_TITLES:
        paragraphs = sections.get(title) or []
        if not paragraphs:
            errors.append(f"{title}为空")
            continue
        if len(paragraphs) < 2:
            warnings.append(f"{title}段落数量偏少")
        total_chars = sum(len(paragraph) for paragraph in paragraphs)
        if total_chars < 120:
            warnings.append(f"{title}篇幅偏短")
        for paragraph in paragraphs:
            if not is_chinese_text(paragraph):
                errors.append(f"{title}含非中文或乱码段落")
            if re.search(r"[A-Za-z]{20,}", paragraph) or english_ratio(paragraph) > 0.1:
                errors.append(f"{title}含异常英文长串")
            if any(marker in paragraph for marker in ("暂无可直接引用", "网站首页", "当前位置", "版权所有", "备案号")):
                errors.append(f"{title}含网页噪声或兜底占位文本")
            if any(marker in paragraph for marker in ("极端最高气温", "极端最低气温", "地震动峰值加速度")):
                warnings.append(f"{title}含精确参数表述，建议改为概述性文字")
    return unique_ordered(errors), unique_ordered(warnings)


def clean_paragraph(value: str) -> str:
    text = re.sub(r"\s+", "", str(value or ""))
    text = re.sub(r"^[：:;；，,。.\-—]+", "", text)
    return text.strip()


def is_chinese_text(text: str) -> bool:
    value = str(text or "")
    chinese_chars = re.findall(r"[\u4e00-\u9fff]", value)
    return len(chinese_chars) >= 20 and len(chinese_chars) / max(len(value), 1) >= 0.75


def english_ratio(text: str) -> float:
    value = str(text or "")
    letters = re.findall(r"[A-Za-z]", value)
    return len(letters) / max(len(value), 1)


def build_docx(payload: Dict[str, Any], numbering: Optional[DocxNumbering] = None) -> Document:
    numbering = numbering or load_numbering(OUTPUT_DIR)
    numbering.begin_section("overview", "项目区域环境概况")
    doc = create_section_document()
    setup_document(doc)
    add_chapter_title(doc, numbering.section_title("overview", "项目区域环境概况"))
    for section in payload.get("sections", []):
        title = section.get("title") or "-"
        add_section_heading(doc, numbering.next_level2_heading("overview", title))
        for paragraph in section.get("paragraphs", []):
            if paragraph:
                add_body_paragraph(doc, paragraph)
    return doc


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


def rebuild_docx_from_output(output_dir: Path, numbering: DocxNumbering) -> Path:
    output_dir = Path(output_dir)
    os.environ["EIA_OUTPUT_DIR"] = str(output_dir)
    text_path = output_dir / "debug_tables" / "project_area_overview_text.json"
    if not text_path.exists():
        raise FileNotFoundError(f"缺少概况文本缓存: {text_path}")

    payload = json.loads(text_path.read_text(encoding="utf-8"))
    sections = payload.get("sections") if isinstance(payload, dict) else {}
    if not isinstance(sections, dict):
        sections = {}

    doc_payload = {
        "sections": [
            {"title": title, "paragraphs": sections.get(title) or []}
            for title in SECTION_TITLES
        ],
    }
    numbering.begin_section("overview", "项目区域环境概况")
    numbering.reset_section_heading_counters("overview")
    doc = build_docx(doc_payload, numbering)
    finalize_section_document(doc)
    out_path = output_dir / "project_area_overview.docx"
    doc.save(out_path)
    return out_path


if __name__ == "__main__":
    main()
