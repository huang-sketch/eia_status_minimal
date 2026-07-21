import re
import json
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional


SKILL_NAME = "highway-eia-status-writing"
SKILL_DISPLAY_NAME = "高速公路环评现状章节写作规范 Skill"

FLOOR_RE = re.compile(r"(?:\d+\s*[层楼]|顶层|底层|一层|二层|三层|四层|五层)")
POSITION_MARKER_RE = re.compile(r"(面向|背向|首排|第二排|本项目|临本项目|临路|道路|公路|铁路|室外|室内)")
ORIENTATION_RE = re.compile(
    r"(面向[^，,；;。]*?(?:首排|第二排|本项目|道路|公路|铁路)|"
    r"背向[^，,；;。]*?(?:首排|第二排|本项目|道路|公路|铁路)|"
    r"临本项目[^，,；;。]*|临路[^，,；;。]*|首排)"
)


def build_noise_formal_text_validation(
    plan_rows: List[Dict[str, Any]],
    monitor_points_table: Dict[str, Any],
    sensitive_results_table: Dict[str, Any],
    attenuation_table: Dict[str, Any],
    texts: Dict[str, Any],
    data_warnings: Optional[List[str]] = None,
) -> Dict[str, Any]:
    issues: List[Dict[str, Any]] = []
    output_text = collect_output_text(
        texts,
        monitor_points_table,
        sensitive_results_table,
        attenuation_table,
    )
    monitor_rows = {
        str(row.get("监测点编号") or "").strip(): row
        for row in monitor_points_table.get("rows", [])
        if isinstance(row, dict)
    }

    for plan in plan_rows:
        code = str(plan.get("point_code") or "").strip()
        if not code:
            continue
        row = monitor_rows.get(code)
        plan_name = clean_text(plan.get("point_name"))
        plan_position = clean_text(plan.get("position"))
        row_name = clean_text((row or {}).get("监测点名称"))
        row_position = clean_text((row or {}).get("监测点位置"))

        if row_name and has_position_detail(row_name):
            issues.append(
                make_issue(
                    "field_placement",
                    "warning",
                    "监测点名称疑似混入楼层、室内外或与项目关系，应归入监测点位置。",
                    point_code=code,
                    evidence=row_name,
                    recommendation="名称保留敏感点/监测对象，楼层、首排、面向本项目等写入监测点位置。",
                )
            )
        if re.fullmatch(r"\d+", row_position):
            issues.append(
                make_issue(
                    "floor_unit_normalization",
                    "warning",
                    "声环境监测点位置为纯数字，疑似楼层缺少“层”单位。",
                    point_code=code,
                    evidence=row_position,
                    recommendation="若该字段表示楼层，应统一写为“X层”；距离、桩号和衰减断面不应按楼层处理。",
                )
            )
        if plan_name and row_name and not names_compatible(plan_name, row_name):
            issues.append(
                make_issue(
                    "plan_report_priority",
                    "review",
                    "输出监测点名称与方案名称不一致，需确认是否误用报告长描述。",
                    point_code=code,
                    evidence=f"方案名称：{plan_name}；输出名称：{row_name}",
                    recommendation="监测点名称原则上以方案为准，报告点位文本用于匹配和补充实测信息。",
                )
            )
        for phrase in required_position_phrases(plan_position):
            if phrase and phrase not in row_position and phrase not in output_text:
                issues.append(
                    make_issue(
                        "missing_required_expression",
                        "warning",
                        "方案中的关键位置关系未在输出表格或正文中体现。",
                        point_code=code,
                        evidence=phrase,
                        recommendation="保留方案中的首排、面向本项目、临本项目侧、楼层和室内外信息。",
                    )
                )
        plan_floors = floor_tokens(plan_position)
        row_floors = floor_tokens(row_position)
        if len(plan_floors) > 1 and row_floors and len(row_floors) == 1:
            issues.append(
                make_issue(
                    "multi_floor_expression",
                    "review",
                    "方案包含多个楼层，但输出位置只体现了单个楼层。",
                    point_code=code,
                    evidence=f"方案位置：{plan_position}；输出位置：{row_position}",
                    recommendation="同一敏感点多楼层应保留楼层差异，避免合并为笼统点位。",
                )
            )

    for warning in data_warnings or []:
        issues.append(
            make_issue(
                "plan_report_consistency",
                "review",
                str(warning),
                recommendation="优先核对监测方案和监测报告是否为同一项目、同一批次、同一套点位。",
            )
        )

    return build_validation_payload(
        "noise",
        issues,
        checks=[
            "字段归位：监测点名称不应混入楼层、室内外、首排和面向本项目等位置描述。",
            "方案优先：点位名称、监测位置、执行标准原则上以方案为准。",
            "报告优先：监测时间、检测值、交通流量等实测信息以报告为准。",
            "输出校核：关键位置关系、多楼层表达和方案报告对应关系需显式提示。",
        ],
    )


def build_surface_water_formal_text_validation(
    monitor_points_table: Dict[str, Any],
    monitor_results_table: Dict[str, Any],
    compliance_table: Dict[str, Any],
    standard_config: Dict[str, Any],
    compliance_results: List[Dict[str, Any]],
    texts: Dict[str, Any],
) -> Dict[str, Any]:
    issues: List[Dict[str, Any]] = []
    output_text = collect_output_text(texts, monitor_points_table, monitor_results_table, compliance_table)
    rows = standard_config.get("points") or standard_config.get("point_configs") or []
    if isinstance(rows, dict):
        rows = list(rows.values())
    for item in rows if isinstance(rows, list) else []:
        code = clean_text(item.get("point_code") or item.get("code"))
        standard = clean_text(item.get("standard_class") or item.get("standard"))
        if code and not standard:
            issues.append(
                make_issue(
                    "missing_standard_class",
                    "review",
                    "监测方案中未识别该地表水点位的水质目标/标准类别。",
                    point_code=code,
                    recommendation="地表水达标判断应以方案水质目标或标准类别为依据，缺失时需人工补充。",
                )
            )

    for result in compliance_results:
        if not isinstance(result, dict):
            continue
        code = clean_text(result.get("point_code"))
        factor = clean_text(result.get("factor"))
        warning = clean_text(result.get("warning"))
        if result.get("needs_review") and warning:
            issues.append(
                make_issue(
                    "surface_water_review",
                    "review",
                    warning,
                    point_code=code,
                    evidence=factor,
                    recommendation="保留为人工校核项，不应将不确定计算结果写成确定性达标结论。",
                )
            )
        if factor == "溶解氧" and ("水温" in warning or "DO" in warning):
            issues.append(
                make_issue(
                    "do_temperature_dependency",
                    "warning",
                    "溶解氧标准指数计算依赖同点位、同日期水温，当前结果需校核。",
                    point_code=code,
                    recommendation="补齐水温或在正式文本中说明该项无法自动完成标准指数计算。",
                )
            )

    if compliance_table and "需校核" in output_text and "人工" not in output_text:
        issues.append(
            make_issue(
                "formal_expression",
                "info",
                "输出包含需校核项，建议在调试清单中明确人工确认责任。",
                recommendation="正式文本不直接写入调试语气，公测结果中保留校核清单。",
            )
        )

    return build_validation_payload(
        "surface_water",
        issues,
        checks=[
            "字段归位：断面/点位、河流名称、标准类别和监测结果分字段表达。",
            "方案优先：水质目标和标准类别以监测方案为准。",
            "报告优先：采样日期、监测值、单位和样品状态以监测报告为准。",
            "输出校核：缺少标准类别、异常值和溶解氧水温依赖需提示人工确认。",
        ],
    )


def merge_formal_text_validations(*payloads: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    sections = [payload for payload in payloads if isinstance(payload, dict)]
    issues: List[Dict[str, Any]] = []
    checks: List[str] = []
    for payload in sections:
        issues.extend(payload.get("issues") or [])
        checks.extend(payload.get("checks") or [])
    return {
        "skill": SKILL_NAME,
        "display_name": SKILL_DISPLAY_NAME,
        "valid": not any(issue.get("severity") in {"warning", "review", "error"} for issue in issues),
        "section_count": len(sections),
        "issue_count": len(issues),
        "sections": sections,
        "checks": unique(checks),
        "issues": issues,
    }


def write_formal_text_validation(debug_dir: Path, payload: Dict[str, Any]) -> None:
    debug_dir.mkdir(parents=True, exist_ok=True)
    section = str(payload.get("section") or "").strip()
    if section:
        section_path = debug_dir / f"{section}_formal_text_validation.json"
        section_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    aggregate_path = debug_dir / "formal_text_validation.json"
    existing_sections: List[Dict[str, Any]] = []
    if aggregate_path.exists():
        try:
            existing = json.loads(aggregate_path.read_text(encoding="utf-8"))
            if isinstance(existing, dict):
                existing_sections = [
                    item for item in existing.get("sections", [])
                    if isinstance(item, dict) and item.get("section") != section
                ]
        except json.JSONDecodeError:
            existing_sections = []
    aggregate = merge_formal_text_validations(*existing_sections, payload)
    aggregate_path.write_text(json.dumps(aggregate, ensure_ascii=False, indent=2), encoding="utf-8")


def build_validation_payload(section: str, issues: List[Dict[str, Any]], checks: List[str]) -> Dict[str, Any]:
    return {
        "skill": SKILL_NAME,
        "display_name": SKILL_DISPLAY_NAME,
        "section": section,
        "valid": not any(issue.get("severity") in {"warning", "review", "error"} for issue in issues),
        "issue_count": len(issues),
        "checks": checks,
        "issues": issues,
    }


def make_issue(
    category: str,
    severity: str,
    message: str,
    point_code: str = "",
    evidence: str = "",
    recommendation: str = "",
) -> Dict[str, str]:
    issue = {
        "category": category,
        "severity": severity,
        "message": message,
    }
    if point_code:
        issue["point_code"] = point_code
    if evidence:
        issue["evidence"] = evidence
    if recommendation:
        issue["recommendation"] = recommendation
    return issue


def collect_output_text(*items: Any) -> str:
    parts: List[str] = []
    for item in items:
        collect_text_parts(item, parts)
    return "\n".join(parts)


def collect_text_parts(item: Any, parts: List[str]) -> None:
    if item is None:
        return
    if isinstance(item, str):
        if item.strip():
            parts.append(item.strip())
        return
    if isinstance(item, dict):
        for value in item.values():
            collect_text_parts(value, parts)
        return
    if isinstance(item, list):
        for value in item:
            collect_text_parts(value, parts)
        return
    if isinstance(item, (int, float)):
        parts.append(str(item))


def has_position_detail(text: str) -> bool:
    value = clean_text(text)
    return bool(FLOOR_RE.search(value) or POSITION_MARKER_RE.search(value))


def required_position_phrases(text: str) -> List[str]:
    value = clean_text(text)
    phrases: List[str] = []
    orientation = ORIENTATION_RE.search(value)
    if orientation:
        phrases.append(normalize_phrase(orientation.group(1)))
    floor = FLOOR_RE.search(value)
    if floor:
        phrases.append(normalize_phrase(floor.group(0)))
    if "室外" in value:
        phrases.append("室外")
    if "室内" in value:
        phrases.append("室内")
    return unique([phrase for phrase in phrases if phrase])


def floor_tokens(text: str) -> List[str]:
    value = clean_text(text)
    tokens = []
    for match in FLOOR_RE.finditer(value):
        token = normalize_phrase(match.group(0))
        token = token.replace("楼", "层")
        tokens.append(token)
    return unique(tokens)


def names_compatible(plan_name: str, output_name: str) -> bool:
    plan = normalize_name(plan_name)
    output = normalize_name(output_name)
    if not plan or not output:
        return True
    return plan in output or output in plan


def normalize_name(text: str) -> str:
    value = clean_text(text)
    value = FLOOR_RE.sub("", value)
    value = POSITION_MARKER_RE.sub("", value)
    value = re.sub(r"(敏感点|监测点|点位编号|室外|室内)", "", value)
    value = re.sub(r"[，,；;。:：\\s()（）-]", "", value)
    return value


def normalize_phrase(text: str) -> str:
    return re.sub(r"\s+", "", clean_text(text)).replace("楼", "层")


def clean_text(value: Any) -> str:
    return str(value or "").strip()


def unique(values: Iterable[str]) -> List[str]:
    seen = set()
    result = []
    for value in values:
        if value in seen:
            continue
        seen.add(value)
        result.append(value)
    return result
