import json
from pathlib import Path
from typing import Any, Dict


def load_text_polish_guidance(path: Path) -> Dict[str, Any]:
    if not path.exists():
        return {}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {"config_error": f"{path} 不是合法 JSON，已忽略该配置。"}
    if not isinstance(payload, dict):
        return {"config_error": f"{path} 顶层必须是 JSON 对象，已忽略该配置。"}
    return payload


def build_text_polish_prompt(
    payload: Dict[str, Any],
    *,
    role: str,
    output_keys: str,
    extra_rules: str = "",
) -> str:
    rules = [
        f"你是环评报告{role}文字润色助手。只润色文字，不修改任何数字、表号、点位/断面编号、监测日期、标准名称或达标/超标结论。",
        "严格遵循 text_guidance 中的 style_exemplars、length_targets、section_points 和 hard_constraints。",
        "style_exemplars.example 仅作句式、长度和语气参考，不得照搬其中的项目名、点位、日期、河流、道路和数值。",
        "所有事实数据以 rule_texts 和 summary 为准；二者与 exemplar 冲突时，必须服从 rule_texts 和 summary。",
        f"输出必须是 JSON 对象，键名必须完整保留：{output_keys}。",
        "禁止输出 Markdown；不得编造输入中不存在的事实。",
    ]
    if extra_rules.strip():
        rules.append(extra_rules.strip())
    rules.append(f"输入数据：\n{json.dumps(payload, ensure_ascii=False)}")
    return "\n".join(rules)
