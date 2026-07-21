---
name: highway-eia-status-writing
description: 高速公路环评现状章节写作规范。Use when Codex or the EIA status agent needs to generate, review, or validate highway EIA status-analysis chapters, especially surface water, noise, regional environmental status, monitoring plan/report correspondence, field placement, standard-class usage, and formal-text gap checks.
---

# 高速公路环评现状章节写作规范

Use this skill to make generated highway EIA status-analysis chapters closer to formal report language.

Core workflow:

1. Treat the monitoring plan as the source of truth for point layout, point names, monitoring positions, and standard classes.
2. Treat the monitoring report as the source of truth for measured dates, values, units, traffic flow, and actual result records.
3. Split long point descriptions into point code, point name, floor/indoor-outdoor position, project-facing relationship, and report-side point code.
4. Keep names concise. Do not put floor, indoor/outdoor, first-row, or project-facing relationship into the monitoring point name.
5. Keep positions complete. Preserve phrases such as `面向本项目首排`, `临本项目侧`, `1层室外`, `顶层`, and similar formal monitoring-position details.
6. Let rules handle extraction, matching, calculation, and hard validation. Use LLMs only for understanding public text, organizing expression, and polishing.
7. Emit a validation checklist whenever the output still differs from formal EIA text or needs human confirmation.

For detailed rule groups, read `references/writing-rules.md`.
