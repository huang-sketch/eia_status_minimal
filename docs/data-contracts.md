# 数据契约

本文记录 FastAPI Web 主流程当前实际使用的数据接口和 JSON 文件。契约分为三个稳定级别：

| 级别 | 含义 |
|---|---|
| 稳定 | 业务流程或外部调用依赖；只能向后兼容扩展 |
| 半稳定 | 当前前端或章节生成器直接读取；允许演进，但变更前必须同步调用方和测试 |
| 调试 | 用于诊断、审计或人工复核；除非升级稳定级别，否则不保证字段兼容 |

本文是 Markdown 契约，不代表项目已接入 JSON Schema 或运行时 schema 校验。

## 1. Web API

### `GET /api/health`

稳定接口。返回服务、磁盘、队列和限制信息。

```json
{
  "status": "ok",
  "disk": {"total_mb": 1000, "free_mb": 500},
  "jobs": {"running": 0, "current_job_id": null, "queued": 0, "queued_jobs": []},
  "limits": {"max_upload_mb": 30, "job_retention_count": 30, "job_retention_days": 7}
}
```

### `POST /api/jobs`

稳定接口。请求类型为 `multipart/form-data`。

| 字段 | 类型 | 必需 | 语义 |
|---|---|---:|---|
| `report_name` | string | 是 | 项目名称，也是最终报告文件名的基础 |
| `admin_division` | string | 是 | 项目所在行政区划，用于区域概况和地表水公报匹配 |
| `run_surface_water` | boolean | 否，默认 `true` | 是否运行地表水链 |
| `run_noise` | boolean | 否，默认 `true` | 是否运行声环境链 |
| `enable_llm_text_polish` | boolean | 否，默认 `false` | 是否尝试使用 LLM 润色规则文本 |
| `enable_llm_extraction` | boolean | 否，默认 `true` | 是否启用表头结构 LLM 兜底；不允许 LLM 决定监测值 |
| `data_source_type` | enum | 否，默认 `docx_report` | `docx_report`或 `xlsx_data` |
| `monitoring_report` | file | 是 | DOCX 模式为监测报告，XLSX 模式为监测数据工作簿 |
| `monitoring_plan` | file | 是 | DOCX 监测方案 |

约束：至少选择一个生成模块；数据文件扩展名必须与模式一致；方案必须为 `.docx`；单文件默认最大 30 MB，空文件被拒绝。

成功响应：

```json
{"job_id": "20260101_120000_ab12cd34", "status": "queued", "queue_position": 1}
```

### 任务查询接口

| 接口 | 稳定级别 | 返回内容 |
|---|---|---|
| `GET /api/jobs/{job_id}` | 稳定 | 状态、结果文件组、LLM/表头/输入校验状态和 `project_meta` |
| `GET /api/jobs/{job_id}/preview` | 半稳定 | 地表水、声环境结果表和达标预览 |
| `GET /api/jobs/{job_id}/logs` | 稳定 | UTF-8 纯文本任务日志 |
| `GET /api/jobs/{job_id}/download/{file_path}` | 稳定 | 下载任务 `output/`内文件；禁止目录穿越 |

不存在的任务返回 HTTP 404。上传类型、空文件和模块选择错误返回 HTTP 400，文件过大返回 HTTP 413。

## 2. 任务状态契约

### `input/project_meta.json`

稳定文件，由 `web_app.py`在任务创建时生成，任务恢复和结果命名会读取它。

| 字段 | 类型 | 说明 |
|---|---|---|
| `job_id` | string | 任务标识 |
| `report_name` | string | 前端项目名称 |
| `admin_division` | string | 行政区划 |
| `run_surface_water` | boolean | 是否运行地表水 |
| `run_noise` | boolean | 是否运行声环境 |
| `enable_llm_text_polish` | boolean | 是否请求正文润色 |
| `enable_llm_extraction` | boolean | 是否请求表头结构兜底 |
| `data_source_type` | enum | `docx_report`或 `xlsx_data` |
| `monitoring_data_original_name` | string | 用户上传的数据文件原名 |
| `created_at` | string | 本地 ISO 8601 时间，精确到秒 |

### `status.json`

稳定文件，由 `web_app.py`使用临时文件和 `os.replace`原子更新。前端通过 API 读取。

| 字段 | 类型/可空 | 说明 |
|---|---|---|
| `job_id` | string | 任务标识 |
| `status` | enum | `queued`、`running`、`success`、`failed` |
| `current_step` | string | 面向用户的当前步骤 |
| `error` | string/null | 失败原因 |
| `queue_position` | integer/null | 排队位置；运行时为 0 |
| `queued_at` | string | 入队时间 |
| `started_at` / `finished_at` | string，可选 | 开始和结束时间 |
| `updated_at` | string | 最近状态写入时间 |
| `queue_elapsed_seconds` / `run_elapsed_seconds` | number，可选 | 排队和运行耗时 |
| `result_groups` | object | 当前可下载结果分组 |
| `llm_text_polish` | object | 正文润色状态摘要 |
| `schema_fallback` | object | 表头规则/LLM 解析状态摘要 |
| `input_validation` | object | XLSX 输入校验摘要；DOCX 模式通常为 `not_run` |
| `final_report_filename` | string，可选 | 成功任务的最终 Word 文件名 |

```text
queued -> running -> success
                  -> failed
```

服务启动时，遗留 `running`任务被标记为 `failed`；遗留 `queued`任务按照原排队时间恢复。缺少 `project_meta.json`的排队任务不能恢复。

## 3. 输入文件契约

### DOCX 模式

- 监测方案和监测报告必须是可由 `python-docx`读取的 `.docx`。
- 表格按 Word 正文顺序解析；表格前最近的段落作为表题和上下文。
- 方案负责点位名称、位置、标准类别、监测因子和频次等计划信息。
- 报告负责日期、实测值、单位、检测机构和车流量等实测信息。
- 表头首先使用正式别名匹配；缺少必需字段时才允许 LLM 解释表头。

### XLSX 模式

- 监测方案仍为 DOCX，监测数据为 `.xlsx`。
- 工作簿需要包含所选模块对应的地表水和/或噪声工作表。
- 方案点位集合必须与 XLSX 点位集合一一对应；缺失、额外、重复或元数据冲突会阻断生成。
- 校验失败前仍会写出 `xlsx_input_validation.json`和 `point_correspondence.json`。

### 声环境编号

`NJX-Y-Z`中，`X`为敏感点序号，`Y`为敏感点中的位置序号，`Z`为某位置中的楼层序号而非实际楼层号。

| 位置/楼层关系 | 编号形式 |
|---|---|
| 单一位置、单一楼层 | `NJX` |
| 单一位置、多个楼层 | `NJX-Z`，例如 `NJ4-1`、`NJ4-2` |
| 多个位置、各位置单一楼层 | `NJX-Y` |
| 多个位置、某位置多个楼层 | `NJX-Y-Z` |

编号顺序由方案中的位置和楼层关系确定，XLSX 不能自行改变方案对应关系。

## 4. 稳定地表水业务 JSON

### `standard_config.json`

生成者：`surface_water_pipeline.py`。读取者：地表水达标计算和 Word 生成器。

```json
{
  "standard_name": "地表水环境质量标准 GB3838-2002",
  "points": {
    "WJ1": {
      "point_code": "WJ1",
      "river_name": "示例河",
      "standard_name": "地表水环境质量标准 GB3838-2002",
      "standard_class": "Ⅲ类",
      "section_name": "示例断面",
      "source_table": "plan.docx:table:0",
      "evidence": {},
      "limits": {}
    }
  },
  "warnings": [],
  "source_file": "plan.docx"
}
```

`points`以规范化点位编号为键。`limits`是该标准类别当前支持因子的限值映射；缺少标准类别时允许为空，但会进入人工复核状态。

### `monitoring_records.json`

生成者：`surface_water_pipeline.py`。读取者：地表水 Word 生成器和 Web 结果文件组。当前该文件只承载地表水记录，不是跨环境要素的统一记录库。

| 字段 | 类型/可空 | 说明 |
|---|---|---|
| `source_type` | string | 监测报告或 XLSX 监测数据 |
| `monitor_type` | string | 当前固定为 `surface_water` |
| `point_code` | string | 规范化点位编号，如 `WJ1` |
| `point` | string/null | 原始点位描述 |
| `sample_date` | string/null | 采样日期文本 |
| `factor` / `raw_factor` | string | 规范化和原始因子名称 |
| `value` | string | 原始结果文本，保留 `<`等符号 |
| `numeric_value` | number/null | 计算值；低于检出限时按当前规则取检出限的一半 |
| `unit` | string/null | 原始单位 |
| `sample_character` | string/null | 样品性状 |
| `evidence` | object/其他 | 来源行证据 |
| `source_file` / `source_table` | string/null | 来源记录；不能假设为可移植路径 |
| `source_row` | integer，可选 | XLSX 来源行号 |
| `needs_review` | boolean | 是否需要人工复核 |
| `warning` | string | 复核原因，无警告时为空字符串 |

### `compliance_results.json`

生成者：`surface_water_pipeline.py`。读取者：地表水结果表、结论和 Web 预览。它保留监测记录字段，并增加：

| 字段 | 类型/可空 | 说明 |
|---|---|---|
| `river_name` | string/null | 方案中的水体名称 |
| `standard_name` | string/null | 标准名称 |
| `standard_class` | string/null | Ⅰ至Ⅴ类之一 |
| `limit_value` | number/object/null | 单值限值或 pH 上下限对象 |
| `standard_index` | number/null | 四位小数的单项标准指数 |
| `is_compliant` | boolean/null | 达标、不达标或未完成判断 |
| `method` | enum | `ph_standard_index`、`do_standard_index`、`value_div_limit`、`not_applicable`、`invalid_value`或 `missing_standard` |

`水温`和 `悬浮物`当前为 `not_applicable`。溶解氧指数依赖同点位同日期水温，缺失时 `is_compliant`为 `null`并要求复核。

## 5. 半稳定前端结果 JSON

这些文件位于 `output/debug_tables/`。虽然目录名为 debug，但当前 Web 预览直接读取，变更必须同步修改 Web、前端和测试。

| 文件 | 主要顶层字段 | 使用方 |
|---|---|---|
| `surface_water_monitor_results_table.json` | `table_key`、`caption_suffix`、`title`、`headers`、`rows` | 地表水预览、Word 重建 |
| `surface_water_compliance_table.json` | 上述字段及 `evaluated_factors`、`review_cells` | 地表水达标预览、Word 重建 |
| `noise_sensitive_points_result_table.json` | 表元数据、`headers`、`rows`、`subtables`、`flow_label`、`warnings` | 声环境预览、Word 重建 |
| `traffic_noise_attenuation_table.json` | 表元数据、`headers`、`rows`、`flow_label`、`warnings` | 交通噪声预览、Word 重建 |
| `noise_compliance_summary.json` | `sensitive`、`attenuation`、`monitoring_meta` | 前端摘要、声环境结论 |

通用表外形：

```json
{
  "table_key": "domain_table_key",
  "caption_suffix": "表题",
  "title": "带编号表题",
  "headers": ["列1", "列2"],
  "rows": [{"列1": "示例", "列2": 1, "needs_review": false, "warning": ""}]
}
```

`headers`决定前端列顺序。前端最多预览前 50 行，完整内容通过下载接口获得。

噪声当前没有与地表水 `monitoring_records.json`等价的统一顶层正式记录文件。原始规则化数据主要存放在 `flattened_table_*.json`，前端和 Word 行为由上述结果表及摘要承担。调用方不得假设 `monitoring_records.json`包含噪声记录。

## 6. XLSX 校验契约

### `xlsx_input_validation.json`

半稳定文件，即使校验失败也会写出。

| 字段 | 类型 | 说明 |
|---|---|---|
| `source_file` / `plan_file` | string | XLSX 和 DOCX 来源记录 |
| `selected_modules` | object | `surface_water`、`noise`布尔值 |
| `sheets` | string[] | 工作表名称 |
| `noise_record_count` | integer，可选 | 解析出的噪声行数 |
| `surface_water_record_count` | integer，可选 | 地表水因子记录数 |
| `valid` | boolean | 是否允许继续生成 |
| `errors` | string[] | 阻断性错误 |
| `warnings` | string[] | 非阻断复核提示 |

Web API 会补充 `state`、`report_path`和 `correspondence_path`，但不写回原 JSON。

### `point_correspondence.json`

```json
{
  "noise": {
    "valid": true,
    "plan_codes": ["NJ1"],
    "data_codes": ["NJ1"],
    "missing_in_xlsx": [],
    "extra_in_xlsx": [],
    "errors": []
  },
  "surface_water": {
    "valid": true,
    "plan_codes": ["WJ1"],
    "data_codes": ["WJ1"],
    "missing_in_xlsx": [],
    "extra_in_xlsx": [],
    "errors": []
  }
}
```

未选择的模块可以不出现。

## 7. 生成者与读取者矩阵

| 数据/文件 | 生成者 | 主要读取者 | 稳定级别 | 失败行为 |
|---|---|---|---|---|
| `project_meta.json` | Web 创建任务 | 队列恢复、项目概况、报告命名 | 稳定 | 缺失时排队任务恢复失败 |
| `status.json` | Web 状态管理 | 任务查询 API、前端 | 稳定 | 无法解析时任务查询或恢复不可用 |
| `extraction/records.json` | DOCX 通用抽取 | 地表水适配器 | 调试 | 缺失时专用规则解析仍会运行 |
| `extraction/xlsx_surface_water_records.json` | XLSX 解析器 | 地表水管线 | 半稳定 | XLSX 模式缺失时任务失败 |
| `standard_config.json` | 地表水管线 | 达标判断、地表水 Word | 稳定 | 缺少标准时标记复核 |
| `monitoring_records.json` | 地表水管线 | 地表水 Word、下载 | 稳定 | 无记录时跳过地表水章节 |
| `compliance_results.json` | 地表水管线 | 地表水 Word、Web 预览 | 稳定 | 不能判定时使用 `null`并标记复核 |
| 噪声结果表和摘要 | 声环境生成器 | Web 预览、Word 重建 | 半稳定 | 必需扁平表缺失时声环境任务失败 |
| 地表水结果表 | 地表水生成器 | Web 预览、Word 重建 | 半稳定 | 缺失时预览为空或重建失败 |
| `docx_numbering_state.json` | 编号器 | 章节重建、合并 | 调试 | 每个任务开始时重置 |
| 最终 Word | Web 发布步骤 | 用户下载、ZIP | 稳定 | 合并失败时任务失败 |

## 8. 输出文件契约

章节顺序为项目区域概况、声环境、地表水，只对实际存在的章节连续编号。

| 文件 | 条件 | 说明 |
|---|---|---|
| `project_area_overview.docx` | LLM 概况生成成功 | 可选章节，失败不阻断主流程 |
| `noise_section.docx` | 选择声环境并成功生成 | 声环境独立章节 |
| `surface_water_section.docx` | 选择地表水且存在有效记录 | 地表水独立章节 |
| `<项目名称>现状调查与评价.docx` | 至少一个章节可合并 | 面向用户的最终报告 |
| `eia_outputs.zip` | 最终报告生成成功 | 包含最终 Word 和 JSON/debug 成果，不重复包含独立章节 Word |

最终文件名会去除输入末尾 `.docx`和重复后缀，替换 Windows 非法字符，项目名主体最长 100 个字符；空名称回退为“项目”。

## 9. 调试文件附录

以下文件默认不保证长期字段兼容：

| 分类 | 典型文件 |
|---|---|
| DOCX 分块和通用抽取 | `debug_chunks/*.json`、`extraction/eia_result.json`、`extraction/meta.json`、`extraction_logs/*.jsonl` |
| 噪声表预处理 | `original_table_*.json`、`flattened_table_*.json`、`noise_flattened_table_detection.json` |
| 表头映射 | `table_schema_detection.json`、`table_schema_llm_input.json`、`table_schema_llm_output.json`、`table_schema_validation.json` |
| 候选治理 | `table_schema_candidates.json`、`eia_router_diagnostics.json` |
| 影子分类 | `table_llm_classification.json`、`unclassified_candidate_tables.json` |
| LLM 正文 | `noise_llm_text_*.json`、`surface_water_llm_text_*.json` |
| 项目概况 | `project_area_overview_*.json` |
| 地表水公报 | `surface_water_local_status.json` |
| 正式文本校验 | `noise_formal_text_validation.json`、`surface_water_formal_text_validation.json`、`formal_text_validation.json` |
| 编码和编号 | `encoding_health_check.json`、`docx_numbering_state.json` |

`table_schema_candidates.json`是人工审核输入，不是正式配置。只有显式执行候选提升命令后，候选才可能进入正式别名。

## 10. 兼容与变更规则

1. 稳定对象可以新增可选字段，但不得在未迁移调用方时删除、改名或改变既有字段语义。
2. 半稳定对象变更时，必须同时更新 Web 预览、Word 重建逻辑和对应测试。
3. 状态、数据来源和达标方法新增枚举值时，前端必须能够容忍未知值。
4. 业务结论不得由 LLM 覆盖规则计算结果；LLM 校验失败必须回退规则文本。
5. 稳定契约变化必须同步更新本文、脱敏 fixture 和回归测试。
6. 真实输入、绝对业务路径、API key、运行日志和任务输出不得进入文档示例或 Git 提交。
