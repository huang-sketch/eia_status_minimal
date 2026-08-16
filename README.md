# 高速公路环评现状分析自动化系统

本项目是一个本地运行的 FastAPI 应用，用于从监测方案和监测数据中生成高速公路环评的声环境、地表水现状调查与评价章节。系统以确定性规则完成数据解析、方案与报告对应、标准判断和 Word 表格生成；大模型只用于表头解释、可选正文润色和可选项目区域环境概况。

当前没有 Streamlit 入口。Web 服务入口是 `web_app.py`，默认地址为 `http://127.0.0.1:8010`。

## 当前能力

### 已接入 Web 主流程

- 声环境监测方案解析、监测结果解析、昼夜平均值和超标量计算。
- 地表水监测方案解析、监测结果解析、GB 3838-2002 单项标准指数和达标判断。
- 两种监测数据来源：监测方案 DOCX + 监测报告 DOCX；监测方案 DOCX + 规则化监测数据 XLSX。
- XLSX 点位一一对应校验、声环境 `NJX-Y-Z`编号校验和输入错误阻断。
- 本地单进程 FIFO 队列、任务状态持久化、服务重启后的排队任务恢复。
- 声环境、地表水章节 Word 生成、章节重新编号、合并报告和 ZIP 打包。
- 最终报告名称使用前端项目名称，格式为 `<项目名称>现状调查与评价.docx`。

### 尚未形成完整业务链

- 环境空气只有通用识别和路由基础，没有专用标准判断、Word 章节或前端入口。
- 公报功能只有本地静态缓存和简单网页检索，没有公报文件归档、PDF 解析、数据库或来源版本管理。
- 当前没有知识库、向量库或 RAG。
- 内置任务队列只支持本地单机、单 Uvicorn worker，不适用于多 worker 或分布式部署。

## 处理流程

```text
前端项目信息和上传文件
  -> DOCX 分块或 XLSX 输入校验
  -> 方案表头与点位解析
  -> 监测值、日期、单位和车流量解析
  -> 确定性标准判断
  -> JSON 中间数据和校验数据
  -> 可选 LLM 正文润色
  -> 声环境/地表水 Word 章节
  -> 章节重新编号和合并
  -> <项目名称>现状调查与评价.docx + eia_outputs.zip
```

方案是点位名称、位置、标准类别和监测要求的主要依据；监测报告或 XLSX 是监测日期、实测值、单位和车流量的主要依据。LLM 不计算标准指数，也不决定最终达标结论。

## 安装

```powershell
python -m pip install -r requirements.txt
python -m pip install -r requirements-dev.txt
```

## 启动 Web 服务

```powershell
python web_app.py
```

默认监听 `http://127.0.0.1:8010`。Web 表单需要填写项目名称和行政区划，选择声环境、地表水生成模块，并选择一种输入模式：

| 输入模式 | 监测方案 | 监测数据 |
|---|---|---|
| Word 监测报告 | `.docx` | `.docx` |
| XLSX 监测数据 | `.docx` | `.xlsx` |

单个上传文件默认最大 30 MB。至少选择声环境或地表水中的一个模块。

## 独立 CLI

通用 DOCX 抽取入口仍可独立运行：

```powershell
python main.py report.docx plan.docx -o output/eia_result_test.json
```

该入口输出通用抽取记录和元数据，不等同于完整 Web 章节生成流程。Web 的 DOCX 地表水链会通过 `monitoring_extraction.py`调用通用抽取器，但会强制使用规则抽取；表头 LLM 兜底由后续专用映射器处理。

## LLM 职责

LLM 调用统一由 `llm_client.py`管理，使用 OpenAI 兼容的 Chat Completions 接口。

| 能力 | 默认状态 | 失败行为 |
|---|---|---|
| 表头结构兜底 | Web 默认勾选 | 无 API key 或调用失败时继续使用正式规则；无法满足必需字段时标记人工复核 |
| 声环境/地表水正文润色 | 默认关闭 | 回退到确定性规则文本，不改变监测值和结论 |
| 项目区域环境概况 | 有 API key 时尝试 | 作为可选步骤跳过，不阻断声环境和地表水报告 |
| 表格影子分类 | 默认关闭 | 仅生成诊断数据，不参与正式结果 |

LLM 表头映射只允许引用输入中真实存在的表头。运行时接受的新映射写入任务目录下的 `debug_tables/table_schema_candidates.json`，不会自动修改正式配置。

```powershell
# 只审核，不修改正式配置
python scripts/promote_schema_candidates.py --input <candidate-json>

# 显式应用无冲突候选
python scripts/promote_schema_candidates.py --input <candidate-json> --apply
```

## 主要环境变量

| 变量 | 默认值 | 说明 |
|---|---:|---|
| `EIA_WEB_PORT` | `8010` | Web 服务端口 |
| `EIA_WEB_RELOAD` | `false` | 是否启用 Uvicorn reload |
| `EIA_MAX_UPLOAD_MB` | `30` | 单个上传文件大小限制 |
| `EIA_JOB_RETENTION_COUNT` | `30` | 最多保留的已结束任务数 |
| `EIA_JOB_RETENTION_DAYS` | `7` | 已结束任务保留天数 |
| `EIA_QUEUE_POLL_SECONDS` | `2` | 队列轮询间隔 |
| `WEB_CONCURRENCY` / `UVICORN_WORKERS` | `1` | 必须为 1，否则服务启动失败 |
| `EIA_LLM_API_KEY` | 无 | LLM API key；未配置时规则主流程仍可运行 |
| `EIA_LLM_ENDPOINT` | `https://api.openai.com/v1/chat/completions` | OpenAI 兼容接口地址 |
| `EIA_LLM_MODEL` | 按 profile 选择 | 文本润色、概况或抽取模型 |
| `EIA_LLM_FAST_MODEL` | `qwen-flash` | 表头映射和快速分类模型 |
| `EIA_LLM_TIMEOUT_SECONDS` | `25` | 普通 LLM 请求超时秒数 |
| `EIA_LLM_TEXT_TIMEOUT_SECONDS` | `120` | 文本类 LLM 请求超时秒数 |
| `EIA_SURFACE_WATER_WEB_SEARCH` | `true` | 缓存未命中时是否检索地表水公报网页 |
| `EIA_SURFACE_WATER_WEB_TIMEOUT_SECONDS` | `12` | 公报网页请求超时秒数 |
| `ENABLE_LLM_TABLE_CLASSIFICATION_SHADOW` | `false` | 是否运行表格影子分类诊断 |

Web 会为子进程设置 `EIA_INPUT_DIR`、`EIA_OUTPUT_DIR`、`EIA_DATA_SOURCE_TYPE`、`EIA_RUN_NOISE`、`EIA_RUN_SURFACE_WATER`、`ENABLE_SCHEMA_FALLBACK`和 `ENABLE_LLM_TEXT_POLISH`，通常不需要手工配置。

## 目录与输出

```text
config/                  正式表头别名、Word 布局、润色规则和公报静态缓存
static/                  Web 前端
tests/                   脱敏回归测试
runs/web_jobs/<job_id>/  Web 任务输入、状态、日志和输出
scripts/                 人工治理命令
```

单个 Web 任务的主要输出：

```text
output/
  standard_config.json
  monitoring_records.json
  compliance_results.json
  debug_tables/
  project_area_overview.docx        # 可选
  noise_section.docx                # 选择声环境时
  surface_water_section.docx        # 选择地表水且存在数据时
  <项目名称>现状调查与评价.docx
  eia_outputs.zip
```

`input/`、`output/`和 `runs/`默认被 Git 忽略，不应提交真实任务输入和运行结果。详细字段、稳定级别和上下游关系见 [数据契约](docs/data-contracts.md)。

## 测试

```powershell
python -m pytest -q
```

测试覆盖 Word 分块、表头映射、地表水指数、声环境楼层/频次/车流量、XLSX 对应关系、任务队列、原子状态写入和脱敏离线生成流程。

## 兼容性原则

- 稳定业务字段只进行向后兼容扩展。
- 删除、改名或改变稳定字段语义时，必须同步更新数据契约和回归测试。
- `debug_tables`中的诊断文件默认不是长期稳定 API；前端正在读取的文件除外，具体分级见数据契约。
- 代码行为与文档冲突时，以经过测试的当前代码为事实来源，并及时修正文档。
