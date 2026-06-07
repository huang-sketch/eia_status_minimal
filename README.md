# EIA Status Minimal

环评现状分析模块的最小版本。当前目标是从 `.docx` 文件中读取段落和表格，按切片提取现状监测记录，并输出 JSON。

## 安装依赖

建议在项目目录中创建并启用虚拟环境，然后安装依赖：

```powershell
pip install -r requirements.txt
```

## 输入文件

将需要分析的 Word 文件放到项目目录，或在运行命令中传入文件的完整路径。

当前仅支持 `.docx` 文件。

## 运行

使用本地 mock 流程运行：

```powershell
python main.py input.docx -o eia_result.json
```

也可以一次传入多个文件：

```powershell
python main.py report.docx plan.docx -o eia_result.json
```

输出文件默认为 `eia_result.json`，内容包括：

- `records`: 提取到的监测记录
- `meta`: 输入文件、切片数量、记录数量

## LLM 环境变量

LLM 调用已统一到 [llm_client.py](llm_client.py)。未配置 `EIA_LLM_API_KEY` 时，规则路径仍可运行；需要 LLM 的功能会回退或跳过。

配置 API key 后，程序会调用 OpenAI 兼容的 chat completions 接口：

```powershell
$env:EIA_LLM_API_KEY="your-api-key"
$env:EIA_LLM_ENDPOINT="https://api.openai.com/v1/chat/completions"
$env:EIA_LLM_MODEL="gpt-4o-mini"
python main.py input.docx -o eia_result.json
```

### 公共变量

| 变量 | 说明 |
|------|------|
| `EIA_LLM_API_KEY` | LLM API key |
| `EIA_LLM_ENDPOINT` | OpenAI 兼容接口地址，默认 `https://api.openai.com/v1/chat/completions` |
| `EIA_LLM_TIMEOUT_SECONDS` | 单次请求超时（秒），默认 `25` |
| `EIA_LLM_MAX_RETRIES` | 失败重试次数，默认 `1` |
| `EIA_LLM_DISABLE` | 设为 `1` 可禁用 LLM（CLI 抽取路径） |

### 场景 Profile 默认值

| Profile | 用途 | 默认 model | temperature | max_tokens 环境变量 |
|---------|------|------------|-------------|---------------------|
| `extraction` | CLI 监测数据抽取（`main.py`） | `gpt-4o-mini` | 0 | `EIA_LLM_MAX_TOKENS`（默认 2048） |
| `text_polish` | 地表水/声环境 Word 文本润色 | `qwen-plus` | 0.2 | `EIA_LLM_TEXT_MAX_TOKENS`（默认 1800） |
| `project_overview` | 项目区域环境概况 | `qwen-plus` | 0.2 | `EIA_PROJECT_OVERVIEW_MAX_TOKENS`（默认 3500） |
| `fast` | 复杂噪声表分片抽取 | `qwen-flash` | 0 | `EIA_LLM_FAST_MAX_TOKENS`（默认 4096） |

各 Profile 的 model 仍可通过 `EIA_LLM_MODEL` 或 `EIA_LLM_FAST_MODEL` 覆盖。

### Web 服务

```powershell
python web_app.py
```

默认监听 `http://127.0.0.1:8010`，可通过 `EIA_WEB_PORT` 修改端口。

地表水任务流程：

1. `monitoring_extraction.py` — CLI 抽取（方案 + 报告），输出到 `output/extraction/`
2. `surface_water_pipeline.py` — CLI 记录为主、规则解析 fallback，写入 `monitoring_records.json`
3. `surface_water_section_generator.py` — 生成 Word 章节

Web 表单选项：

- **启用 LLM 表格抽取**：对应 `ENABLE_LLM_EXTRACTION=true`，需配置 `EIA_LLM_API_KEY`；未勾选时仍运行 CLI 抽取，但仅使用规则 fallback
- **启用 LLM 文本润色**：对应 `ENABLE_LLM_TEXT_POLISH=true`

Web 抽取相关环境变量：

| 变量 | 说明 |
|------|------|
| `ENABLE_LLM_EXTRACTION` | Web 表单「启用 LLM 表格抽取」 |
| `EIA_MAX_CHUNKS_PER_RUN` | Web 任务默认可处理 chunk 数，默认 `100`（CLI `main.py` 默认 `20`） |
| `EIA_OUTPUT_DIR/extraction/` | 存放 `eia_result.json`、`records.json`、`meta.json` |
| `EIA_OUTPUT_DIR/debug_tables/extraction_summary.json` | CLI/规则合并统计 |
