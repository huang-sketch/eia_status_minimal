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

未配置 `EIA_LLM_API_KEY` 时，程序会使用内置 mock 提取器，只用于演示流程和本地冒烟测试。

配置 API key 后，程序会调用 OpenAI 兼容的 chat completions 接口：

```powershell
$env:EIA_LLM_API_KEY="your-api-key"
$env:EIA_LLM_ENDPOINT="https://api.openai.com/v1/chat/completions"
$env:EIA_LLM_MODEL="gpt-4o-mini"
python main.py input.docx -o eia_result.json
```

变量说明：

- `EIA_LLM_API_KEY`: LLM API key。设置后启用 LLM 提取。
- `EIA_LLM_ENDPOINT`: OpenAI 兼容接口地址。未设置时使用 `https://api.openai.com/v1/chat/completions`。
- `EIA_LLM_MODEL`: 模型名称。未设置时使用 `gpt-4o-mini`。
