# OfficeDoc Translator

使用 OpenAI-compatible API 翻译 PowerPoint (`.pptx`) 和 Word (`.docx`) 文档。

## Features

- 支持 PPTX / DOCX 翻译
- 使用 `requests` 调用自定义 endpoint
- 全局 cache 减少重复文本请求
- retry + exponential backoff + request interval，降低 429 / SSL EOF 影响
- resume state 支持失败后重跑，不必从头重新请求
- PPT 默认按段落翻译；样式差异明显时回退到 run 级，减少 API 调用次数

## Quick Start

### 1. Install dependencies

```powershell
uv sync
```

### 2. Configure `.env`

```powershell
Copy-Item .env.example .env
```

当前实际使用的 LLM API 推荐设置：

```dotenv
MODEL_NAME=qwen/qwen3-32b
ENDPOINT=https://api.groq.com/openai/v1
TEMPERATURE=0.7
ENABLE_THINKING=false
REQUEST_TIMEOUT_SECONDS=60
MAX_RETRIES=4
INITIAL_BACKOFF_SECONDS=1
MAX_BACKOFF_SECONDS=20
MIN_REQUEST_INTERVAL_SECONDS=0.75
CONSECUTIVE_FAILURE_LIMIT=3
SAVE_PROGRESS_EVERY_N=20
RESUME_ENABLED=true
```

主要配置项：

| Variable | Description | Default |
|----------|-------------|---------|
| `LLM_API_KEY` | API key | Required |
| `MODEL_NAME` | Translation model | `qwen/qwen3-32b` |
| `ENDPOINT` | OpenAI-compatible endpoint | `https://api.groq.com/openai/v1` |
| `TEMPERATURE` | Sampling temperature | `0.7` |
| `ENABLE_THINKING` | Provider-specific thinking toggle | `false` |
| `REQUEST_TIMEOUT_SECONDS` | Connect/read timeout | `60` |
| `MAX_RETRIES` | Retry count for 429 / 5xx / network errors | `4` |
| `INITIAL_BACKOFF_SECONDS` | First backoff delay | `1` |
| `MAX_BACKOFF_SECONDS` | Backoff cap | `20` |
| `MIN_REQUEST_INTERVAL_SECONDS` | Minimum gap between requests | `0.75` |
| `CONSECUTIVE_FAILURE_LIMIT` | Abort threshold after exhausted failures | `3` |
| `SAVE_PROGRESS_EVERY_N` | Cache/state checkpoint interval | `20` |
| `RESUME_ENABLED` | Enable resume state by default | `true` |

### 3. Customize prompt

编辑 `llm_prompt.txt`。其中 `{target_language}` 会在运行时替换。

## Usage

```powershell
uv run python OfficeDoc_Translator.py <input_file> <target_language>
```

示例：

```powershell
uv run python OfficeDoc_Translator.py .\input.pptx zh-CN
uv run python OfficeDoc_Translator.py .\document.docx en-US
uv run python OfficeDoc_Translator.py .\file.pptx ja-JP --type ppt
uv run python OfficeDoc_Translator.py .\input.pptx zh-CN --max-retries 6 --min-request-interval 1.2
uv run python OfficeDoc_Translator.py .\input.pptx zh-CN --no-resume --no-fail-fast
```

命令参数：

| Argument | Description |
|----------|-------------|
| `input_file` | 输入 PPTX 或 DOCX 文件 |
| `target_language` | 目标语言代码，默认 `zh-CN` |
| `--type` | 强制文件类型：`ppt` / `word` |
| `--no-cache` | 禁用持久化 cache |
| `--resume` / `--no-resume` | 开启或关闭 resume state |
| `--fail-fast` / `--no-fail-fast` | 失败后立即中止，或保留原文继续 |
| `--max-retries` | 覆盖 `.env` 中的最大重试次数 |
| `--min-request-interval` | 覆盖 `.env` 中的最小请求间隔 |

## Stability Notes

- `cache/` 只能减少重复文本请求，不能防止瞬时 burst。
- 当前 README 的推荐配置基于本地实际在用的 `Groq + qwen/qwen3-32b` 组合。
- `429 Too Many Requests` 主要靠去重、串行限速、读取 `Retry-After`、以及 exponential backoff 缓解。
- `SSLEOFError / UNEXPECTED_EOF_WHILE_READING` 属于 transient provider/network error，主要靠 `Session` 复用、timeout 和 retry 缓解。
- 当前只使用主 endpoint，不做自动 provider fallback。
- 默认启用 fail-fast：单个翻译单元在耗尽重试后会中止任务、保存 cache/state，并输出失败统计。

## Verification

```powershell
uv run python -m compileall -q .
uv run python -m unittest discover -v
```
