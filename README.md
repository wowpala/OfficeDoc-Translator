# OfficeDoc Translator

Translate PowerPoint (`.pptx`) and Word (`.docx`) documents using an OpenAI-compatible API.

## Features

- PPTX / DOCX translation support
- Uses `requests` to call a custom endpoint
- Global cache to avoid duplicate text requests
- Retry + exponential backoff + request interval to mitigate 429 / SSL EOF errors
- Resume state — rerun after failure without starting over
- PPT translates by paragraph by default; falls back to run-level when style differences are significant, reducing API calls

## Quick Start

### 1. Install dependencies

```powershell
uv sync
```

### 2. Configure `.env`

```powershell
Copy-Item .env.example .env
```

Groq LLM API is recommended: get your key at [console.groq.com/keys](https://console.groq.com/keys), model `qwen/qwen3-32b` is suggested.

```dotenv
MODEL_NAME=qwen/qwen3-32b
ENDPOINT=https://api.groq.com/openai/v1
TEMPERATURE=0.7
REQUEST_TIMEOUT_SECONDS=60
MAX_RETRIES=4
INITIAL_BACKOFF_SECONDS=1
MAX_BACKOFF_SECONDS=20
MIN_REQUEST_INTERVAL_SECONDS=0.75
CONSECUTIVE_FAILURE_LIMIT=3
SAVE_PROGRESS_EVERY_N=20
RESUME_ENABLED=true
```

Configuration reference:

| Variable | Description | Default |
|----------|-------------|---------|
| `LLM_API_KEY` | API key | Required |
| `MODEL_NAME` | Translation model | `qwen/qwen3-32b` |
| `ENDPOINT` | OpenAI-compatible endpoint | `https://api.groq.com/openai/v1` |
| `TEMPERATURE` | Sampling temperature | `0.7` |
| `MAX_TOKENS` | Max output tokens per request, keep below the provider's OTPM limit | `900` |
| `REQUEST_TIMEOUT_SECONDS` | Connect/read timeout | `60` |
| `MAX_RETRIES` | Retry count for 429 / 5xx / network errors | `4` |
| `INITIAL_BACKOFF_SECONDS` | First backoff delay | `1` |
| `MAX_BACKOFF_SECONDS` | Backoff cap | `20` |
| `MIN_REQUEST_INTERVAL_SECONDS` | Minimum gap between requests | `0.75` |
| `CONSECUTIVE_FAILURE_LIMIT` | Abort threshold after exhausted failures | `3` |
| `SAVE_PROGRESS_EVERY_N` | Cache/state checkpoint interval | `20` |
| `RESUME_ENABLED` | Enable resume state by default | `true` |

### 3. Customize prompt

Edit `llm_prompt.txt`. The `{target_language}` placeholder is replaced at runtime.

## Usage

```powershell
uv run python OfficeDoc_Translator.py <input_file> <target_language>
```

Examples:

```powershell
uv run python OfficeDoc_Translator.py .\input.pptx zh-CN
uv run python OfficeDoc_Translator.py .\document.docx en-US
uv run python OfficeDoc_Translator.py .\file.pptx ja-JP --type ppt
uv run python OfficeDoc_Translator.py .\input.pptx zh-CN --max-retries 6 --min-request-interval 1.2
uv run python OfficeDoc_Translator.py .\input.pptx zh-CN --no-resume --no-fail-fast
```

Arguments:

| Argument | Description |
|----------|-------------|
| `input_file` | Input PPTX or DOCX file |
| `target_language` | Target language code, defaults to `zh-CN` |
| `--type` | Force file type: `ppt` / `word` |
| `--no-cache` | Disable persistent cache |
| `--resume` / `--no-resume` | Enable or disable resume state |
| `--fail-fast` / `--no-fail-fast` | Abort on failure vs keep original text and continue |
| `--max-retries` | Override max retries from `.env` |
| `--min-request-interval` | Override min request interval from `.env` |

## Stability Notes

- `cache/` only reduces duplicate text requests, cannot prevent instantaneous bursts.
- `429 Too Many Requests` is mitigated by dedup, serial rate limiting, reading `Retry-After`, and exponential backoff.
- `SSLEOFError / UNEXPECTED_EOF_WHILE_READING` are transient provider/network errors, mitigated by `Session` reuse, timeout, and retry.
- Only the primary endpoint is used; no automatic provider fallback.
- Thinking is always disabled: endpoints containing `groq.com` get `reasoning_effort: "none"`, all other endpoints get `thinking: {"type": "disabled"}`.
- Beyond that field the client does no provider-specific adaptation and no "retry with another field when rejected" downgrade; multi-channel routing and field translation behind a relay gateway are handled on the gateway side.
- Fail-fast is enabled by default: a single translation unit aborts the task after exhausting retries, saves cache/state, and outputs failure statistics.

## Verification

```powershell
uv run python -m compileall -q .
uv run python -m unittest discover -v
```
