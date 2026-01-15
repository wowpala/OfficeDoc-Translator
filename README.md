# Use LLM to translate Office documents

Translate PowerPoint (`pptx`) and Word (`docx`) files using LLM.

## Features

- Translate PPTX and DOCX files
- Translation caching for faster repeated translations
- Customizable prompts via `llm_prompt.txt`
- Configurable via `.env` file

## Setup

### 1. Apply an API Key

Get your API key from [SiliconFlow](https://cloud.siliconflow.cn/account/ak).

### 2. Choose a Model

Recommended models: `Qwen/Qwen3-8B`

Browse available models at [SiliconFlow Models](https://cloud.siliconflow.cn/models).

### 3. Install Dependencies

```shell
pip install openai python-pptx python-docx
```

### 4. Configure

Copy `.env.example` to `.env` and fill in your settings:

```shell
cp .env.example .env
```

**`.env` Configuration:**

| Variable | Description | Default |
|----------|-------------|---------|
| `LLM_API_KEY` | Your API key | Required |
| `MODEL_NAME` | Translation model | `Qwen/Qwen3-8B` |
| `ENDPOINT` | API endpoint | `https://api.siliconflow.cn/v1` |
| `TEMPERATURE` | Response creativity (0.0-2.0) | `0.7` |
| `ENABLE_THINKING` | Enable thinking mode | `false` |

**Customize Prompt:**

Edit `llm_prompt.txt` to modify the translation instructions. Use `{target_language}` as a placeholder for the target language.

## Usage

```shell
python OfficeDoc_Translator.py <input_file> <target_language>
```

**Examples:**

```shell
# Translate to Chinese
python OfficeDoc_Translator.py ./input.pptx zh-CN

# Translate to English
python OfficeDoc_Translator.py ./document.docx en-US

# Specify file type explicitly
python OfficeDoc_Translator.py ./file.pptx ja-JP --type ppt
```

**Arguments:**

| Argument | Description |
|----------|-------------|
| `input_file` | Input PPTX or DOCX file |
| `target_language` | Target language code (default: `zh-CN`) |
| `--type` | Force file type (`ppt` or `word`) |
| `--no-cache` | Disable translation cache |

**Supported Languages:**

`zh-CN`, `en-US`, `ja-JP`, `ko-KR`, `fr-FR`, `de-DE`, `es-ES`, etc.

## Cache

Translations are cached in the `cache/` directory. Cache is automatically loaded and saved.

Use `--no-cache` to disable caching.
