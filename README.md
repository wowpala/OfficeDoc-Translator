# OfficeDoc Translator

A tool to translate PowerPoint (`pptx`) and Word (`docx`) files using LLM.

## Features

- Translate PPTX and DOCX files
- Translation caching for faster repeated translations
- Customizable prompts via `llm_prompt.txt`
- Configurable via `.env` file
- Support for mutual translation between multiple languages
- Preserve original document format and layout

## Quick Start

### 1. Get an API Key

Obtain your API key from your chosen LLM service provider.

### 2. Set Up Environment

#### 2.1 Install Dependencies

Use `uv` to manage dependencies:

```powershell
uv sync
```

#### 2.2 Configure Environment Variables

Copy `.env.example` to `.env` and fill in your settings:

```powershell
Copy-Item .env.example .env
```

**`.env` Configuration:**

| Variable | Description | Default |
|----------|-------------|---------|
| `LLM_API_KEY` | Your API key | Required |
| `MODEL_NAME` | Translation model | `gemini-3-pro-high` |
| `ENDPOINT` | API endpoint | `https://api.siliconflow.cn/v1` |
| `TEMPERATURE` | Response creativity (0.0-2.0) | `1` |
| `ENABLE_THINKING` | Enable thinking mode | `false` |

#### 2.3 Customize Translation Prompt

Edit the `llm_prompt.txt` file to customize translation instructions. Use `{target_language}` as a placeholder for the target language.

## Usage

```powershell
python OfficeDoc_Translator.py <input_file> <target_language>
```

**Examples:**

```powershell
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

## Caching Mechanism

Translation results are cached in the `cache/` directory. The cache is automatically loaded and saved.

Use the `--no-cache` parameter to disable caching.

## Technical Details

- Uses OpenAI-compatible API format to call LLM
- Uses `requests` library for HTTP calls
- Supports custom API endpoints (configured via environment variables)
- Minimizes LLM API calls (through caching mechanism)

## Project Structure

```
OfficeDoc_Translator/
├── OfficeDoc_Translator.py  # Main script
├── llm_prompt.txt          # Translation prompt template
├── .env.example            # Environment variable example
├── .gitignore              # Git ignore file
├── .python-version         # Python version
├── pyproject.toml          # Project dependency configuration
└── README.md               # Project documentation
```

## Notes

- Please ensure your API key is securely stored in the `.env` file and not committed to version control
- Translation quality depends on the capabilities of the selected model
- Translation may take longer for large documents
- The caching mechanism stores translation results in the `cache/` directory; regular cleaning can save disk space
