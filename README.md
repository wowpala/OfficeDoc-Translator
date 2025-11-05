# Use LLM to translate Office document

Support `pptx` and `docx`.

## Update 20250623

Add cache.

## Usage

### Apply an API Key
https://cloud.siliconflow.cn/account/ak

### Choose a free model
https://cloud.siliconflow.cn/models

`tencent/Hunyuan-MT-7B` or `Qwen/Qwen3-8B` is recommended.

### Install package and dependencies
```shell
pip install openai
pip install python-pptx
pip install python-docx
```

### Run
```shell
python OfficeDoc_Translator.py '.\input.pptx' zh
```

You can specify the language you want to translate to. For example:
`zh`, `en`, `zh-Hant`, `ja`, `ko`, `fr`, `es`, etc.