# Use LLM to translate Office document

Support `pptx` and `docx`.

## Update 20250623

Add cache.

## Usage

### Apply an API Key
https://cloud.siliconflow.cn/account/ak

### Choose a free model
https://cloud.siliconflow.cn/models

`Qwen/Qwen2.5-7B-Instruct` or `Qwen/Qwen3-8B` is recommended.

### Install package and dependencies
```shell
pip install openai
pip install python-pptx
pip install python-docx
```

### Run
```shell
python OfficeDoc_Translator.py '.\input.pptx' zh-CN
```

You can specify the language you want to translate to. For example:
`zh-CN`, `en-US`, `ja-JP`, `ko-KR`, `fr-FR`, `de-DE`, `es-ES`, etc.