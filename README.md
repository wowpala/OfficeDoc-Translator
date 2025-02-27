# LLM-PPT-Translator

### 申请 API Key
https://cloud.siliconflow.cn/account/ak

### 选择免费模型
https://siliconflow.cn/zh-cn/models

推荐`Qwen/Qwen2.5-7B-Instruct`

### 安装依赖包
```shell
pip install openai
pip install python-pptx
```

### 使用方法
```shell
python PPT_Translator_siliconflow.py zh-CN '.\input.pptx'
```
支持指定语言，如`zh-CN`, `en-US`, `ja-JP`, `ko-KR`, `fr-FR`, `de-DE`, `es-ES`等