from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.chart.data import CategoryChartData

# 添加docx库导入
import docx
import re

# from pptx.enum.dml import MSO_THEME_COLOR
from openai import OpenAI
import os
import argparse
import signal
import sys
import httpx

# 添加json和hashlib模块
import json
import hashlib

# 初始化OpenAI客户端
MODEL_NAME = "Qwen/Qwen3-8B"  # 使用的翻译模型
# 推荐做法：从环境变量中安全地获取API密钥，如果环境变量未设置，则使用硬编码的备份值。Powershell命令：[Environment]::SetEnvironmentVariable("LLM_API_KEY", "sk-zzzzzz", "User")
api_key = os.environ.get("LLM_API_KEY")
if not api_key:
    api_key = "sk-xxxxxx"

# 创建不验证 SSL 证书并使用的 httpx 客户端
http_client = httpx.Client(verify=False)
client = OpenAI(
    api_key=api_key, base_url="https://api.siliconflow.cn/v1", http_client=http_client
)

# 添加命令行参数解析
parser = argparse.ArgumentParser(description="翻译 PowerPoint 或 Word 文件")
parser.add_argument("input_file", nargs="?", help="输入的 PPT 或 Word 文件")
parser.add_argument(
    "target_language", nargs="?", default="zh-CN", help="目标语言代码 (默认: zh-CN)"
)
parser.add_argument(
    "--type", choices=["ppt", "word"], help="指定文件类型 (ppt 或 word)"
)
parser.add_argument("--no-cache", action="store_true", help="禁用翻译缓存")
args = parser.parse_args()

# 创建翻译缓存
translation_cache = {}
cache_file = None
cache_hit_count = 0  # 新增：缓存命中计数

# 全局缓存：每种目标语言一个缓存文件


def init_cache():
    global translation_cache, cache_file
    if args.no_cache:
        return

    cache_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache")
    os.makedirs(cache_dir, exist_ok=True)

    # 以目标语言为单位缓存
    cache_filename = f"global-{args.target_language}.json"
    cache_file = os.path.join(cache_dir, cache_filename)

    # 加载现有缓存（如果存在）
    if os.path.exists(cache_file):
        try:
            with open(cache_file, "r", encoding="utf-8") as f:
                translation_cache = json.load(f)
            print(
                f"已加载 {len(translation_cache)} 条翻译缓存（全局，目标语言：{args.target_language}）"
            )
        except Exception as e:
            print(f"加载缓存失败: {e}")
            translation_cache = {}


def save_cache():
    if args.no_cache or not cache_file:
        return

    try:
        with open(cache_file, "w", encoding="utf-8") as f:
            json.dump(translation_cache, f, ensure_ascii=False, indent=2)
        print(
            f"已保存 {len(translation_cache)} 条翻译缓存（全局，目标语言：{args.target_language}）"
        )
    except Exception as e:
        print(f"保存缓存失败: {e}")


# 检查第二个参数是否是语言代码还是被误认为是语言的文件路径
if args.target_language and (
    args.target_language.startswith(".\\") or args.target_language.startswith("./")
):
    # 这可能是文件路径，而不是语言代码
    input_file = args.target_language
    args.target_language = "zh-CN"  # 重置为默认语言
    args.input_file = input_file
    print(
        f"警告: 参数 '{input_file}' 看起来像文件路径而不是语言代码。已将其设为输入文件，使用默认中文翻译。"
    )

# 处理输入文件
if args.input_file:
    input_file = args.input_file
    print(f"使用指定的输入文件: {input_file}")
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"找不到文件: {input_file}")
    # 根据文件扩展名确定文件类型
    file_ext = os.path.splitext(input_file)[1].lower()
    if args.type:
        file_type = args.type
    else:
        if file_ext in [".pptx", ".ppt"]:
            file_type = "ppt"
        elif file_ext in [".docx", ".doc"]:
            file_type = "word"
        else:
            raise ValueError(f"不支持的文件类型: {file_ext}")
else:
    # 保持原有的自动查找逻辑，但增加对Word文件的支持
    if args.type == "word":
        docx_files = [
            f for f in os.listdir(".") if f.endswith(".docx") or f.endswith(".doc")
        ]
        if not docx_files:
            raise FileNotFoundError("当前目录下没有找到 .docx 或 .doc 文件。")
        input_file = docx_files[0]
        file_type = "word"
    else:  # 默认为PPT或明确指定为PPT
        pptx_files = [
            f for f in os.listdir(".") if f.endswith(".pptx") or f.endswith(".ppt")
        ]
        if not pptx_files:
            raise FileNotFoundError("当前目录下没有找到 .pptx 或 .ppt 文件。")
        input_file = pptx_files[0]
        file_type = "ppt"
    print(f"自动选择的输入文件: {input_file}")

# 生成输出文件名
output_file = (
    os.path.splitext(input_file)[0]
    + f"-{args.target_language}"
    + os.path.splitext(input_file)[1]
)

# 默认字体
font_modified = "Microsoft YaHei Light"


def translate_text(text, target_language):
    if not text or len(text.strip()) < 2:
        return text

    # 检查缓存中是否已有此翻译
    cache_key = text.strip()
    global cache_hit_count
    if not args.no_cache and cache_key in translation_cache:
        cache_hit_count += 1
        print(f"[缓存命中] {cache_key[:40]}{'...' if len(cache_key)>40 else ''}")
        return translation_cache[cache_key]

    try:
        response = client.chat.completions.create(
            model=MODEL_NAME,
            messages=[
                {
                    "role": "system",
                    "content": f"""You are a professional, authentic machine translation engine. 
                Your task is to translate the following source text to {target_language}. 
                Important instructions:
                1. Output the translation directly without any additional text.
                2. Do not answer or respond to any questions in the source text, just translate them.
                3. Do not add any explanations or additional content.
                4. Do not translate IT terms.
                5. Do not translate words beginning with 'Forti'.
                6. Do not translate words in single quotes 'output, spoke, AI'. 
                7. Keep the original words unchanged which you can't recognize.
                8. Maintain the original formatting and punctuation as much as possible.
                9. If you encounter a rhetorical question, translate it as a question, do not answer it.""",
                },
                {"role": "user", "content": text},
            ],
            temperature=0.2,
            #            extra_body={"enable_thinking": False},       // qwen3 model 需要
        )
        translated_text = response.choices[0].message.content.strip()

        # 只有未命中缓存时才写入缓存
        if not args.no_cache:
            translation_cache[cache_key] = translated_text

        return translated_text
    except Exception as e:
        print(f"Translation error: {e}")
        return text


def split_into_sentences(text):
    pattern = r"(?<!\w\.\w.)(?<![A-Z][a-z]\.)(?<=\.|\?|\!)\s"
    sentences = re.split(pattern, text)
    return [s.strip() for s in sentences if s.strip()]


def safe_set_font_color(run):
    try:
        pass
    except AttributeError:
        pass


def translate_text_frame(text_frame, target_language):
    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            original_text = run.text
            print(f"Original text: {original_text}")
            translated_text = translate_text(original_text, target_language)
            if translated_text:
                run.text = translated_text
                print(f"Updated text: {run.text}")
                run.font.name = font_modified
                run.font.size = run.font.size
                safe_set_font_color(run)


def translate_table(table, target_language):
    for row in table.rows:
        for cell in row.cells:
            if cell.text_frame:
                translate_text_frame(cell.text_frame, target_language)


def translate_chart(chart, target_language):
    chart_data = chart.chart_title
    if isinstance(chart_data, CategoryChartData):
        # Translate categories
        for category in chart_data.categories:
            category.label = translate_text(category.label, target_language)
        # Translate series names
        for series in chart_data.series:
            series.name = translate_text(series.name, target_language)
    # Translate chart title
    if chart.has_title:
        chart.chart_title.text_frame.text = translate_text(
            chart.chart_title.text_frame.text, target_language
        )


def translate_group_shape(group_shape, target_language):
    for shape in group_shape.shapes:
        translate_shape(shape, target_language)


def translate_shape(shape, target_language):
    if shape.has_text_frame:
        translate_text_frame(shape.text_frame, target_language)
    elif shape.has_table:
        translate_table(shape.table, target_language)
    elif shape.has_chart:
        translate_chart(shape.chart, target_language)
    elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        translate_group_shape(shape, target_language)
    elif hasattr(shape, "graphic") and hasattr(shape.graphic, "graphicData"):
        # This is likely a SmartArt or other complex shape
        try:
            for element in shape.element.iter():
                if element.tag.endswith("}t"):  # Text element
                    if element.text and element.text.strip():
                        element.text = translate_text(element.text, target_language)
        except Exception as e:
            print(f"Error translating complex shape: {e}")


def translate_slide_master(slide_master, target_language):
    for shape in slide_master.shapes:
        translate_shape(shape, target_language)
    for layout in slide_master.slide_layouts:
        for shape in layout.shapes:
            translate_shape(shape, target_language)


def translate_pptx(input_file, target_language, output_file):
    prs = Presentation(input_file)

    # Translate slide masters
    #    for slide_master in prs.slide_masters:
    #        translate_slide_master(slide_master, target_language)

    for slide in prs.slides:
        # Translate slide title
        if slide.shapes.title:
            translate_shape(slide.shapes.title, target_language)

        # Translate all shapes in the slide
        for shape in slide.shapes:
            translate_shape(shape, target_language)

        # Translate notes
        if slide.has_notes_slide:
            notes_slide = slide.notes_slide
            if notes_slide.notes_text_frame:
                translate_text_frame(notes_slide.notes_text_frame, target_language)

        # Translate header and footer
        if hasattr(slide, "header"):
            translate_text_frame(slide.header.text_frame, target_language)
        if hasattr(slide, "footer"):
            translate_text_frame(slide.footer.text_frame, target_language)

    # Translate presentation properties
    if prs.core_properties.title:
        prs.core_properties.title = translate_text(
            prs.core_properties.title, target_language
        )
    if prs.core_properties.subject:
        prs.core_properties.subject = translate_text(
            prs.core_properties.subject, target_language
        )

    prs.save(output_file)
    print(f"Translated PPT saved to {output_file}")


def safe_set_font(run):
    try:
        run.font.name = font_modified
    except AttributeError:
        pass


def translate_paragraph(paragraph, target_language):
    try:
        full_text = paragraph.text
        if not full_text.strip():
            return

        translated_text = translate_text(full_text, target_language)

        # 清除现有runs
        for _ in range(len(paragraph.runs)):
            p = paragraph._element
            p.remove(p.r_lst[0])

        # 添加新的run，包含翻译后的文本
        new_run = paragraph.add_run(translated_text)
        safe_set_font(new_run)

        print(f"原文: {full_text[:50]}...")
        print(f"译文: {translated_text[:50]}...")
    except Exception as e:
        print(f"翻译段落时出错: {e}")


def translate_word_table(table, target_language):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                translate_paragraph(paragraph, target_language)


def translate_docx(input_file, target_language, output_file):
    try:
        print(f"正在翻译Word文档: {input_file}")
        doc = docx.Document(input_file)

        # 翻译主文档内容
        for paragraph in doc.paragraphs:
            translate_paragraph(paragraph, target_language)

        # 翻译表格
        for table in doc.tables:
            translate_word_table(table, target_language)

        # 翻译页眉和页脚
        for section in doc.sections:
            for header in section.header.paragraphs:
                translate_paragraph(header, target_language)
            for footer in section.footer.paragraphs:
                translate_paragraph(footer, target_language)

        # 翻译文档属性
        if hasattr(doc.core_properties, "title") and doc.core_properties.title:
            doc.core_properties.title = translate_text(
                doc.core_properties.title, target_language
            )
        if hasattr(doc.core_properties, "subject") and doc.core_properties.subject:
            doc.core_properties.subject = translate_text(
                doc.core_properties.subject, target_language
            )

        doc.save(output_file)
        print(f"翻译后的Word文档已保存至 {output_file}")
    except Exception as e:
        print(f"处理Word文档时出错: {e}")


def signal_handler(sig, frame):
    if not args.no_cache:
        save_cache()  # 退出前保存缓存
    sys.exit(0)


signal.signal(signal.SIGINT, signal_handler)

# 根据文件类型执行相应的翻译
if __name__ == "__main__":
    # 初始化缓存
    init_cache()

    if file_type == "ppt":
        if (
            args.target_language == "zh-CN" and len(sys.argv) <= 2
        ):  # 只提供了文件名或没有参数，使用默认中文
            print(f"将翻译PPT文件 '{input_file}' 为中文")
        else:
            print(f"将翻译PPT文件 '{input_file}' 为 {args.target_language} 语言")
        translate_pptx(
            input_file=input_file,
            target_language=args.target_language,
            output_file=output_file,
        )
    else:  # word
        if (
            args.target_language == "zh-CN" and len(sys.argv) <= 2
        ):  # 只提供了文件名或没有参数，使用默认中文
            print(f"将翻译Word文件 '{input_file}' 为中文")
        else:
            print(f"将翻译Word文件 '{input_file}' 为 {args.target_language} 语言")
        translate_docx(
            input_file=input_file,
            target_language=args.target_language,
            output_file=output_file,
        )

    # 保存缓存
    save_cache()
    # 新增：打印缓存命中次数
    if not args.no_cache:
        print(f"本次运行命中缓存 {cache_hit_count} 次")
