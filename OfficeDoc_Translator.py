from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.chart.data import CategoryChartData

# Add docx library import
import docx
import re

# from pptx.enum.dml import MSO_THEME_COLOR
from openai import OpenAI
import os
import argparse
import signal
import sys
import httpx

# Add json and hashlib modules
import json
import hashlib


# Load config from .env file
ENV_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")

def load_env():
    config = {}
    if os.path.exists(ENV_PATH):
        with open(ENV_PATH, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    key, value = line.split("=", 1)
                    config[key.strip()] = value.strip()
    return config

env_config = load_env()

MODEL_NAME = env_config.get("MODEL_NAME", "Qwen/Qwen3-8B")
api_key = env_config.get("LLM_API_KEY", "")
ENDPOINT = env_config.get("ENDPOINT", "https://api.siliconflow.cn/v1")
TEMPERATURE = float(env_config.get("TEMPERATURE", "0.7"))
ENABLE_THINKING = env_config.get("ENABLE_THINKING", "false").lower() == "true"

# Load prompt template
PROMPT_TEMPLATE_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "llm_prompt.txt"
)
if os.path.exists(PROMPT_TEMPLATE_PATH):
    with open(PROMPT_TEMPLATE_PATH, "r", encoding="utf-8") as f:
        PROMPT_TEMPLATE = f.read()
else:
    PROMPT_TEMPLATE = (
        "You are a professional translator. Translate to {target_language}."
    )


def get_prompt(target_language):
    return PROMPT_TEMPLATE.format(target_language=target_language)


# Create httpx client without SSL verification
http_client = httpx.Client(verify=False)
client = OpenAI(api_key=api_key, base_url=ENDPOINT, http_client=http_client)

# Parse command line arguments
parser = argparse.ArgumentParser(description="Translate PowerPoint or Word files")
parser.add_argument("input_file", nargs="?", help="Input PPT or Word file")
parser.add_argument(
    "target_language",
    nargs="?",
    default="zh-CN",
    help="Target language code (default: zh-CN)",
)
parser.add_argument(
    "--type", choices=["ppt", "word"], help="Specify file type (ppt or word)"
)
parser.add_argument("--no-cache", action="store_true", help="Disable translation cache")
args = parser.parse_args()

# Create translation cache
translation_cache = {}
cache_file = None
cache_hit_count = 0  # Cache hit count

# Global cache: one cache file per target language


def init_cache():
    global translation_cache, cache_file
    if args.no_cache:
        return

    cache_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache")
    os.makedirs(cache_dir, exist_ok=True)

    # Cache per target language
    cache_filename = f"global-{args.target_language}.json"
    cache_file = os.path.join(cache_dir, cache_filename)

    # Load existing cache if exists
    if os.path.exists(cache_file):
        try:
            with open(cache_file, "r", encoding="utf-8") as f:
                translation_cache = json.load(f)
            print(
                f"Loaded {len(translation_cache)} translation cache entries (global, target language: {args.target_language})"
            )
        except Exception as e:
            print(f"Failed to load cache: {e}")
            translation_cache = {}


def save_cache():
    if args.no_cache or not cache_file:
        return

    try:
        with open(cache_file, "w", encoding="utf-8") as f:
            json.dump(translation_cache, f, ensure_ascii=False, indent=2)
        print(
            f"Saved {len(translation_cache)} translation cache entries (global, target language: {args.target_language})"
        )
    except Exception as e:
        print(f"Failed to save cache: {e}")


# Check if second argument is a language code or a file path
if args.target_language and (
    args.target_language.startswith(".\\") or args.target_language.startswith("./")
):
    # This might be a file path, not a language code
    input_file = args.target_language
    args.target_language = "zh-CN"  # Reset to default language
    args.input_file = input_file
    print(
        f"Warning: Argument '{input_file}' looks like a file path, not a language code. Using as input file with default Chinese translation."
    )

# Handle input file
if args.input_file:
    input_file = args.input_file
    print(f"Using specified input file: {input_file}")
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"File not found: {input_file}")
    # Determine file type by extension
    file_ext = os.path.splitext(input_file)[1].lower()
    if args.type:
        file_type = args.type
    else:
        if file_ext in [".pptx", ".ppt"]:
            file_type = "ppt"
        elif file_ext in [".docx", ".doc"]:
            file_type = "word"
        else:
            raise ValueError(f"Unsupported file type: {file_ext}")
else:
    # Keep original auto-find logic, add Word file support
    if args.type == "word":
        docx_files = [
            f for f in os.listdir(".") if f.endswith(".docx") or f.endswith(".doc")
        ]
        if not docx_files:
            raise FileNotFoundError(
                "No .docx or .doc files found in current directory."
            )
        input_file = docx_files[0]
        file_type = "word"
    else:  # Default to PPT or explicitly specified as PPT
        pptx_files = [
            f for f in os.listdir(".") if f.endswith(".pptx") or f.endswith(".ppt")
        ]
        if not pptx_files:
            raise FileNotFoundError(
                "No .pptx or .ppt files found in current directory."
            )
        input_file = pptx_files[0]
        file_type = "ppt"
    print(f"Auto-selected input file: {input_file}")

# Generate output filename
output_file = (
    os.path.splitext(input_file)[0]
    + f"-{args.target_language}"
    + os.path.splitext(input_file)[1]
)

# Default font
font_modified = "Microsoft YaHei"


def translate_text(text, target_language):
    if not text or len(text.strip()) < 2:
        return text

    # Check if translation already exists in cache
    cache_key = text.strip()
    global cache_hit_count
    if not args.no_cache and cache_key in translation_cache:
        cache_hit_count += 1
        print(f"[Cache hit] {cache_key[:40]}{'...' if len(cache_key)>40 else ''}")
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
                4. Do not translate the following terms:
                - IT terms;
                - Numerical digits;
                - Words beginning with 'Forti'；
                - Words in single quotes 'output, spoke, AI, Fabric, SD-WAN, SASE, ZTNA'. 
                5. Keep the original words unchanged which you can't recognize.
                6. Maintain the original formatting and punctuation as much as possible.
                7. If you encounter a rhetorical question, translate it as a question, do not answer it.""",
                },
                {"role": "user", "content": text},
            ],
            temperature=TEMPERATURE,
        )
        if ENABLE_THINKING:
            response = client.chat.completions.create(
                model=MODEL_NAME,
                messages=[
                    {
                        "role": "system",
                        "content": get_prompt(target_language),
                    },
                    {"role": "user", "content": text},
                ],
                temperature=TEMPERATURE,
                extra_body={"enable_thinking": ENABLE_THINKING},
            )
        else:
            response = client.chat.completions.create(
                model=MODEL_NAME,
                messages=[
                    {
                        "role": "system",
                        "content": get_prompt(target_language),
                    },
                    {"role": "user", "content": text},
                ],
                temperature=TEMPERATURE,
            )
        translated_text = response.choices[0].message.content.strip()

        # Only write to cache on cache miss
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

        # Clear existing runs
        for _ in range(len(paragraph.runs)):
            p = paragraph._element
            p.remove(p.r_lst[0])

        # Add new run with translated text
        new_run = paragraph.add_run(translated_text)
        safe_set_font(new_run)

        print(f"Original: {full_text[:50]}...")
        print(f"Translated: {translated_text[:50]}...")
    except Exception as e:
        print(f"Error translating paragraph: {e}")


def translate_word_table(table, target_language):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                translate_paragraph(paragraph, target_language)


def translate_docx(input_file, target_language, output_file):
    try:
        print(f"Translating Word document: {input_file}")
        doc = docx.Document(input_file)

        # Translate main document content
        for paragraph in doc.paragraphs:
            translate_paragraph(paragraph, target_language)

            # Translate tables
        for table in doc.tables:
            translate_word_table(table, target_language)

        # Translate header and footer
        for section in doc.sections:
            for header in section.header.paragraphs:
                translate_paragraph(header, target_language)
            for footer in section.footer.paragraphs:
                translate_paragraph(footer, target_language)

        # Translate document properties
        if hasattr(doc.core_properties, "title") and doc.core_properties.title:
            doc.core_properties.title = translate_text(
                doc.core_properties.title, target_language
            )
        if hasattr(doc.core_properties, "subject") and doc.core_properties.subject:
            doc.core_properties.subject = translate_text(
                doc.core_properties.subject, target_language
            )

        doc.save(output_file)
        print(f"Translated Word document saved to {output_file}")
    except Exception as e:
        print(f"Error processing Word document: {e}")


def signal_handler(sig, frame):
    if not args.no_cache:
        save_cache()  # Save cache before exit
    sys.exit(0)


signal.signal(signal.SIGINT, signal_handler)

# Translate based on file type
if __name__ == "__main__":
    # Initialize cache
    init_cache()

    if file_type == "ppt":
        if (
            args.target_language == "zh-CN" and len(sys.argv) <= 2
        ):  # Only filename or no args, use default Chinese
            print(f"Translating PPT file '{input_file}' to Chinese")
        else:
            print(f"Translating PPT file '{input_file}' to {args.target_language}")
        translate_pptx(
            input_file=input_file,
            target_language=args.target_language,
            output_file=output_file,
        )
    else:  # word
        if (
            args.target_language == "zh-CN" and len(sys.argv) <= 2
        ):  # Only filename or no args, use default Chinese
            print(f"Translating Word file '{input_file}' to Chinese")
        else:
            print(f"Translating Word file '{input_file}' to {args.target_language}")
        translate_docx(
            input_file=input_file,
            target_language=args.target_language,
            output_file=output_file,
        )

    # Save cache
    save_cache()
    # Print cache hit count
    if not args.no_cache:
        print(f"Cache hits this run: {cache_hit_count}")
