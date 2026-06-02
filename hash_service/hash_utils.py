import hashlib
import datetime
import os
import re
import json
import pytz
import nltk
from nltk.tokenize import sent_tokenize

# Проверка наличия punkt
try:
    nltk.data.find('tokenizers/punkt')
except LookupError:
    nltk.download('punkt', quiet=True)

# Импорт библиотек
try:
    from docx import Document
except ImportError:
    Document = None

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

try:
    from striprtf.striprtf import rtf_to_text
except ImportError:
    rtf_to_text = None


def calculate_sha256(filepath):
    """Вычисляет SHA256 хэш файла."""
    sha256_hash = hashlib.sha256()

    with open(filepath, "rb") as f:
        for byte_block in iter(lambda: f.read(4096), b""):
            sha256_hash.update(byte_block)

    return sha256_hash.hexdigest()


def detect_language(text):
    """
    Определяет язык текста.
    """
    if re.search(r'[а-яА-ЯёЁ]', text):
        return 'russian'
    return 'english'


def has_multiple_sentences(text):
    """
    Проверяет наличие минимум двух предложений.
    """
    if not text:
        return False

    text = text.strip()

    if len(text) < 50:
        return False

    try:
        lang = detect_language(text)

        sentences = sent_tokenize(text, language=lang)

        meaningful = [
            s.strip()
            for s in sentences
            if len(s.strip()) > 10
        ]

        return len(meaningful) >= 2

    except Exception:
        punctuation_count = len(
            re.findall(r'[.!?]', text)
        )

        return punctuation_count >= 2


def clean_text(text):
    """
    Очистка текста.
    """
    if not text:
        return ""

    text = text.replace('\xa0', ' ')
    text = re.sub(r'\s+', ' ', text)

    return text.strip()


def is_bad_block(text):
    """
    Отсеивает служебные блоки.
    """
    if not text:
        return True

    text_lower = text.lower().strip()

    bad_patterns = [
        r'^doi',
        r'^udk',
        r'^удк',
        r'^issn',
        r'^isbn',
        r'^keywords',
        r'^ключевые слова',
        r'^references',
        r'^список литературы',
        r'^author',
        r'^автор',
        r'^journal',
        r'^журнал',
        r'^abstract$',
        r'^аннотация$',
        r'^introduction$',
        r'^введение$'
    ]

    for pattern in bad_patterns:
        if re.search(pattern, text_lower):
            return True

    return False


def split_into_blocks(text):
    """
    Разбивает текст на абзацы.
    """
    blocks = re.split(r'\n\s*\n', text)

    cleaned_blocks = []

    for block in blocks:
        block = clean_text(block)

        if block:
            cleaned_blocks.append(block)

    return cleaned_blocks


def find_annotation(blocks):
    """
    Ищет первый информативный абзац.
    """

    for i, block in enumerate(blocks):

        if is_bad_block(block):
            continue

        if has_multiple_sentences(block):
            return block

        # Если встретили "Аннотация" или "Abstract",
        # берем следующий содержательный блок
        lower = block.lower().strip()

        if lower in ['abstract', 'аннотация']:

            for next_block in blocks[i + 1:]:

                if (
                    not is_bad_block(next_block)
                    and has_multiple_sentences(next_block)
                ):
                    return next_block

    return "Аннотация не найдена."


def extract_text_from_docx(filepath):
    doc = Document(filepath)

    paragraphs = []

    for p in doc.paragraphs:
        text = clean_text(p.text)

        if text:
            paragraphs.append(text)

    return '\n\n'.join(paragraphs)


def extract_text_from_txt(filepath):

    encodings = [
        'utf-8',
        'cp1251',
        'windows-1251',
        'latin-1'
    ]

    for enc in encodings:

        try:
            with open(filepath, 'r', encoding=enc) as f:
                return f.read()

        except UnicodeDecodeError:
            continue

    raise Exception("Не удалось определить кодировку TXT файла")


def extract_text_from_pdf(filepath):

    reader = PdfReader(filepath)

    text = ""

    pages_to_read = min(5, len(reader.pages))

    for i in range(pages_to_read):

        page = reader.pages[i]

        page_text = page.extract_text()

        if page_text:
            text += page_text + "\n\n"

    return text


def extract_text_from_rtf(filepath):

    with open(
        filepath,
        'r',
        encoding='latin-1',
        errors='ignore'
    ) as f:

        rtf_content = f.read()

    return rtf_to_text(rtf_content)


def get_annotation(filepath):
    """
    Извлекает аннотацию из документа.
    """

    extension = os.path.splitext(filepath)[1].lower()

    if os.name == 'nt':
        filepath = os.path.normpath(filepath)

    try:

        if extension == '.docx':

            if Document is None:
                return "python-docx не установлен"

            text = extract_text_from_docx(filepath)

        elif extension == '.txt':

            text = extract_text_from_txt(filepath)

        elif extension == '.pdf':

            if PdfReader is None:
                return "pypdf не установлен"

            text = extract_text_from_pdf(filepath)

        elif extension == '.rtf':

            if rtf_to_text is None:
                return "striprtf не установлен"

            text = extract_text_from_rtf(filepath)

        else:
            return f"Формат {extension} не поддерживается"

        blocks = split_into_blocks(text)

        annotation = find_annotation(blocks)

        return annotation

    except Exception as e:
        return f"Ошибка извлечения аннотации: {e}"


def process_document(filepath, output_folder="output"):
    """
    Основная функция обработки документа.
    """

    if not filepath or not os.path.exists(filepath):
        return {
            "error": "Файл не найден"
        }

    original_filename = os.path.basename(filepath)

    base_name = os.path.splitext(
        original_filename
    )[0]

    os.makedirs(output_folder, exist_ok=True)

    try:

        # Хэш
        doc_hash = calculate_sha256(filepath)

        # Время МСК
        moscow_tz = pytz.timezone('Europe/Moscow')

        now_msk = datetime.datetime.now(
            moscow_tz
        )

        date_str = now_msk.strftime(
            "%Y-%m-%d %H:%M:%S %Z%z"
        )

        # Аннотация
        annotation = get_annotation(filepath)

        # Количество предложений
        try:

            lang = detect_language(annotation)

            annotation_sentences = len(
                sent_tokenize(
                    annotation,
                    language=lang
                )
            )

        except Exception:
            annotation_sentences = 0

        # Размер
        hash_bytes = len(doc_hash) // 2

        annotation_bytes = len(
            annotation.encode('utf-8')
        )

        total_weight = (
            hash_bytes +
            annotation_bytes
        )

        result = {
            "filename": original_filename,
            "hash": doc_hash,
            "datetime_msk": date_str,
            "annotation": annotation,
            "annotation_length_chars": len(annotation),
            "annotation_sentences": annotation_sentences,
            "total_weight_bytes": total_weight,
            "processed_at": now_msk.isoformat()
        }

        # JSON
        json_filename = (
            f"{base_name}_result.json"
        )

        json_path = os.path.join(
            output_folder,
            json_filename
        )

        with open(
            json_path,
            'w',
            encoding='utf-8'
        ) as f:

            json.dump(
                result,
                f,
                ensure_ascii=False,
                indent=2
            )

        # TXT
        txt_filename = (
            f"{base_name}_result.txt"
        )

        txt_path = os.path.join(
            output_folder,
            txt_filename
        )

        with open(
            txt_path,
            'w',
            encoding='utf-8'
        ) as f:

            f.write(
                f"Файл: {original_filename}\n"
            )

            f.write(
                f"Дата обработки (МСК): "
                f"{date_str}\n"
            )

            f.write(
                f"SHA256 Хэш: "
                f"{doc_hash}\n"
            )

            f.write(
                f"Общий вес "
                f"(хэш + аннотация): "
                f"{total_weight} байт\n"
            )

            f.write("-" * 40 + "\n")

            f.write("АННОТАЦИЯ:\n")

            f.write(annotation)

        return result

    except Exception as e:

        return {
            "error": str(e)
        }

