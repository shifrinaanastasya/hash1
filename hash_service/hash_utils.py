import hashlib
import datetime
import os
import pytz
import nltk
from nltk.tokenize import sent_tokenize

# Загрузка токенизатора Punkt для русского языка (потребуется один раз)
try:
    nltk.data.find('tokenizers/punkt')
except LookupError:
    nltk.download('punkt', quiet=True)

# Загрузка специфического ресурса 'punkt_tab' для русской токенизации, если он требуется
try:
    nltk.data.find('tokenizers/punkt_tab')
except LookupError:
    nltk.download('punkt_tab', quiet=True)

from docx import Document
from pypdf import PdfReader
from striprtf.striprtf import rtf_to_text


def calculate_sha256(filepath):
    """Вычисляет SHA256 хэш файла."""
    sha256_hash = hashlib.sha256()
    try:
        with open(filepath, "rb") as f:
            for byte_block in iter(lambda: f.read(4096), b""):
                sha256_hash.update(byte_block)
        return sha256_hash.hexdigest()
    except FileNotFoundError:
        return "Файл не найден"
    except Exception as e:
        return f"Ошибка при вычислении хэша: {e}"


def _has_multiple_sentences(text):
    """Проверяет, содержит ли текст более одного предложения, используя NLTK."""
    sentences = sent_tokenize(text, language='russian')
    meaningful_sentences = [s.strip() for s in sentences if len(s.strip()) > 5]
    return len(meaningful_sentences) >= 2


def get_annotation(filepath):
    """
    Извлекает аннотацию (первый абзац, который содержит два или более предложения)
    из документа. Поддерживаются файлы .docx, .txt, .pdf и .rtf.
    """
    file_extension = os.path.splitext(filepath)[1].lower()

    if file_extension == '.docx':
        try:
            doc = Document(filepath)
            for paragraph in doc.paragraphs:
                cleaned_text = paragraph.text.strip()
                if cleaned_text and _has_multiple_sentences(cleaned_text):
                    return cleaned_text
            return "В документе .docx не найдено абзацев с двумя или более предложениями."
        except Exception as e:
            return f"Ошибка при извлечении аннотации из .docx: {e}"
    elif file_extension == '.txt':
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                for line in f:
                    cleaned_line = line.strip()
                    if cleaned_line and _has_multiple_sentences(cleaned_line):
                        return cleaned_line
            return "В документе .txt не найдено строк с двумя или более предложениями."
        except Exception as e:
            return f"Ошибка при извлечении аннотации из .txt: {e}"
    elif file_extension == '.pdf':
        try:
            reader = PdfReader(filepath)
            # Собираем текст с нескольких первых страниц для поиска аннотации
            full_text = ""
            for i in range(min(len(reader.pages), 3)):
                full_text += reader.pages[i].extract_text() + "\n\n"

            # Разбиваем текст на "абзацы" (блоки текста, разделенные несколькими переносами строки)
            # и ищем первый подходящий
            text_blocks = full_text.split('\n\n')
            for block in text_blocks:
                cleaned_block = block.strip()
                if cleaned_block and _has_multiple_sentences(cleaned_block):
                    return cleaned_block
            return "В документе .pdf не найдено абзацев с двумя или более предложениями на первых страницах."
        except Exception as e:
            return f"Ошибка при извлечении аннотации из .pdf: {e}"
    elif file_extension == '.rtf':
        try:
            with open(filepath, 'r', encoding='latin-1') as f:
                rtf_content = f.read()
            plain_text = rtf_to_text(rtf_content)

            # Разбиваем текст на "абзацы" (блоки текста, разделенные несколькими переносами строки)
            text_blocks = plain_text.split('\n\n')
            for block in text_blocks:
                cleaned_block = block.strip()
                if cleaned_block and _has_multiple_sentences(cleaned_block):
                    return cleaned_block
            return "В документе .rtf не найдено абзацев с двумя или более предложениями."
        except Exception as e:
            return f"Ошибка при извлечении аннотации из .rtf: {e}"
    else:
        return f"Извлечение аннотации не поддерживается для файлов '{file_extension}'. В настоящее время поддерживаются только .docx, .txt, .pdf и .rtf."


def process_document(filepath):
    """
    Обрабатывает документ и возвращает результат в виде словаря.
    """
    document_hash = calculate_sha256(filepath)

    moscow_timezone = pytz.timezone('Europe/Moscow')
    calculation_datetime_moscow = datetime.datetime.now(moscow_timezone)
    calculation_date = calculation_datetime_moscow.strftime("%Y-%m-%d %H:%M:%S %Z%z")

    document_annotation = get_annotation(filepath)

    # Calculate total weight in bytes
    hash_bytes_length = len(document_hash) // 2
    annotation_bytes_length = len(document_annotation.encode('utf-8'))
    total_weight_bytes = hash_bytes_length + annotation_bytes_length

    return {
        'hash': document_hash,
        'datetime_msk': calculation_date,
        'annotation': document_annotation,
        'annotation_length_chars': len(document_annotation),
        'annotation_sentences': len(sent_tokenize(document_annotation, language='russian')),
        'total_weight_bytes': total_weight_bytes
    }
