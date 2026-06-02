import hashlib
import datetime
import os
import pytz
import re
import nltk
from nltk.tokenize import sent_tokenize

# Попытка загрузки токенизатора при импорте модуля
try:
    nltk.data.find('tokenizers/punkt')
except LookupError:
    try:
        nltk.download('punkt', quiet=True)
        nltk.download('punkt_tab', quiet=True)
    except Exception:
        pass # Игнорируем ошибки загрузки при импорте, попробуем позже

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
    """Проверяет, содержит ли текст более одного предложения."""
    if not text:
        return False
    try:
        sentences = sent_tokenize(text, language='russian')
    except Exception:
        # Fallback для английского или простого разбиения, если русский не загружен
        try:
            sentences = sent_tokenize(text, language='english')
        except Exception:
            sentences = re.split(r'[.!?]+', text)
    
    meaningful_sentences = [s.strip() for s in sentences if len(s.strip()) > 10]
    return len(meaningful_sentences) >= 2

def _is_metadata_line(line):
    """Проверяет, является ли строка служебной информацией (метаданными)."""
    line = line.strip()
    if not line:
        return True
    
    # Паттерны для исключения служебных строк
    patterns = [
        r'^\d{4}$', # Только год
        r'^\d{4}-\d{2}-\d{2}$', # Дата
        r'^(УДК|UDC|DOI|ISBN|ISSN)\s*[:\.]?\s*\S+', # Библиографические индексы
        r'^https?://', # Ссылки
        r'^\S+@\S+\.\S+', # Email
        r'^(\([А-ЯA-Z][а-яa-z]+\s+[А-ЯA-Z][а-яa-z]+\)|[А-ЯA-Z][а-яa-z]+\s+[А-ЯA-Z]\.[А-ЯA-Z]\.)$', # Имена авторов (упрощенно)
        r'(аспирант|студент|доцент|профессор|кандидат|доктор)', # Ученые степени/звания
        r'(г\.|город|ул\.|улица|проспект|пер\.|переулок)', # Адреса
        r'^\d+\.\s+\d+$', # Номера страниц типа "1. 23"
        r'^\[\d+\]$', # Ссылки вида [1]
        r'^Vol\.\s*\d+|^No\.\s*\d+|^Issue\s*\d+', # Информация о выпуске журнала
    ]
    
    for pattern in patterns:
        if re.search(pattern, line, re.IGNORECASE):
            return True
            
    # Если строка слишком короткая (менее 15 символов) и не выглядит как начало предложения
    if len(line) < 15 and not line[0].isupper():
        return True
        
    return False

def get_annotation(filepath):
    """
    Извлекает аннотацию (первый информативный абзац) из документа.
    Игнорирует заголовки, авторов, издателей и другую служебную информацию.
    """
    file_extension = os.path.splitext(filepath)[1].lower()
    text_content = ""

    try:
        if file_extension == '.docx':
            doc = Document(filepath)
            paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
            # Объединяем для поиска, но сохраняем структуру для анализа
            for paragraph in paragraphs:
                if not _is_metadata_line(paragraph) and _has_multiple_sentences(paragraph):
                    return paragraph
            # Если не нашли по предложениям, вернем первый подходящий блок
            for paragraph in paragraphs:
                if not _is_metadata_line(paragraph) and len(paragraph) > 50:
                    return paragraph
                    
        elif file_extension == '.txt':
            # Пробуем разные кодировки для кириллицы
            encodings = ['utf-8', 'cp1251', 'utf-8-sig']
            for enc in encodings:
                try:
                    with open(filepath, 'r', encoding=enc) as f:
                        lines = f.readlines()
                    break
                except UnicodeDecodeError:
                    continue
            else:
                return "Не удалось прочитать файл в поддерживаемых кодировках."
            
            for line in lines:
                cleaned = line.strip()
                if not _is_metadata_line(cleaned) and _has_multiple_sentences(cleaned):
                    return cleaned
            # Fallback
            for line in lines:
                cleaned = line.strip()
                if not _is_metadata_line(cleaned) and len(cleaned) > 50:
                    return cleaned

        elif file_extension == '.pdf':
            reader = PdfReader(filepath)
            full_text = ""
            # Читаем первые 3 страницы
            for i in range(min(len(reader.pages), 3)):
                page_text = reader.pages[i].extract_text()
                if page_text:
                    full_text += page_text + "\n\n"
            
            blocks = full_text.split('\n\n')
            for block in blocks:
                cleaned = block.strip()
                # Разбиваем блок на строки, если он многострочный, и проверяем первую значимую
                if not _is_metadata_line(cleaned) and _has_multiple_sentences(cleaned):
                    return cleaned
            
            # Попытка найти аннотацию внутри больших блоков
            for block in blocks:
                lines = block.split('\n')
                for line in lines:
                    cleaned = line.strip()
                    if not _is_metadata_line(cleaned) and _has_multiple_sentences(cleaned):
                        return cleaned

        elif file_extension == '.rtf':
            with open(filepath, 'r', encoding='latin-1', errors='ignore') as f:
                rtf_content = f.read()
            plain_text = rtf_to_text(rtf_content)
            
            blocks = plain_text.split('\n\n')
            for block in blocks:
                cleaned = block.strip()
                if not _is_metadata_line(cleaned) and _has_multiple_sentences(cleaned):
                    return cleaned
                    
        else:
            return f"Формат '{file_extension}' не поддерживается."
            
        return "Аннотация не найдена (возможно, структура документа отличается от стандартной)."
        
    except Exception as e:
        return f"Ошибка при обработке файла: {str(e)}"

def process_document(filepath):
    """
    Основная функция обработки документа.
    Возвращает словарь с результатами: хэш, дата, аннотация, вес.
    """
    if not os.path.exists(filepath):
        return {"error": "Файл не найден"}

    # Вычисление хэша
    doc_hash = calculate_sha256(filepath)
    
    # Дата и время (МСК)
    moscow_tz = pytz.timezone('Europe/Moscow')
    now_msk = datetime.datetime.now(moscow_tz)
    date_str = now_msk.strftime("%Y-%m-%d %H:%M:%S %Z%z")
    
    # Аннотация
    annotation = get_annotation(filepath)
    
    # Расчет веса
    hash_bytes = len(doc_hash) // 2
    annotation_bytes = len(annotation.encode('utf-8')) if annotation else 0
    total_weight = hash_bytes + annotation_bytes
    
    return {
        "filename": os.path.basename(filepath),
        "hash": doc_hash,
        "date": date_str,
        "annotation": annotation,
        "total_weight_bytes": total_weight,
        "annotation_length_chars": len(annotation) if annotation else 0
    }