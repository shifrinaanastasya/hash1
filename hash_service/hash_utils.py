import hashlib
import datetime
import os
import pytz
import re
import nltk
from nltk.tokenize import sent_tokenize

# Проверка и загрузка ресурсов NLTK при импорте модуля
try:
    nltk.data.find('tokenizers/punkt')
except LookupError:
    nltk.download('punkt', quiet=True)

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

def _is_meaningful_annotation(text):
    """
    Проверяет, является ли текст подходящей аннотацией.
    Исключает служебную информацию: авторов, издателей, DOI, годы и т.д.
    """
    if not text or len(text.strip()) < 20:
        return False
    
    clean_text = text.strip()
    
    # Списки стоп-слов и паттернов для отсева служебной информации
    skip_patterns = [
        r'^\s*([А-ЯA-Z][а-яa-z]+\s+){1,}[А-ЯA-Z][а-яa-z]+', # ФИО (несколько слов с заглавной)
        r'^\s*(Аннотация|Abstract|Анотація)', # Заголовки секций
        r'^\s*(Ключевые слова|Keywords|УДК|DOI|ISBN|ISSN)', # Мета-данные
        r'^\s*©', # Копирайт
        r'^\s*\d{4}\s*г\.?', # Год издания в начале
        r'^\s*[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}', # Email
        r'^\s*https?://', # Ссылки
        r'^\s*\[.*\]', # Ссылки в скобках типа [1]
        r'^\s*(Vol\.|No\.|Том|№)', # Номера томов/выпусков
        r'^\s*([А-ЯA-Z]\.)+\s*[А-ЯA-Z][а-яa-z]+', # Инициалы + Фамилия
    ]
    
    for pattern in skip_patterns:
        if re.search(pattern, clean_text, re.IGNORECASE):
            return False
            
    # Проверяем наличие хотя бы двух предложений
    try:
        sentences = sent_tokenize(clean_text, language='russian')
        meaningful_sentences = [s for s in sentences if len(s.strip()) > 10]
        if len(meaningful_sentences) < 2:
            # Попробуем английский, если русский не сработал (для смешанных текстов)
            sentences_en = sent_tokenize(clean_text, language='english')
            if len(sentences_en) < 2:
                return False
    except Exception:
        # Если токенизация не сработала, проверяем просто по точкам
        if clean_text.count('.') < 1:
            return False
            
    return True

def get_annotation(filepath):
    """
    Извлекает аннотацию (первый информативный абзац) из документа.
    Поддерживаются файлы .docx, .txt, .pdf и .rtf.
    Исправлена работа с кириллическими путями и кодировками.
    """
    # Получаем расширение корректно
    _, file_extension = os.path.splitext(filepath)
    file_extension = file_extension.lower()

    if file_extension == '.docx':
        try:
            doc = Document(filepath)
            for paragraph in doc.paragraphs:
                text = paragraph.text.strip()
                if _is_meaningful_annotation(text):
                    return text
            return "В документе .docx не найдено подходящей аннотации."
        except Exception as e:
            return f"Ошибка при чтении .docx: {str(e)}"
            
    elif file_extension == '.txt':
        try:
            # Пробуем разные кодировки для поддержки кириллицы
            encodings = ['utf-8', 'cp1251', 'latin-1']
            content = ""
            for enc in encodings:
                try:
                    with open(filepath, 'r', encoding=enc) as f:
                        content = f.read()
                    break
                except UnicodeDecodeError:
                    continue
            
            if not content:
                return "Не удалось определить кодировку файла .txt"
                
            lines = content.split('\n')
            for line in lines:
                text = line.strip()
                if _is_meaningful_annotation(text):
                    return text
            return "В документе .txt не найдено подходящей аннотации."
        except Exception as e:
            return f"Ошибка при чтении .txt: {str(e)}"

    elif file_extension == '.pdf':
        try:
            reader = PdfReader(filepath)
            full_text = ""
            # Читаем первые 5 страниц для поиска аннотации
            for i in range(min(len(reader.pages), 5)):
                page_text = reader.pages[i].extract_text()
                if page_text:
                    full_text += page_text + "\n\n"
            
            # Разбиваем на абзацы по двойному переносу строки или явно видимым блокам
            # Иногда в PDF абзацы разделены одиночным \n, но тогда строки короткие
            blocks = re.split(r'\n\s*\n', full_text)
            
            for block in blocks:
                text = block.strip()
                if _is_meaningful_annotation(text):
                    return text
            
            # Если строгая разбивка не помогла, попробуем по строкам, если они длинные
            if not blocks:
                 lines = full_text.split('\n')
                 for line in lines:
                     if _is_meaningful_annotation(line):
                         return line

            return "В документе .pdf не найдено подходящей аннотации."
        except Exception as e:
            return f"Ошибка при чтении .pdf: {str(e)}"

    elif file_extension == '.rtf':
        try:
            with open(filepath, 'r', encoding='latin-1', errors='ignore') as f:
                rtf_content = f.read()
            plain_text = rtf_to_text(rtf_content)
            
            blocks = re.split(r'\n\s*\n', plain_text)
            for block in blocks:
                text = block.strip()
                if _is_meaningful_annotation(text):
                    return text
                    
            return "В документе .rtf не найдено подходящей аннотации."
        except Exception as e:
            return f"Ошибка при чтении .rtf: {str(e)}"
    
    else:
        # Здесь была ошибка: если расширение пустое или не распознано
        if not file_extension:
            return "Файл не имеет расширения или имя файла повреждено."
        return f"Извлечение аннотации не поддерживается для файлов '{file_extension}'. Поддерживаются: .docx, .txt, .pdf, .rtf."