import hashlib
import datetime
import os
import re
import json
import pytz
import nltk
from nltk.tokenize import sent_tokenize

# Попытка загрузки токенизатора при импорте
try:
    nltk.data.find('tokenizers/punkt')
except LookupError:
    try:
        nltk.download('punkt', quiet=True)
        nltk.download('punkt_tab', quiet=True)
    except Exception:
        pass

# Импорт библиотек для работы с форматами
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

def _has_multiple_sentences(text):
    """Проверяет, содержит ли текст более одного предложения."""
    if not text or len(text) < 20:
        return False
    try:
        sentences = sent_tokenize(text, language='russian')
        meaningful_sentences = [s.strip() for s in sentences if len(s.strip()) > 10]
        return len(meaningful_sentences) >= 2
    except Exception:
        # Фоллбэк, если токенизация не сработала
        return '.' in text and text.count('.') > 1

def _is_metadata_line(line):
    """
    Проверяет, является ли строка служебной информацией (авторы, издательство, DOI и т.д.),
    а не частью аннотации.
    """
    if not line:
        return True
    
    line_lower = line.lower().strip()
    
    # Паттерны служебной информации
    patterns = [
        r'^\d{4}', # Начинается с года
        r'^удк\s', r'^doi:\s?', r'^isbn', r'^issn',
        r'^©', r'^все права защищены',
        r'^аннотация:', r'^abstract:', 
        r'^ключевые слова:', r'^keywords:',
        r'^автор(ы)?[:\.]', r'^author(s)?[:\.]',
        r'^редактор', r'^editor',
        r'^журнал', r'^journal', r'^вестник', r'^bulletin',
        r'^выпуск', r'^issue', r'^том', r'^vol',
        r'^стр\.', r'^p\.', r'^с\.\s*\d',
        r'^http', r'^www\.',
        r'^@', r'^email', r'^e-mail',
        r'^г\.\s*москва', r'^г\.\s*санкт-петербург',
        r'^издательство', r'^publisher',
        r'^лицензия', r'^license',
        r'^статья получена', r'^received', r'^принята к печати', r'^accepted'
    ]
    
    # Проверка на список авторов
    if re.match(r'^([А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.[А-ЯЁ]?\.?\s*,?\s*)+', line_lower):
        return True
        
    for pattern in patterns:
        if re.search(pattern, line_lower):
            return True
            
    # Если строка слишком короткая и похожа на заголовок или имя
    if len(line) < 15 and not any(c in line for c in '.!?'):
        if re.match(r'^[А-ЯЁ][а-яё]+(\s+[А-ЯЁ][а-яё]+)*$', line):
            return True

    return False

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

def get_annotation(filepath):
    """
    Извлекает аннотацию (первый информативный абзац) из документа по пути.
    Поддерживаются файлы .docx, .txt, .pdf и .rtf.
    Корректно работает с путями на кириллице.
    """
    file_extension = os.path.splitext(filepath)[1].lower()
    
    # Нормализация пути для Windows
    if os.name == 'nt':
        filepath = os.path.normpath(filepath)

    if file_extension == '.docx':
        if Document is None:
            return "Библиотека python-docx не установлена."
        try:
            doc = Document(filepath)
            for paragraph in doc.paragraphs:
                text = paragraph.text.strip()
                if text and not _is_metadata_line(text) and _has_multiple_sentences(text):
                    return text
            return "В документе .docx не найдено подходящей аннотации."
        except Exception as e:
            return f"Ошибка при чтении .docx: {e}"
            
    elif file_extension == '.txt':
        try:
            encodings = ['utf-8', 'cp1251', 'latin-1']
            content = None
            for enc in encodings:
                try:
                    with open(filepath, 'r', encoding=enc) as f:
                        content = f.readlines()
                    break
                except UnicodeDecodeError:
                    continue
            
            if content is None:
                return "Не удалось определить кодировку файла."

            for line in content:
                text = line.strip()
                if text and not _is_metadata_line(text) and _has_multiple_sentences(text):
                    return text
            return "В документе .txt не найдено подходящей аннотации."
        except Exception as e:
            return f"Ошибка при чтении .txt: {e}"
            
    elif file_extension == '.pdf':
        if PdfReader is None:
            return "Библиотека pypdf не установлена."
        try:
            reader = PdfReader(filepath)
            full_text = ""
            for i in range(min(len(reader.pages), 5)):
                page_text = reader.pages[i].extract_text()
                if page_text:
                    full_text += page_text + "\n\n"
            
            lines = full_text.split('\n')
            for line in lines:
                text = line.strip()
                if len(text) < 20:
                    continue
                if text.isdigit():
                    continue
                if not _is_metadata_line(text) and _has_multiple_sentences(text):
                    return text
            return "В документе .pdf не найдено подходящей аннотации."
        except Exception as e:
            return f"Ошибка при чтении .pdf: {e}"
            
    elif file_extension == '.rtf':
        if rtf_to_text is None:
            return "Библиотека striprtf не установлена."
        try:
            with open(filepath, 'r', encoding='latin-1', errors='ignore') as f:
                rtf_content = f.read()
            plain_text = rtf_to_text(rtf_content)
            
            blocks = plain_text.split('\n')
            for block in blocks:
                text = block.strip()
                if text and not _is_metadata_line(text) and _has_multiple_sentences(text):
                    return text
            return "В документе .rtf не найдено подходящей аннотации."
        except Exception as e:
            return f"Ошибка при чтении .rtf: {e}"
    else:
        return f"Формат '{file_extension}' не поддерживается."

def process_document(filepath, output_folder="output"):
    """
    Основная функция обработки файла по пути.
    
    Args:
        filepath: Путь к сохраненному файлу (строка)
        output_folder: Папка для сохранения результатов
        
    Returns:
        dict: результаты обработки или ошибка
    """
    if not filepath or not os.path.exists(filepath):
        return {"error": "Файл не найден или путь не указан"}

    original_filename = os.path.basename(filepath)
    base_name = os.path.splitext(original_filename)[0]
    
    os.makedirs(output_folder, exist_ok=True)
    
    try:
        # 1. Вычисляем хэш
        doc_hash = calculate_sha256(filepath)
        
        # 2. Получаем дату и время (МСК)
        moscow_tz = pytz.timezone('Europe/Moscow')
        now_msk = datetime.datetime.now(moscow_tz)
        date_str = now_msk.strftime("%Y-%m-%d %H:%M:%S %Z%z")
        
        # 3. Получаем аннотацию
        annotation = get_annotation(filepath)
        
        # 4. Считаем количество предложений
        try:
            sentences_count = len(sent_tokenize(annotation, language='russian')) if annotation else 0
        except:
            sentences_count = 0
            
        # 5. Считаем вес
        hash_bytes = len(doc_hash) // 2
        annotation_bytes = len(annotation.encode('utf-8')) if annotation else 0
        total_weight = hash_bytes + annotation_bytes
        
        result = {
            "filename": original_filename,
            "hash": doc_hash,
            "datetime_msk": date_str,
            "annotation": annotation,
            "annotation_length_chars": len(annotation) if annotation else 0,
            "annotation_sentences": sentences_count,
            "total_weight_bytes": total_weight,
            "processed_at": now_msk.isoformat()
        }
        
        # Сохраняем результат в JSON
        result_filename = f"{base_name}_result.json"
        result_path = os.path.join(output_folder, result_filename)
        
        with open(result_path, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
            
        # Сохраняем результат в TXT
        txt_filename = f"{base_name}_result.txt"
        txt_path = os.path.join(output_folder, txt_filename)
        
        with open(txt_path, 'w', encoding='utf-8') as f:
            f.write(f"Файл: {original_filename}\n")
            f.write(f"Дата обработки (МСК): {date_str}\n")
            f.write(f"SHA256 Хэш: {doc_hash}\n")
            f.write(f"Общий вес (хэш + аннотация): {total_weight} байт\n")
            f.write("-" * 40 + "\n")
            f.write("АННОТАЦИЯ:\n")
            f.write(f"{annotation}\n")
            
        return result
        
    except Exception as e:
        return {"error": str(e)}