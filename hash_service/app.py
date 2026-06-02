import os
import json
from flask import Flask, request, render_template, send_from_directory, redirect, url_for, flash
from werkzeug.utils import secure_filename
from hash_utils import process_document

app = Flask(__name__)
app.secret_key = 'super-secret-key-change-in-production'

# Конфигурация
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'output')
ALLOWED_EXTENSIONS = {'docx', 'txt', 'pdf', 'rtf'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # Максимальный размер файла 50MB

# Убедимся, что папки существуют
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('Файл не найден в запросе')
        return redirect(url_for('index'))

    file = request.files['file']

    if file.filename == '':
        flash('Файл не выбран')
        return redirect(url_for('index'))

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        try:
            # Обрабатываем документ
            result = process_document(filepath)
            result['filename'] = filename

            # Сохраняем результат в output папку
            base_name = os.path.splitext(filename)[0]
            output_filename = f"{base_name}_hash_annotation.json"
            output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)

            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(result, f, ensure_ascii=False, indent=2)

            # Также создаем текстовый файл с читаемым выводом
            txt_output_filename = f"{base_name}_hash_annotation.txt"
            txt_output_path = os.path.join(app.config['OUTPUT_FOLDER'], txt_output_filename)

            with open(txt_output_path, 'w', encoding='utf-8') as f:
                f.write(f"Файл: {filename}\n")
                f.write(f"SHA256 Хэш: {result['hash']}\n")
                f.write(f"Дата вычисления (МСК): {result['datetime_msk']}\n")
                f.write(f"\nАннотация документа:\n{result['annotation']}\n")
                f.write(f"\n--- Дополнительная информация об аннотации ---\n")
                f.write(f"Длина аннотации: {result['annotation_length_chars']} символов\n")
                f.write(f"Количество предложений в аннотации: {result['annotation_sentences']}\n")
                f.write(f"\nОбщий вес (хэш + аннотация): {result['total_weight_bytes']} байт\n")

            flash(f'Файл успешно обработан! Результаты сохранены в папке output.')
            return render_template('result.html', result=result, output_json=output_filename, output_txt=txt_output_filename)

        except Exception as e:
            flash(f'Ошибка при обработке файла: {str(e)}')
            return redirect(url_for('index'))
        finally:
            # Удаляем загруженный файл после обработки
            if os.path.exists(filepath):
                os.remove(filepath)
    else:
        flash('Недопустимый тип файла. Разрешены: docx, txt, pdf, rtf')
        return redirect(url_for('index'))


@app.route('/output/<filename>')
def download_output(filename):
    return send_from_directory(app.config['OUTPUT_FOLDER'], filename, as_attachment=True)


@app.route('/output')
def list_output():
    files = []
    for filename in os.listdir(app.config['OUTPUT_FOLDER']):
        filepath = os.path.join(app.config['OUTPUT_FOLDER'], filename)
        if os.path.isfile(filepath):
            files.append({
                'name': filename,
                'size': os.path.getsize(filepath),
                'url': url_for('download_output', filename=filename)
            })
    return render_template('output_list.html', files=files)


if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=5000, threaded=True)
