from flask import Flask, render_template, request, jsonify
import pandas as pd
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Создаем папку для загрузок если не существует
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def process_files(proforma_file, invoice_file):
    """Обработка файлов счет-проформы и накладной"""
    try:
        # Чтение счет-проформы
        proforma_df = pd.read_excel(proforma_file, header=None)
        
        # Поиск данных товаров в счет-проформе
        goods_data = []
        for index, row in proforma_df.iterrows():
            if isinstance(row[2], str) and any(keyword in row[2] for keyword in [
                'Диван', 'Кресло', 'Комод', 'Сервант', 'Тумбочка', 'Туалетный', 
                'Зеркало', 'Кровать', 'Шкаф', 'ПУФ', 'Стол', 'Стул'
            ]):
                code = proforma_df.iloc[index, 8] if len(proforma_df.columns) > 8 else ''
                name = row[2]
                quantity = proforma_df.iloc[index, 18] if len(proforma_df.columns) > 18 else 0
                price = proforma_df.iloc[index, 22] if len(proforma_df.columns) > 22 else 0
                cost = proforma_df.iloc[index, 27] if len(proforma_df.columns) > 27 else 0
                
                if code and str(code).startswith(('70', '94')) and len(str(code)) == 10:
                    goods_data.append({
                        'code': str(code),
                        'name': name,
                        'quantity': quantity,
                        'cost': cost
                    })
        
        # Чтение накладной
        invoice_df = pd.read_excel(invoice_file, header=None)
        
        # Поиск данных о массе в накладной
        mass_data = {}
        for index, row in invoice_df.iterrows():
            if isinstance(row[0], str) and any(keyword in row[0] for keyword in [
                'Диван', 'Кресло', 'Комод', 'Сервант', 'Тумбочка', 'Туалетный', 
                'Зеркало', 'Кровать', 'Шкаф', 'ПУФ', 'Стол', 'Стул'
            ]):
                name = row[0]
                mass = invoice_df.iloc[index, 8] if len(invoice_df.columns) > 8 else 0
                mass_data[name] = mass
        
        # Сборка итоговой таблицы
        result = []
        for i, item in enumerate(goods_data, 1):
            mass = mass_data.get(item['name'], 0)
            result.append({
                '№': i,
                'Код ТНВЭД': item['code'],
                'Наименование': item['name'],
                'Масса': round(float(mass), 2) if mass else 0,
                'Количество': int(item['quantity']),
                'Стоимость': int(item['cost'])
            })
        
        return result
        
    except Exception as e:
        raise Exception(f"Ошибка обработки файлов: {str(e)}")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    if 'proforma' not in request.files or 'invoice' not in request.files:
        return jsonify({'error': 'Необходимо загрузить оба файла'}), 400
    
    proforma_file = request.files['proforma']
    invoice_file = request.files['invoice']
    
    if proforma_file.filename == '' or invoice_file.filename == '':
        return jsonify({'error': 'Файлы не выбраны'}), 400
    
    try:
        # Сохраняем файлы временно
        proforma_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(proforma_file.filename))
        invoice_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(invoice_file.filename))
        
        proforma_file.save(proforma_path)
        invoice_file.save(invoice_path)
        
        # Обрабатываем файлы
        result = process_files(proforma_path, invoice_path)
        
        # Удаляем временные файлы
        os.remove(proforma_path)
        os.remove(invoice_path)
        
        return jsonify({'data': result})
        
    except Exception as e:
        # Удаляем временные файлы в случае ошибки
        if os.path.exists(proforma_path):
            os.remove(proforma_path)
        if os.path.exists(invoice_path):
            os.remove(invoice_path)
        
        return jsonify({'error': str(e)}), 500

@app.route('/templates/<path:filename>')
def serve_static(filename):
    return render_template(filename)

if __name__ == '__main__':
    app.run(debug=True)
