import pandas as pd
import re
from flask import Flask, request, send_file, render_template_string, flash
from io import BytesIO
import logging

app = Flask(__name__)
app.secret_key = 'your-secret-key-here'
logging.basicConfig(level=logging.DEBUG)

# === HTML форма для загрузки файлов ===
UPLOAD_FORM = """
<!doctype html>
<title>Сводная таблица</title>
<h2>Загрузите два Excel файла</h2>
{% with messages = get_flashed_messages() %}
  {% if messages %}
    <div style="color: red;">
      {% for message in messages %}
        <p>{{ message }}</p>
      {% endfor %}
    </div>
  {% endif %}
{% endwith %}
<form method=post enctype=multipart/form-data>
  <p>Файл счета: <input type=file name=invoice required></p>
  <p>Файл накладной: <input type=file name=waybill required></p>
  <p><input type=submit value="Сформировать сводную таблицу"></p>
</form>
"""

def extract_tnved(text):
    """Извлекает код ТНВЭД из текста, исключая коды начинающиеся с 26"""
    try:
        matches = re.findall(r"\b\d{10}\b", str(text))
        for match in matches:
            # Исключаем коды, начинающиеся с 26
            if not match.startswith('26'):
                return match
        return None
    except Exception as e:
        logging.error(f"Ошибка при извлечении ТНВЭД: {e}")
        return None

def process_invoice(df):
    """Обрабатывает данные из счета"""
    invoice_data = []
    for _, row in df.iterrows():
        try:
            line = " ".join(str(v) for v in row if pd.notna(v))
            tnved = extract_tnved(line)
            if tnved:
                parts = line.split(tnved)
                name = parts[0].strip().split(maxsplit=1)[-1] if len(parts) > 0 else "Неизвестно"
                
                # Ищем числовые значения в строке
                numbers = re.findall(r"\d+[.,]?\d*", line)
                numeric_values = []
                for num in numbers:
                    try:
                        numeric_values.append(float(num.replace(",", ".")))
                    except ValueError:
                        continue
                
                quantity = int(numeric_values[0]) if numeric_values else None
                cost = numeric_values[-1] if len(numeric_values) > 1 else None
                
                invoice_data.append({
                    "№ п/п": len(invoice_data) + 1,
                    "Код ТНВЭД": tnved,
                    "Наименование товара": name,
                    "Кол-во": quantity,
                    "Стоимость": cost
                })
        except Exception as e:
            logging.warning(f"Ошибка обработки строки счета: {e}")
            continue
    
    return pd.DataFrame(invoice_data)

def process_waybill(df):
    """Обрабатывает данные из накладной"""
    waybill_data = []
    for _, row in df.iterrows():
        try:
            line = " ".join(str(v) for v in row if pd.notna(v))
            tnved = extract_tnved(line)
            if tnved:
                # Ищем числовые значения (предполагаем, что вес - последнее число)
                numbers = re.findall(r"\d+[.,]?\d*", line)
                numeric_values = []
                for num in numbers:
                    try:
                        numeric_values.append(float(num.replace(",", ".")))
                    except ValueError:
                        continue
                
                weight = numeric_values[-1] if numeric_values else None
                
                waybill_data.append({
                    "Код ТНВЭД": tnved,
                    "Вес (кг)": weight
                })
        except Exception as e:
            logging.warning(f"Ошибка обработки строки накладной: {e}")
            continue
    
    return pd.DataFrame(waybill_data)

def make_summary(invoice_file, waybill_file):
    """Создает сводную таблицу"""
    try:
        # Чтение файлов
        df_invoice = pd.read_excel(invoice_file)
        df_waybill = pd.read_excel(waybill_file)
        
        # Обработка данных
        df_invoice_clean = process_invoice(df_invoice)
        df_waybill_clean = process_waybill(df_waybill)
        
        if df_invoice_clean.empty:
            raise ValueError("Не удалось извлечь данные из файла счета")
        if df_waybill_clean.empty:
            raise ValueError("Не удалось извлечь данные из файла накладной")

        # Объединение данных
        df_summary = pd.merge(df_invoice_clean, df_waybill_clean, on="Код ТНВЭД", how="left")

        # Добавление итогов
        if not df_summary.empty:
            totals = {
                "№ п/п": "ИТОГО",
                "Код ТНВЭД": "",
                "Наименование товара": "",
                "Кол-во": df_summary["Кол-во"].sum(skipna=True),
                "Стоимость": df_summary["Стоимость"].sum(skipna=True),
                "Вес (кг)": df_summary["Вес (кг)"].sum(skipna=True)
            }
            df_summary = pd.concat([df_summary, pd.DataFrame([totals])], ignore_index=True)
        
        return df_summary
        
    except Exception as e:
        logging.error(f"Ошибка при создании сводной таблицы: {e}")
        raise

@app.route("/", methods=["GET", "POST"])
def upload():
    if request.method == "POST":
        try:
            invoice = request.files.get("invoice")
            waybill = request.files.get("waybill")
            
            if not invoice or not waybill:
                flash("Пожалуйста, загрузите оба файла")
                return render_template_string(UPLOAD_FORM)
            
            # Проверка расширений файлов
            if not (invoice.filename.lower().endswith(('.xls', '.xlsx')) and 
                   waybill.filename.lower().endswith(('.xls', '.xlsx'))):
                flash("Пожалуйста, загрузите файлы в формате Excel (.xls или .xlsx)")
                return render_template_string(UPLOAD_FORM)
            
            df_summary = make_summary(invoice, waybill)
            
            # Сохранение в память как Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_summary.to_excel(writer, index=False, sheet_name='Сводная таблица')
            output.seek(0)
            
            return send_file(
                output, 
                as_attachment=True, 
                download_name="summary.xlsx",
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            
        except Exception as e:
            flash(f"Произошла ошибка при обработке файлов: {str(e)}")
            logging.error(f"Ошибка обработки: {e}")
            return render_template_string(UPLOAD_FORM)
    
    return render_template_string(UPLOAD_FORM)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
