import pandas as pd
import re
from flask import Flask, request, send_file, render_template_string
from io import BytesIO

app = Flask(__name__)

# === HTML форма для загрузки файлов ===
UPLOAD_FORM = """
<!doctype html>
<html>
<head>
  <title>Сводная таблица</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 40px; }
    h2 { color: #333; }
    form { margin-top: 20px; }
    input[type=file] { margin: 5px 0; }
    input[type=submit] {
      background: #4CAF50; color: white; border: none; padding: 10px 15px;
      cursor: pointer; border-radius: 5px;
    }
    input[type=submit]:hover { background: #45a049; }
  </style>
</head>
<body>
  <h2>Загрузите два Excel файла</h2>
  <form method=post enctype=multipart/form-data>
    <p><input type=file name=invoice>
       <input type=file name=waybill>
       <input type=submit value="Сформировать сводную таблицу">
  </form>
</body>
</html>
"""

# === HTML для вывода результата ===
RESULT_TEMPLATE = """
<!doctype html>
<html>
<head>
  <title>Сводная таблица</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 40px; }
    h2 { color: #333; }
    .table-container { 
      max-width: 95%; 
      overflow-x: auto; 
      border: 1px solid #ccc; 
      padding: 10px; 
      margin-top: 20px;
    }
    table { border-collapse: collapse; width: 100%; }
    th, td { border: 1px solid #ccc; padding: 8px; text-align: center; }
    th { background: #f2f2f2; }
    tr:nth-child(even) { background: #f9f9f9; }
    tr:hover { background: #f1f1f1; }
    .download-btn {
      display: inline-block;
      margin-top: 20px;
      background: #2196F3; color: white;
      padding: 10px 20px; text-decoration: none;
      border-radius: 5px; font-weight: bold;
    }
    .download-btn:hover { background: #0b7dda; }
  </style>
</head>
<body>
  <h2>Сводная таблица</h2>
  <div class="table-container">
    {{ table|safe }}
  </div>
  <a href="/download" class="download-btn">📥 Скачать Excel</a>
</body>
</html>
"""

# Глобальное хранилище итогового DataFrame
SUMMARY_DF = None

def extract_tnved(text):
    """Ищет 10-значный код ТНВЭД в тексте"""
    match = re.search(r"\b\d{10}\b", str(text))
    return match.group(0) if match else None

def make_summary(invoice_file, waybill_file):
    # читаем счет: пропустить первые 8 строк и последние 11
    df_invoice = pd.read_excel(invoice_file, skiprows=8, skipfooter=11, engine="openpyxl")
    # накладная: только пропуск первых 8 строк
    df_waybill = pd.read_excel(waybill_file, skiprows=8)

    # --- обработка счета ---
    invoice_data = []
    for _, row in df_invoice.iterrows():
        line = " ".join(str(v) for v in row if pd.notna(v))
        tnved = extract_tnved(line)
        if tnved:
            parts = line.split(tnved)
            name = parts[0].strip().split(maxsplit=1)[-1]
            qty = pd.to_numeric(row.astype(str).str.replace(",", "."), errors="coerce").dropna()
            quantity = int(qty.iloc[0]) if not qty.empty else None
            cost = qty.iloc[-1] if len(qty) > 1 else None
            invoice_data.append({
                "№ п/п": len(invoice_data) + 1,
                "Код ТНВЭД": tnved,
                "Наименование товара": name,
                "Кол-во": quantity,
                "Стоимость": cost
            })

    df_invoice_clean = pd.DataFrame(invoice_data)

    # --- обработка накладной (вес) ---
    waybill_data = []
    for _, row in df_waybill.iterrows():
        line = " ".join(str(v) for v in row if pd.notna(v))
        tnved = extract_tnved(line)
        if tnved:
            qty = pd.to_numeric(row.astype(str).str.replace(",", "."), errors="coerce").dropna()
            weight = qty.iloc[-1] if not qty.empty else None
            waybill_data.append({
                "Код ТНВЭД": tnved,
                "Вес (кг)": weight
            })

    df_waybill_clean = pd.DataFrame(waybill_data)

    # --- объединяем ---
    df_summary = df_invoice_clean.merge(df_waybill_clean, on="Код ТНВЭД", how="left")

    # --- добавляем итого ---
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

@app.route("/", methods=["GET", "POST"])
def upload():
    global SUMMARY_DF
    if request.method == "POST":
        invoice = request.files.get("invoice")
        waybill = request.files.get("waybill")
        if not invoice or not waybill:
            return "Пожалуйста, загрузите оба файла"
        SUMMARY_DF = make_summary(invoice, waybill)

        # показываем таблицу в браузере
        html_table = SUMMARY_DF.to_html(index=False, border=0, justify="center")
        return render_template_string(RESULT_TEMPLATE, table=html_table)

    return render_template_string(UPLOAD_FORM)

@app.route("/download")
def download():
    global SUMMARY_DF
    if SUMMARY_DF is None:
        return "Нет данных для выгрузки. Сначала загрузите файлы."
    output = BytesIO()
    SUMMARY_DF.to_excel(output, index=False)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name="summary.xlsx")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
