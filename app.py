import pandas as pd
import re
from flask import Flask, request, send_file, render_template_string
from io import BytesIO

app = Flask(__name__)

# глобальная переменная для хранения последнего результата
last_summary = None

# === HTML форма ===
UPLOAD_FORM = """
<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <title>Сводная таблица</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body class="p-4">
  <div class="container">
    <h2 class="mb-4">Загрузите два Excel файла</h2>
    <form method=post enctype=multipart/form-data class="mb-3">
      <div class="mb-3">
        <label class="form-label">Счет</label>
        <input type=file name=invoice class="form-control">
      </div>
      <div class="mb-3">
        <label class="form-label">Накладная</label>
        <input type=file name=waybill class="form-control">
      </div>
      <button type="submit" class="btn btn-primary">Сформировать сводную таблицу</button>
    </form>
  </div>
</body>
</html>
"""

def extract_tnved(text):
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
        tnved = extract_tnved(line)  # может быть None
        parts = line.split(tnved) if tnved else [line]
        name = parts[0].strip().split(maxsplit=1)[-1] if parts[0].strip() else ""

        # --- если нет кода, пробуем найти по названию аналогичного товара ---
        if not tnved and name:
            name_first_word = name.lower().split()[0]
            for item in invoice_data:
                if item["Наименование товара"] and item["Код ТНВЭД"]:
                    if item["Наименование товара"].lower().split()[0] == name_first_word:
                        tnved = item["Код ТНВЭД"]
                        break

        qty = pd.to_numeric(row.astype(str).str.replace(",", "."), errors="coerce").dropna()
        quantity = int(qty.iloc[0]) if not qty.empty else None
        cost = qty.iloc[-1] if len(qty) > 1 else None
        invoice_data.append({
            "№ п/п": len(invoice_data) + 1,
            "Код ТНВЭД": tnved if tnved else "",
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
    global last_summary
    if request.method == "POST":
        invoice = request.files.get("invoice")
        waybill = request.files.get("waybill")
        if not invoice or not waybill:
            return "Пожалуйста, загрузите оба файла"
        df_summary = make_summary(invoice, waybill)

        # сохраняем глобально
        last_summary = df_summary

        # HTML-таблица + кнопка скачать
        html_table = df_summary.to_html(index=False, classes="table table-striped table-bordered align-middle")
        return render_template_string(f"""
        <!doctype html>
        <html lang="ru">
        <head>
          <meta charset="utf-8">
          <title>Сводная таблица</title>
          <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
        </head>
        <body class="p-4">
          <div class="container">
            <h2 class="mb-4">Результат</h2>
            <div class="table-responsive">
              {html_table}
            </div>
            <a href="/download" class="btn btn-success mt-3">📥 Скачать Excel</a>
            <a href="/" class="btn btn-secondary mt-3">🔄 Загрузить новые файлы</a>
          </div>
        </body>
        </html>
        """)

    return render_template_string(UPLOAD_FORM)

@app.route("/download")
def download():
    global last_summary
    if last_summary is None:
        return "Нет данных для скачивания"
    output = BytesIO()
    last_summary.to_excel(output, index=False)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name="summary.xlsx")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
