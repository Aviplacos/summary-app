import pandas as pd
import re
from flask import Flask, request, send_file, render_template_string
from io import BytesIO

app = Flask(__name__)

# === HTML форма для загрузки файлов ===
UPLOAD_FORM = """
<!doctype html>
<title>Сводная таблица</title>
<h2>Загрузите два Excel файла</h2>
<form method=post enctype=multipart/form-data>
  <p><input type=file name=invoice>
     <input type=file name=waybill>
     <input type=submit value="Сформировать сводную таблицу">
</form>
"""

def extract_tnved(text):
    match = re.search(r"\b\d{10}\b", str(text))
    return match.group(0) if match else None

def make_summary(invoice_file, waybill_file):
    df_invoice = pd.read_excel(invoice_file)
    df_waybill = pd.read_excel(waybill_file)

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
    if request.method == "POST":
        invoice = request.files.get("invoice")
        waybill = request.files.get("waybill")
        if not invoice or not waybill:
            return "Пожалуйста, загрузите оба файла"
        df_summary = make_summary(invoice, waybill)

        # сохраняем в память как Excel
        output = BytesIO()
        df_summary.to_excel(output, index=False)
        output.seek(0)
        return send_file(output, as_attachment=True, download_name="summary.xlsx")

    return render_template_string(UPLOAD_FORM)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
