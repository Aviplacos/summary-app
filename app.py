import os
import io
import pandas as pd
import pdfplumber
from flask import Flask, request, render_template_string, send_file

app = Flask(__name__)

HTML_TEMPLATE = """
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>Сводная таблица</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 40px; }
    table { border-collapse: collapse; width: 100%; }
    th, td { border: 1px solid #999; padding: 6px; text-align: left; }
    th { background: #eee; }
    pre { background: #f8f8f8; padding: 10px; border: 1px solid #ccc; }
  </style>
</head>
<body>
  <h2>Загрузка документов</h2>
  <form action="/" method="post" enctype="multipart/form-data">
    УПД (PDF): <input type="file" name="upd"><br><br>
    ТОРГ-12 (PDF): <input type="file" name="torg"><br><br>
    <input type="submit" value="Обработать">
  </form>
  {% if table %}
    <h2>Сводная таблица товаров</h2>
    {{ table|safe }}
    <br>
    <a href="/download">📥 Скачать Excel</a><br>
    <a href="/download_raw/upd">📥 Скачать сырые данные УПД (CSV)</a><br>
    <a href="/download_raw/torg">📥 Скачать сырые данные ТОРГ-12 (CSV)</a>
  {% endif %}
  {% if debug %}
    <h2>Отладка (первые строки из PDF)</h2>
    <pre>{{ debug }}</pre>
  {% endif %}
</body>
</html>
"""

summary_df = None
raw_upd = None
raw_torg = None


def auto_code(name: str) -> str:
    """Автоподстановка кода по названию (только диван/стул/пуф/банкетка)"""
    if not name:
        return "—"
    n = str(name).lower()
    if "диван" in n:
        return "9401410000"
    if "стул" in n:
        return "9401800009"
    if "пуф" in n or "банкетка" in n:
        return "9401800009"
    return "—"


def parse_upd(file):
    rows = []
    debug = []
    raw = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if not table:
                continue
            raw.extend(table)
            debug.append("=== UPD PAGE ===")
            for r in table[:10]:
                debug.append(str(r))
            for row in table:
                if not row or len(row) < 2:
                    continue
                code = None
                for cell in row:
                    if isinstance(cell, str) and cell.strip().isdigit() and len(cell.strip()) == 10:
                        code = cell.strip()
                name = row[1] if len(row) > 1 else ""
                if not code:
                    code = auto_code(name)
                qty, cost = None, None
                for cell in row:
                    if not isinstance(cell, str):
                        continue
                    val = cell.replace(" ", "").replace(",", ".")
                    try:
                        num = float(val)
                        if qty is None:
                            qty = num
                        else:
                            cost = num
                    except:
                        continue
                if name and qty and cost:
                    rows.append([code, name, qty, cost])
    return pd.DataFrame(rows, columns=["Код вида товара", "Наименование", "Кол-во", "Стоимость (₽)"]), "\n".join(debug), raw


def parse_torg(file):
    rows = []
    debug = []
    raw = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if not table:
                continue
            raw.extend(table)
            debug.append("=== TORG PAGE ===")
            for r in table[:10]:
                debug.append(str(r))
            for row in table:
                if not row or len(row) < 2:
                    continue
                name = row[1]
                weight = None
                for cell in row:
                    if not isinstance(cell, str):
                        continue
                    val = cell.replace(" ", "").replace(",", ".")
                    try:
                        weight = float(val)
                        break
                    except:
                        continue
                if name and weight:
                    rows.append([name, weight])
    return pd.DataFrame(rows, columns=["Наименование", "Масса нетто (кг)"]), "\n".join(debug), raw


def build_summary(upd_df, torg_df):
    if upd_df.empty:
        return upd_df
    df = pd.merge(upd_df, torg_df, on="Наименование", how="left")
    total_mass = df["Масса нетто (кг)"].sum(skipna=True)
    total_qty = df["Кол-во"].sum(skipna=True)
    total_cost = df["Стоимость (₽)"].sum(skipna=True)
    df.loc[len(df)] = ["ИТОГО", "-", total_qty, total_cost, total_mass]
    return df


@app.route("/", methods=["GET", "POST"])
def index():
    global summary_df, raw_upd, raw_torg
    table_html = None
    debug_output = None
    if request.method == "POST":
        upd_file = request.files.get("upd")
        torg_file = request.files.get("torg")
        if upd_file and torg_file:
            upd_df, debug_upd, raw_upd = parse_upd(upd_file)
            torg_df, debug_torg, raw_torg = parse_torg(torg_file)
            summary_df = build_summary(upd_df, torg_df)
            if not summary_df.empty:
                table_html = summary_df.to_html(index=False, float_format="%.2f")
            else:
                table_html = "<p>⚠️ Не удалось распознать данные из PDF.</p>"
            debug_output = debug_upd + "\n" + debug_torg
    return render_template_string(HTML_TEMPLATE, table=table_html, debug=debug_output)


@app.route("/download")
def download():
    global summary_df
    if summary_df is None or summary_df.empty:
        return "Нет данных, сначала загрузите файлы."
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="Сводная")
    output.seek(0)
    return send_file(output, download_name="summary.xlsx", as_attachment=True)


@app.route("/download_raw/<kind>")
def download_raw(kind):
    global raw_upd, raw_torg
    raw = raw_upd if kind == "upd" else raw_torg
    if raw is None:
        return "Нет данных."
    df = pd.DataFrame(raw)
    output = io.BytesIO()
    df.to_csv(output, index=False, encoding="utf-8-sig")
    output.seek(0)
    return send_file(output, download_name=f"raw_{kind}.csv", as_attachment=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
