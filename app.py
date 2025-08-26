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
    <a href="/download">📥 Скачать Excel</a>
  {% endif %}
</body>
</html>
"""

summary_df = None  # глобальная таблица


def parse_upd(file) -> pd.DataFrame:
    """Читает УПД и возвращает товары: код, наименование, количество, стоимость"""
    rows = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if not table:
                continue
            for row in table:
                # ищем строки с кодом из 10 цифр
                for cell in row:
                    if cell and cell.strip().isdigit() and len(cell.strip()) == 10:
                        try:
                            code = cell.strip()
                            name = row[1]
                            qty = float(row[3].replace(",", "."))
                            cost = float(row[5].replace(",", "."))
                            rows.append([code, name, qty, cost])
                        except Exception:
                            continue
    return pd.DataFrame(rows, columns=["Код вида товара", "Наименование", "Кол-во", "Стоимость (₽)"])


def parse_torg(file) -> pd.DataFrame:
    """Читает ТОРГ-12 и возвращает веса товаров"""
    rows = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if not table:
                continue
            for row in table:
                try:
                    name = row[1]
                    weight = float(row[4].replace(",", "."))
                    rows.append([name, weight])
                except Exception:
                    continue
    return pd.DataFrame(rows, columns=["Наименование", "Масса нетто (кг)"])


def build_summary(upd_df, torg_df):
    """Собирает сводную таблицу по наименованию"""
    df = pd.merge(upd_df, torg_df, on="Наименование", how="left")
    # Итоги
    total_mass = df["Масса нетто (кг)"].sum()
    total_qty = df["Кол-во"].sum()
    total_cost = df["Стоимость (₽)"].sum()
    df.loc[len(df)] = ["ИТОГО", "-", total_qty, total_cost, total_mass]
    return df


@app.route("/", methods=["GET", "POST"])
def index():
    global summary_df
    table_html = None
    if request.method == "POST":
        upd_file = request.files.get("upd")
        torg_file = request.files.get("torg")
        if upd_file and torg_file:
            upd_df = parse_upd(upd_file)
            torg_df = parse_torg(torg_file)
            summary_df = build_summary(upd_df, torg_df)
            table_html = summary_df.to_html(index=False, float_format="%.2f")
    return render_template_string(HTML_TEMPLATE, table=table_html)


@app.route("/download")
def download():
    global summary_df
    if summary_df is None:
        return "Нет данных, сначала загрузите файлы."
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="Сводная")
    output.seek(0)
    return send_file(output, download_name="summary.xlsx", as_attachment=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
