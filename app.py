import os
import io
import re
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
    УПД (PDF): <input type="file" name="upd" accept="application/pdf"><br><br>
    ТОРГ-12 (PDF): <input type="file" name="torg" accept="application/pdf"><br><br>
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

summary_df = None


def safe_float(val):
    """Безопасное преобразование строки в float"""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str):
        s = val.replace("\u00A0", " ").replace(" ", "").replace(",", ".")
        try:
            return float(s)
        except Exception:
            return None
    return None


def extract_10_digit_code(cell):
    """Пробуем извлечь из ячейки 10-значный код (только цифры)"""
    if not cell:
        return None
    digits = re.sub(r"\D", "", str(cell))
    return digits if len(digits) == 10 else None


def parse_upd(file):
    """
    Разбор УПД:
    - Код вида товара: колонка 4 (если нет 10-значного → "—")
    - Наименование:    колонка 3
    - Кол-во:          колонка 7
    - Стоимость:       колонка 9
    """
    rows = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if not table:
                continue
            for row in table:
                if not row or len(row) < 10:
                    continue

                code = extract_10_digit_code(row[4]) or "—"
                name = str(row[3]).strip() if row[3] else ""

                # фильтр: если имя короче 4 символов — пропускаем
                if len(name) < 4:
                    continue

                qty = safe_float(row[7])
                cost = safe_float(row[9])

                if name and qty is not None and cost is not None:
                    rows.append([code, name, qty, cost])

    return pd.DataFrame(rows, columns=["Код вида товара", "Наименование", "Кол-во", "Стоимость (₽)"])


def parse_torg(file):
    """
    Разбор ТОРГ-12:
    - Наименование: колонка 1
    - Масса нетто:  колонка 9
    """
    rows = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if not table:
                continue
            for row in table:
                if not row or len(row) < 10:
                    continue
                name = str(row[1]).strip() if row[1] else ""

                # фильтр: если имя короче 4 символов — пропускаем
                if len(name) < 4:
                    continue

                weight = safe_float(row[9])
                if name and weight is not None:
                    rows.append([name, weight])
    return pd.DataFrame(rows, columns=["Наименование", "Масса нетто (кг)"])


def build_summary(upd_df, torg_df):
    df = pd.merge(upd_df, torg_df, on="Наименование", how="left")
    # Добавляем нумерацию строк
    df.insert(0, "№", range(1, len(df) + 1))
    # Итог
    total_mass = df["Масса нетто (кг)"].sum(skipna=True)
    total_qty = df["Кол-во"].sum(skipna=True)
    total_cost = df["Стоимость (₽)"].sum(skipna=True)
    df.loc[len(df)] = ["ИТОГО", "-", "-", total_qty, total_cost, total_mass]
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
    if summary_df is None or summary_df.empty:
        return "Нет данных, сначала загрузите файлы."
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="Сводная")
    output.seek(0)
    return send_file(output, download_name="summary.xlsx", as_attachment=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
