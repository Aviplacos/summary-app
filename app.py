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


# ---------- helpers ----------

def safe_float(val):
    """Безопасное преобразование строки в float."""
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
    """Извлекаем 10-значный код (только цифры). Возвращаем None если не 10 цифр."""
    if cell is None:
        return None
    digits = re.sub(r"\D", "", str(cell))
    return digits if len(digits) == 10 else None


def has_letters(s: str) -> bool:
    """Есть ли буквы (рус/лат) в строке."""
    return bool(re.search(r"[A-Za-zА-Яа-яЁё]", s or ""))


def get_cell(row, idx):
    return row[idx] if (row is not None and 0 <= idx < len(row)) else None


def cleanup_name(x: str) -> str:
    if not x:
        return ""
    s = str(x).replace("\n", " ").replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def is_total_row(name: str) -> bool:
    n = (name or "").lower()
    return any(key in n for key in ["итого", "всего", "сумма", "к оплате"])


# ---------- parsing ----------

def parse_upd(file):
    """
    УПД: базовые позиции колонок (по твоим RAW):
      name -> col 3, code -> col 4, qty -> col 7, cost -> col 9
    Если в конкретной строке каких-то колонок нет/пусты — применяем резервные эвристики.
    """
    rows = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if not table:
                continue
            for row in table:
                if not row:
                    continue

                # 1) Код: строго из "своей" колонки, иначе '—'
                raw_code_cell = get_cell(row, 4)
                code = extract_10_digit_code(raw_code_cell) or "—"

                # 2) Наименование: сначала из col 3, иначе ищем первую ячейку с буквами
                name = cleanup_name(get_cell(row, 3))
                if not has_letters(name):
                    # резервный поиск имени по строке
                    for i, cell in enumerate(row):
                        cell_str = cleanup_name(cell)
                        if has_letters(cell_str) and len(cell_str) >= 4:
                            name = cell_str
                            break

                # фильтр мусорных строк
                if not name or len(name) < 4 or is_total_row(name):
                    continue

                # 3) Количество и Стоимость:
                #    сначала пробуем фиксированные колонки, затем — скан всей строки
                qty = safe_float(get_cell(row, 7))
                cost = safe_float(get_cell(row, 9))

                # Если не нашли, сканируем числа в строке, игнорируя ячейку кода
                if qty is None or cost is None:
                    numbers = []
                    for i, cell in enumerate(row):
                        if i == 4:  # пропускаем колонку кода, чтобы не поймать 10-значный как число
                            continue
                        num = safe_float(cell)
                        if num is not None:
                            numbers.append((i, num))
                    if numbers:
                        # Грубая, но практичная эвристика:
                        # - количество — первое число в строке,
                        # - стоимость — последнее число в строке.
                        if qty is None:
                            qty = numbers[0][1]
                        if cost is None:
                            cost = numbers[-1][1]

                # Записываем только если есть qty и cost
                if qty is not None and cost is not None:
                    rows.append([code, name, qty, cost])

    return pd.DataFrame(rows, columns=["Код вида товара", "Наименование", "Кол-во", "Стоимость (₽)"])


def parse_torg(file):
    """
    ТОРГ-12: базовые позиции (по твоим RAW):
      name -> col 1, mass -> col 9
    Применяем такие же фильтры по имени и защите от разметки.
    """
    rows = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if not table:
                continue
            for row in table:
                if not row:
                    continue
                name = cleanup_name(get_cell(row, 1))
                if not name or len(name) < 4 or is_total_row(name):
                    continue
                weight = safe_float(get_cell(row, 9))
                if weight is not None:
                    rows.append([name, weight])
    return pd.DataFrame(rows, columns=["Наименование", "Масса нетто (кг)"])


def build_summary(upd_df, torg_df):
    # Левое соединение — всё из УПД остаётся
    df = pd.merge(upd_df, torg_df, on="Наименование", how="left")

    # Нумерация строк
    df.insert(0, "№", range(1, len(df) + 1))

    # ИТОГО
    total_mass = df["Масса нетто (кг)"].sum(skipna=True)
    total_qty = df["Кол-во"].sum(skipna=True)
    total_cost = df["Стоимость (₽)"].sum(skipna=True)
    df.loc[len(df)] = ["—", "ИТОГО", "-", total_qty, total_cost, total_mass]
    return df


# ---------- web ----------

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
