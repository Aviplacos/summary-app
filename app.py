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

# ---------------- helpers ----------------

def safe_float(val):
    """Безопасно приводим значение к float (учитываем пробелы, запятые)."""
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
    """Извлекаем из ячейки 10-значный код (оставляем только цифры). Если не 10 цифр — None."""
    if cell is None:
        return None
    digits = re.sub(r"\D", "", str(cell))
    return digits if len(digits) == 10 else None


def cleanup_text(x: str) -> str:
    if not x:
        return ""
    s = str(x).replace("\n", " ").replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def has_letters(s: str) -> bool:
    return bool(re.search(r"[A-Za-zА-Яа-яЁё]", s or ""))


def word_count(s: str) -> int:
    if not s:
        return 0
    # Считаем слова как последовательности букв/цифр длиной >=1
    tokens = re.findall(r"[A-Za-zА-Яа-яЁё0-9]+", s)
    return len(tokens)


def is_total_row(name: str) -> bool:
    n = (name or "").lower()
    return any(k in n for k in ["итого", "всего", "сумма", "к оплате"])


def get_cell(row, idx):
    return row[idx] if (row is not None and isinstance(row, (list, tuple)) and 0 <= idx < len(row)) else None


def normalize_key(name: str) -> str:
    """Ключ для склейки: нижний регистр + схлопнутые пробелы."""
    s = cleanup_text(name).lower()
    s = re.sub(r"\s+", " ", s)
    return s

# ---------------- parsers ----------------

def iter_rows(page):
    """Аккуратно обходим все таблицы на странице."""
    # Сначала пытаемся извлечь несколько таблиц
    tables = page.extract_tables() or []
    if not tables:
        # на всякий случай – одиночная таблица
        one = page.extract_table()
        if one:
            tables = [one]
    for t in tables:
        if not t:
            continue
        for row in t:
            yield row


def parse_upd(file):
    """
    УПД по RAW:
      name -> col 3, code -> col 4, qty -> col 7, cost -> col 9.
    Правила:
      - Код только из col 4, строго 10 цифр, иначе '—' (строку не отбрасываем).
      - Наименование должно содержать > 2 слов.
    """
    rows = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            for row in iter_rows(page):
                if not row:
                    continue

                raw_name = cleanup_text(get_cell(row, 3))
                raw_code = get_cell(row, 4)

                # Код
                code = extract_10_digit_code(raw_code) or "—"

                # Наименование: если в col 3 слов <= 2, ищем альтернативу в строке
                name = raw_name
                if word_count(name) <= 2 or not has_letters(name):
                    for cell in row:
                        cand = cleanup_text(cell)
                        if has_letters(cand) and word_count(cand) > 2:
                            name = cand
                            break

                # Фильтры по наименованию
                if not name or word_count(name) <= 2 or is_total_row(name):
                    continue

                qty = safe_float(get_cell(row, 7))
                cost = safe_float(get_cell(row, 9))
                if qty is None or cost is None:
                    continue

                rows.append([code, name, qty, cost])

    df = pd.DataFrame(rows, columns=["Код вида товара", "Наименование", "Кол-во", "Стоимость (₽)"])
    if not df.empty:
        df["__key__"] = df["Наименование"].apply(normalize_key)
    return df


def parse_torg(file):
    """
    ТОРГ-12 по RAW:
      name -> col 1, mass -> col 9.
    Наименование также должно содержать > 2 слов.
    """
    rows = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            for row in iter_rows(page):
                if not row:
                    continue
                name = cleanup_text(get_cell(row, 1))
                if not name or word_count(name) <= 2 or is_total_row(name):
                    continue
                mass = safe_float(get_cell(row, 9))
                if mass is None:
                    continue
                rows.append([name, mass])

    df = pd.DataFrame(rows, columns=["Наименование", "Масса нетто (кг)"])
    if not df.empty:
        df["__key__"] = df["Наименование"].apply(normalize_key)
    return df

# ---------------- summary ----------------

def build_summary(upd_df, torg_df):
    if upd_df.empty:
        return upd_df

    # Слейка по нормализованному ключу (оставляем имена из УПД)
    df = pd.merge(
        upd_df,
        torg_df[["__key__", "Масса нетто (кг)"]],
        on="__key__",
        how="left",
        suffixes=("", "_t")
    )
    df.drop(columns=["__key__"], inplace=True)

    # Нумерация
    df.insert(0, "№", range(1, len(df) + 1))

    # ИТОГО
    total_qty = df["Кол-во"].sum(skipna=True)
    total_cost = df["Стоимость (₽)"].sum(skipna=True)
    total_mass = df["Масса нетто (кг)"].sum(skipna=True)

    # Порядок колонок уже: №, Код, Наименование, Кол-во, Стоимость, Масса
    df.loc[len(df)] = ["—", "—", "ИТОГО", total_qty, total_cost, total_mass]
    return df

# ---------------- web ----------------

@app.route("/", methods=["GET", "POST"])
def index():
    global summary_df
    table_html = None
    if request.method == "POST":
        upd_file = request.files.get("upd")
        torg_file = request.files.get("torg")

        if not upd_file or not torg_file:
            return render_template_string(HTML_TEMPLATE, table="<p>Загрузите оба PDF.</p>")

        upd_df = parse_upd(upd_file)
        torg_df = parse_torg(torg_file)

        if upd_df.empty:
            table_html = "<p>Не удалось извлечь строки УПД по заданным правилам.</p>"
        else:
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
