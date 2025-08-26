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
    .note { color: #666; font-size: 12px; }
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
    <p class="note">Основа — ТОРГ-12: все поля, кроме «Код вида товара», взяты из накладной. Код подтянут из УПД.</p>
    <br>
    <a href="/download">📥 Скачать Excel</a>
  {% endif %}
</body>
</html>
"""

summary_df = None

# ---------- Helpers ----------

def safe_float(val):
    """Безопасно приводим значение к float (учитываем пробелы/неразрывные, запятые)."""
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


def cleanup_text(x: str) -> str:
    if not x:
        return ""
    s = str(x).replace("\n", " ").replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def has_letters(s: str) -> bool:
    return bool(re.search(r"[A-Za-zА-Яа-яЁё]", s or ""))


def is_header_or_total(name: str) -> bool:
    n = (name or "").lower()
    return any(k in n for k in [
        "наименование", "товар", "итого", "всего", "сумма", "к оплате", "по счету"
    ])


def extract_10_digit_code(cell):
    """Извлекаем из ячейки ровно 10 цифр, иначе None."""
    if cell is None:
        return None
    digits = re.sub(r"\D", "", str(cell))
    return digits if len(digits) == 10 else None


def normalize_key(name: str) -> str:
    """Ключ для склейки: нижний регистр + схлопнутые пробелы + без лишней пунктуации на краях."""
    s = cleanup_text(name).lower()
    s = re.sub(r"\s+", " ", s).strip(" ,.;:-")
    return s


def get_cell(row, idx):
    return row[idx] if (row is not None and isinstance(row, (list, tuple)) and 0 <= idx < len(row)) else None


def iter_rows(page):
    """Обходим все таблицы на странице устойчиво к разрывам."""
    tables = page.extract_tables() or []
    if not tables:
        one = page.extract_table()
        if one:
            tables = [one]
    for t in tables:
        if not t:
            continue
        for row in t:
            yield row

# ---------- Parsing ----------

def parse_upd(file):
    """
    УПД: нам нужны только пары (Наименование -> Код вида товара).
    По RAW ранее: name -> col 3, code -> col 4.
    Код валиден только если 10 цифр.
    """
    pairs = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            for row in iter_rows(page):
                if not row:
                    continue
                name = cleanup_text(get_cell(row, 3))
                if not name or not has_letters(name) or is_header_or_total(name):
                    continue
                code = extract_10_digit_code(get_cell(row, 4))
                if code:
                    pairs.append((normalize_key(name), code))

    if not pairs:
        return pd.DataFrame(columns=["__key__", "Код вида товара"])

    # Если на одно имя встречается несколько кодов — берём последний (как «самый нижний» в документе).
    df = pd.DataFrame(pairs, columns=["__key__", "Код вида товара"])
    df = df.drop_duplicates(subset="__key__", keep="last")
    return df


def parse_torg(file):
    """
    ТОРГ-12: это «источник истины» по позициям.
    Извлекаем для каждой строки:
      - Наименование (обычно col 1; если пусто — берём первую текстовую ячейку с буквами)
      - Кол-во (пытаемся col 6; если нет — первое число после имени)
      - Стоимость (пытаемся взять «самое большое» число в строке, исключая массу и код)
      - Масса нетто (пытаемся col 9; если нет — None)
    ВАЖНО: строки не фильтруем по длине названия; удаляем только явные заголовки/итоги.
    Порядок строк сохраняем как в ТОРГ-12.
    """
    rows = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            for row in iter_rows(page):
                if not row:
                    continue

                # Наименование
                name = cleanup_text(get_cell(row, 1))
                if not has_letters(name):
                    # ищем первую ячейку с буквами
                    for cell in row:
                        cand = cleanup_text(cell)
                        if has_letters(cand):
                            name = cand
                            break

                if not name or is_header_or_total(name):
                    continue  # явные заголовки/итоги

                # Масса (по опыту — col 9)
                weight = safe_float(get_cell(row, 9))

                # Собираем все числа для эвристик количества/стоимости
                numbers = []
                for i, cell in enumerate(row):
                    # исключим потенциальную ячейку кода вида (index 4) — там 10 цифр склеенных,
                    # и она не должна попадать как число
                    if i == 4:
                        continue
                    num = safe_float(cell)
                    if num is not None:
                        numbers.append((i, num))

                # Кол-во: сначала пробуем col 6, иначе первое число ПОСЛЕ имени
                qty = safe_float(get_cell(row, 6))
                if qty is None:
                    # определим индекс ячейки имени
                    name_idx = None
                    for i, cell in enumerate(row):
                        if cleanup_text(cell) == name:
                            name_idx = i
                            break
                    # берём первое число после имени
                    for i, num in numbers:
                        if name_idx is None or i > name_idx:
                            qty = num
                            break

                # Стоимость: если есть числа — берём максимальное (обычно «Сумма» > «Цена»)
                cost = None
                if numbers:
                    # если вес считался (col 9), исключим его из кандидатов на стоимость
                    candidates = [(i, n) for i, n in numbers if not (i == 9 and weight is not None and n == weight)]
                    if candidates:
                        cost = max(candidates, key=lambda x: x[1])[1]

                rows.append([name, qty, cost, weight])

    # Даже если qty/cost/weight не распознаны — строку оставляем (с None)
    df = pd.DataFrame(rows, columns=["Наименование", "Кол-во", "Стоимость (₽)", "Масса нетто (кг)"])
    # Нормализованный ключ для склейки с УПД
    df["__key__"] = df["Наименование"].apply(normalize_key)
    return df


def build_summary(torg_df, upd_codes_df):
    """
    Финальная таблица: основана ТОЛЬКО на строках ТОРГ-12.
    Подставляем коды из УПД по нормализованному имени.
    Кол-во строк в результате == кол-ву строк ТОРГ-12 (после удаления заголовков/итогов).
    """
    df = pd.merge(
        torg_df,
        upd_codes_df[["__key__", "Код вида товара"]] if not upd_codes_df.empty else pd.DataFrame(columns=["__key__", "Код вида товара"]),
        on="__key__",
        how="left"
    )

    # Если кода нет — ставим "—"
    df["Код вида товара"] = df["Код вида товара"].fillna("—")

    # Переупорядочим колонки и добавим нумерацию
    df = df[["Код вида товара", "Наименование", "Кол-во", "Стоимость (₽)", "Масса нетто (кг)"]]
    df.insert(0, "№", range(1, len(df) + 1))

    # Итоги
    total_qty = df["Кол-во"].sum(skipna=True)
    total_cost = df["Стоимость (₽)"].sum(skipna=True)
    total_mass = df["Масса нетто (кг)"].sum(skipna=True)
    df.loc[len(df)] = ["—", "ИТОГО", "-", total_qty, total_cost, total_mass]

    return df

# ---------- Web ----------

@app.route("/", methods=["GET", "POST"])
def index():
    global summary_df
    table_html = None

    if request.method == "POST":
        upd_file = request.files.get("upd")
        torg_file = request.files.get("torg")

        if not upd_file or not torg_file:
            return render_template_string(HTML_TEMPLATE, table="<p>Загрузите оба PDF файла (УПД и ТОРГ-12).</p>")

        upd_codes_df = parse_upd(upd_file)
        torg_df = parse_torg(torg_file)

        if torg_df.empty:
            table_html = "<p>Не удалось извлечь позиции из ТОРГ-12.</p>"
        else:
            summary_df = build_summary(torg_df, upd_codes_df)
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
