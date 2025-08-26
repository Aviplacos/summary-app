import os
from flask import Flask, render_template_string, send_from_directory, abort
import pandas as pd

APP_TITLE = "Итоговая таблица (накладная + коды из УПД)"
EXCEL_PATH = os.environ.get("EXCEL_PATH", "Итоговая_таблица.xlsx")

app = Flask(__name__)

TABLE_TEMPLATE = """
<!doctype html>
<html lang="ru">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width,initial-scale=1" />
    <title>{{ title }}</title>
    <style>
      body { font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif; margin: 24px; }
      h1 { font-size: 20px; margin: 0 0 12px; }
      .actions { margin: 0 0 16px; }
      table { border-collapse: collapse; width: 100%; }
      th, td { border: 1px solid #ddd; padding: 8px; vertical-align: top; }
      th { background: #f6f6f6; text-align: left; }
      tr:nth-child(even) td { background: #fbfbfb; }
      .nowrap { white-space: nowrap; }
      .muted { color: #666; font-size: 14px; }
      .footer { margin-top: 16px; }
      .chip { display:inline-block; padding:2px 8px; border-radius:999px; background:#eef; color:#334; font-size:12px; }
      .total-row td { font-weight: 700; background: #fff8e1; }
      @media (max-width: 900px) { table { font-size: 14px; } }
      .container { max-width: 1400px; margin: 0 auto; }
      a.btn { display:inline-block; padding:6px 10px; border-radius:6px; border:1px solid #bbb; text-decoration:none; color:#222; background:#fff; }
      a.btn:hover { background:#f3f3f3; }
    </style>
  </head>
  <body>
    <div class="container">
      <h1>{{ title }}</h1>
      <div class="actions">
        <span class="chip">строк (товаров): {{ rows_count }}</span>
        {% if excel_exists %}
        &nbsp; <a class="btn" href="/download" title="Скачать исходный Excel">Скачать Excel</a>
        {% endif %}
      </div>
      {{ table_html|safe }}
      <div class="footer muted">
        /health → OK • Стартовый файл: {{ excel_path }}
      </div>
    </div>
  </body>
</html>
"""

def load_table():
    if not os.path.exists(EXCEL_PATH):
        return None, False
    df = pd.read_excel(EXCEL_PATH)
    # Нормализуем имена колонок на случай иных регистров/пробелов
    df.columns = [str(c).strip() for c in df.columns]

    # Подсветка строки "Итого" (если есть)
    is_total = df.iloc[:, 0].astype(str).str.lower().eq("итого")
    # Чтобы точнее — ищем столбец «№»:
    if "№" in df.columns:
        is_total = df["№"].astype(str).str.lower().eq("итого")

    # Сформируем HTML-таблицу
    # Сохраняем формат значений из файла (pandas сам приведёт их к строкам в to_html)
    table_html = df.to_html(
        index=False,
        escape=False,
        border=0,
        classes="dataframe"
    )

    # Вставим CSS-класс для итоговой строки
    if is_total.any():
        # грубая замена: найдём текст ячейки "Итого" и окрасим её строку
        # (для простоты, заменим <tr><td>Итого на <tr class="total-row"><td>Итого)
        table_html = table_html.replace("<tr>\n      <td>Итого</td>",
                                        '<tr class="total-row">\n      <td>Итого</td>')

    return table_html, True

@app.route("/")
def index():
    table_html, ok = load_table()
    rows_count = "-"
    if ok:
        try:
            # пересчёт количества товарных строк без "Итого"
            df = pd.read_excel(EXCEL_PATH)
            if "№" in df.columns:
                rows_count = int((df["№"].astype(str).str.lower() != "итого").sum())
            else:
                rows_count = len(df)
        except Exception:
            pass

    if not ok:
        msg = f"""
        <p>Файл <code>{EXCEL_PATH}</code> не найден в корне приложения.</p>
        <p>Положите Excel с таблицей рядом с <code>app.py</code> или задайте путь через переменную окружения <code>EXCEL_PATH</code>.</p>
        """
        return render_template_string(TABLE_TEMPLATE, title=APP_TITLE, table_html=msg,
                                      rows_count=rows_count, excel_exists=False, excel_path=EXCEL_PATH)

    return render_template_string(TABLE_TEMPLATE, title=APP_TITLE, table_html=table_html,
                                  rows_count=rows_count, excel_exists=True, excel_path=EXCEL_PATH)

@app.route("/download")
def download():
    if not os.path.exists(EXCEL_PATH):
        abort(404)
    dirname = os.path.dirname(EXCEL_PATH) or "."
    filename = os.path.basename(EXCEL_PATH)
    return send_from_directory(dirname, filename, as_attachment=True)

@app.route("/health")
def health():
    return "OK", 200

if __name__ == "__main__":
    # Локальный запуск: python app.py
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
