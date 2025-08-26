import os
from flask import Flask, request, render_template_string, redirect, url_for
import pandas as pd

app = Flask(__name__)
UPLOAD_FOLDER = "/tmp"   # для Render лучше /tmp (пишущий каталог)
ALLOWED_EXTENSIONS = {"xlsx", "xls"}

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

TEMPLATE = """
<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <title>Загрузка и просмотр таблицы</title>
  <style>
    body { font-family: sans-serif; margin: 2em; }
    table { border-collapse: collapse; width: 100%; margin-top: 1em; }
    th, td { border: 1px solid #ddd; padding: 6px; }
    th { background: #f0f0f0; }
    .total-row td { font-weight: bold; background: #fff8e1; }
  </style>
</head>
<body>
  <h1>Загрузка Excel-файла</h1>
  <form method="post" enctype="multipart/form-data">
    <input type="file" name="file" accept=".xlsx,.xls" required>
    <button type="submit">Загрузить</button>
  </form>
  {% if table_html %}
    <h2>Содержимое файла: {{ filename }}</h2>
    {{ table_html|safe }}
  {% endif %}
</body>
</html>
"""

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route("/", methods=["GET", "POST"])
def index():
    table_html = None
    filename = None

    if request.method == "POST":
        if "file" not in request.files:
            return redirect(request.url)
        file = request.files["file"]
        if file and allowed_file(file.filename):
            filename = file.filename
            filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            file.save(filepath)

            # читаем Excel
            df = pd.read_excel(filepath)
            # подсветка "Итого"
            if "№" in df.columns:
                df_html = df.to_html(index=False, classes="dataframe")
                df_html = df_html.replace("<td>Итого</td>", "<td><b>Итого</b></td>")
                table_html = df_html
            else:
                table_html = df.to_html(index=False, classes="dataframe")

    return render_template_string(TEMPLATE, table_html=table_html, filename=filename)

@app.route("/health")
def health():
    return "OK", 200

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
