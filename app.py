from flask import Flask, request, send_file, render_template_string
import os
import pandas as pd
import re

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        return "Файлы приняты! 🚀"

    return """
    <h2>Проверка сервиса</h2>
    <form method="post" enctype="multipart/form-data">
      <p>ТОРГ-12 (txt): <input type="file" name="torg12"></p>
      <p>УПД (txt): <input type="file" name="upd"></p>
      <p><input type="submit" value="Загрузить"></p>
    </form>
    """

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
