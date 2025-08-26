from flask import Flask, request, Response
import pandas as pd

app = Flask(__name__)

def build_summary(invoice_file, upd_file):
    # Загружаем накладную (вес и количество)
    df_invoice = pd.read_excel(invoice_file)

    # Загружаем УПД (код ТНВЭД, наименование, стоимость)
    df_upd = pd.read_excel(upd_file)

    # --- Здесь нужно подогнать названия колонок под реальные ---
    # Примерные названия (ты подгони под свои):
    # invoice: ["Наименование", "Вес", "Количество"]
    # upd: ["№", "Код ТНВЭД", "Наименование", "Стоимость"]

    df = pd.merge(
        df_upd,
        df_invoice[["Наименование", "Вес", "Количество"]],
        on="Наименование",
        how="left"
    )

    # Итоги
    total_weight = df["Вес"].sum()
    total_price = df["Стоимость"].sum()
    df.loc[len(df)] = ["", "", "**ИТОГО**", total_weight, "", total_price]

    return df


def to_markdown(df):
    cols = df.columns.tolist()
    # Шапка
    md = "| " + " | ".join(cols) + " |\n"
    md += "| " + " | ".join(["---"] * len(cols)) + " |\n"
    # Данные
    for _, row in df.iterrows():
        md += "| " + " | ".join(str(row[c]) if pd.notna(row[c]) else "" for c in cols) + " |\n"
    return md


@app.route("/upload", methods=["POST"])
def upload():
    """
    Принимает два файла:
    - invoice: накладная
    - upd: УПД
    Возвращает Markdown-таблицу
    """
    invoice_file = request.files.get("invoice")
    upd_file = request.files.get("upd")

    if not invoice_file or not upd_file:
        return Response("Нужно загрузить оба файла", status=400)

    df = build_summary(invoice_file, upd_file)
    md = to_markdown(df)

    return Response(md, mimetype="text/plain; charset=utf-8")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
