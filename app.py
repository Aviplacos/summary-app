import os
import re
import pandas as pd
from flask import Flask, request, send_file, render_template_string

app = Flask(__name__)

def ru2f(s: str) -> float:
    return float(s.replace(" ","").replace(" ","").replace(",", "."))

def parse_torg12(text: str):
    row_re = re.compile(
        r'^\s*(\d+)\s+'           
        r'(.+?)\s+'                
        r'([A-ZА-Я0-9\-]+)\s+'     
        r'шт\s+796\s+\S+\s+'       
        r'([\d,]+)\s+'             
        r'([\d,]+)\s+'             
        r'([\d,]+)\s+'             
        r'([\d,]+)\s+'             
        r'(\d[\d\s]*,\d{2})\s+'    
        r'(\d[\d\s]*,\d{2})\s+0%', 
        re.MULTILINE
    )

    rows = []
    for m in row_re.finditer(text):
        rows.append({
            "№": int(m.group(1)),
            "Наименование": m.group(2).strip(),
            "Код товара": m.group(3).strip(),
            "Масса": ru2f(m.group(6)),
            "Количество": ru2f(m.group(7)),
            "Стоимость": ru2f(m.group(9)),
        })
    return pd.DataFrame(rows).sort_values("№").reset_index(drop=True)

def parse_upd(text: str):
    pattern = re.compile(
        r'([A-ZА-Я0-9\-]+)\s+\d+\s+[^\n]+?\s+([0-9]{8,12}|--)\s+796\s+шт\s+[\d,]+\s+\d[\d\s]*,\d{2}',
        re.MULTILINE
    )
    return {c.strip(): kv.strip() for c, kv in pattern.findall(text)}

def build_table(torg_text: str, upd_text: str):
    df_torg = parse_torg12(torg_text)
    upd_map = parse_upd(upd_text)
    df_torg["Код вида товара"] = df_torg["Код товара"].map(upd_map).fillna("--")
    df = df_torg[["№", "Код вида товара", "Код товара", "Наименование", "Масса", "Количество", "Стоимость"]]

    sum_mass = df["Масса"].sum()
    sum_qty = df["Количество"].sum()
    sum_cost = df["Стоимость"].sum()
    df.loc[len(df)] = ["ИТОГО", "", "", "", round(sum_mass, 2), int(sum_qty), round(sum_cost, 2)]
    return df

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        torg12_file = request.files["torg12"]
        upd_file = request.files["upd"]
        torg_text = torg12_file.read().decode("utf-8")
        upd_text = upd_file.read().decode("utf-8")

        df = build_table(torg_text, upd_text)

        # сохраняем в Excel
        out_xlsx = "output.xlsx"
        df.to_excel(out_xlsx, index=False)

        return send_file(out_xlsx, as_attachment=True)

    return render_template_string("""
    <h2>Загрузка ТОРГ-12 и УПД</h2>
    <form method="post" enctype="multipart/form-data">
      <p>ТОРГ-12 (txt): <input type="file" name="torg12" required></p>
      <p>УПД (txt): <input type="file" name="upd" required></p>
      <p><input type="submit" value="Обработать"></p>
    </form>
    """)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
