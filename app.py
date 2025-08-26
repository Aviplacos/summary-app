import pandas as pd
import sys
import os
 
def clean_name(x: str) -> str:
    if not isinstance(x, str):
        return ""
    return " ".join(x.replace("\n", " ").replace("  ", " ").strip().split())
 
def build_final(proforma_path, waybill_path, output_path="Итоговая таблица.xlsx"):
    # Чтение файлов
    dfp = pd.read_excel(proforma_path, sheet_name="Лист_1")
    dfw = pd.read_excel(waybill_path, sheet_name="Лист_1")
 
    # --- Парсинг счета-проформы ---
    pf_items = dfp[["Unnamed: 3", "Unnamed: 22", "Unnamed: 26", "Unnamed: 10"]].copy()
    pf_items.columns = ["name", "qty", "unit_price", "code"]
    pf_items["name_clean"] = pf_items["name"].apply(clean_name)
    pf_items = pf_items[(pf_items["name_clean"] != "") & (pd.to_numeric(pf_items["qty"], errors="coerce").notna())]
    pf_items["qty"] = pd.to_numeric(pf_items["qty"], errors="coerce").astype("Int64")
    pf_items["unit_price"] = pd.to_numeric(pf_items["unit_price"], errors="coerce")
    pf_items["code"] = pf_items["code"].astype(str).str.extract(r"(\d+)", expand=False).fillna("")
    pf_items["cost"] = (pf_items["qty"].astype(float) * pf_items["unit_price"]).round(0).astype("Int64")
    pf_items = pf_items[pf_items["name_clean"].str.len() > 5].reset_index(drop=True)
 
    # --- Парсинг накладной ---
    wb_items = dfw[["Unnamed: 1", "Unnamed: 21"]].copy()
    wb_items.columns = ["name", "mass"]
    wb_items["name_clean"] = wb_items["name"].apply(clean_name)
    wb_items["mass"] = pd.to_numeric(wb_items["mass"], errors="coerce")
    wb_items = wb_items[(wb_items["name_clean"] != "") & wb_items["mass"].notna()]
    wb_items = wb_items[~wb_items["name_clean"].str.fullmatch(r"\d+")].reset_index(drop=True)
 
    # --- Объединение ---
    final = pd.merge(
        pf_items[["name_clean", "name", "qty", "cost", "code"]],
        wb_items[["name_clean", "mass"]],
        on="name_clean",
        how="left"
    )
    final = final[["name", "qty", "mass", "cost", "code"]]
    final.columns = ["Наименование", "Кол-во", "Масса", "Стоимость", "Код вида товара"]
    final["Кол-во"] = final["Кол-во"].astype("Int64")
    final["Масса"] = final["Масса"].round(2)
    final["Стоимость"] = final["Стоимость"].astype("Int64")
 
    # --- Сохранение ---
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        final.to_excel(writer, index=False, sheet_name="Итог")
 
    print(f"Итоговая таблица сохранена: {os.path.abspath(output_path)}")
 
# Если хотите запускать из командной строки:
if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Использование: python merge_tables.py Счет-проформа.xlsx Накладная.xlsx [Итог.xlsx]")
    else:
        proforma_file = sys.argv[1]
        waybill_file = sys.argv[2]
        output_file = sys.argv[3] if len(sys.argv) > 3 else "Итоговая таблица.xlsx"
        build_final(proforma_file, waybill_file, output_file)
