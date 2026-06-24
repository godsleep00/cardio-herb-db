import os
import pandas as pd
from rdkit import Chem
from rdkit.Chem import Draw

CSV_PATH = r"D:\Compounds\sanleng\compounds.csv"
OUTPUT_DIR = r"D:\Compounds\sanleng\images"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# --- 自动读取 CSV 或 Excel ---
if CSV_PATH.lower().endswith(".csv"):
    df = pd.read_csv(CSV_PATH)
else:
    df = pd.read_excel(CSV_PATH)

print("读取成功，检测字段中...")

possible_smiles_cols = ["smiles", "SMILES", "Smiles"]
possible_inchi_cols = ["inchi", "InChI", "Inchi"]
possible_key_cols = ["inchikey", "InChIKey", "InchiKey"]

def find_column(columns, candidates):
    for c in candidates:
        if c in columns:
            return c
    return None

col_smiles = find_column(df.columns, possible_smiles_cols)
col_inchi = find_column(df.columns, possible_inchi_cols)
col_key = find_column(df.columns, possible_key_cols)

if not col_key:
    raise ValueError("❌ CSV/Excel 中没有找到 InChIKey 列")

print(f"检测到字段：SMILES={col_smiles}, InChI={col_inchi}, InChIKey={col_key}")

for _, row in df.iterrows():
    inchikey = row[col_key]
    mol = None

    if col_smiles and pd.notna(row[col_smiles]):
        mol = Chem.MolFromSmiles(row[col_smiles])

    if mol is None and col_inchi and pd.notna(row[col_inchi]):
        mol = Chem.MolFromInchi(row[col_inchi])

    if mol is None:
        print(f"❌ 跳过：无法解析结构 → {inchikey}")
        continue

    outfile = os.path.join(OUTPUT_DIR, f"{inchikey}.png")
    Draw.MolToFile(mol, outfile, size=(400, 400))
    print(f"✔ 输出：{outfile}")

print("\n🎉 完成！")
