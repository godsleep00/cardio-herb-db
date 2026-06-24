import pandas as pd
from rdkit import Chem

input_file = "compounds.csv"     # 你的 CSV 文件名
output_file = "compounds_out.csv"

# 尝试三种编码读取
encodings = ["utf-8", "gbk", "latin-1"]

df = None
for enc in encodings:
    try:
        print(f"尝试编码：{enc}")
        df = pd.read_csv(input_file, encoding=enc)
        print(f"✔ 成功使用编码 {enc} 读取文件")
        break
    except Exception as e:
        print(f"✘ 使用编码 {enc} 失败：{e}")

if df is None:
    print("❌ 无法读取 CSV，请检查文件是否损坏")
    exit()

# 自动识别 inchi 列
possible_inchi_cols = ["InChI", "inchi", "INCHI", "Inchi"]
inchi_col = None

for col in df.columns:
    if col.strip() in possible_inchi_cols:
        inchi_col = col
        break

if inchi_col is None:
    print("❌ 表格中未找到 InChI 列，请检查列名")
    exit()

# 如果没有 inchikey 列则自动创建
if "InChIKey" not in df.columns:
    df["InChIKey"] = ""

print(f"识别到 InChI 列：{inchi_col}")

# 逐行转换
for idx, row in df.iterrows():
    inchi = str(row[inchi_col]).strip()
    if not inchi or inchi.startswith("nan"):
        print(f"跳过第 {idx} 行（无有效 InChI）")
        continue

    mol = Chem.MolFromInchi(inchi)
    if mol:
        inchikey = Chem.InchiToInchiKey(inchi)
        df.at[idx, "InChIKey"] = inchikey
        print(f"✔ 第 {idx} 行 转换成功：{inchikey}")
    else:
        print(f"❌ 第 {idx} 行 无法解析 InChI：{inchi}")

# 保存
df.to_csv(output_file, index=False, encoding="utf-8-sig")
print(f"\n🎉 转换完成！已输出：{output_file}")
