import os
import csv
from rdkit import Chem
from rdkit.Chem import Draw

# -----------------------------
# 配置
# -----------------------------
CSV_FILE = r"D:\Compounds\Compounds_results.csv"
OUTPUT_DIR = r"D:\Compounds\compound_images"
STRUCTURE_COLUMN = "smiles"     # 可选： "smiles" / "inchi" / "inchikey"
NAME_COLUMN = "name"            # 化合物名称所在列名


# InChIKey → mol 需要查询 InChI（这里不使用在线查询，只提示）
def mol_from_inchikey(inchikey):
    print(f"[警告] RDKit 不能直接从 InChIKey 重建分子: {inchikey}")
    return None


# 根据不同类型的输入转换结构
def get_mol(value, col):
    if value is None or value.strip() == "":
        return None

    try:
        if col == "smiles":
            return Chem.MolFromSmiles(value)
        elif col == "inchi":
            return Chem.MolFromInchi(value)
        elif col == "inchikey":
            return mol_from_inchikey(value)
        else:
            return None
    except:
        return None


def main():
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

    with open(CSV_FILE, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        rows = list(reader)

    print(f"读取 {len(rows)} 条记录\n")

    for i, row in enumerate(rows):
        name = row.get(NAME_COLUMN, f"compound_{i}").replace("/", "_")

        struct_value = row.get(STRUCTURE_COLUMN)
        mol = get_mol(struct_value, STRUCTURE_COLUMN)

        if mol is None:
            print(f"[跳过] {name} - 无法从 {STRUCTURE_COLUMN} 解析结构")
            continue

        # 生成 PNG
        try:
            img_path = os.path.join(OUTPUT_DIR, f"{name}.png")
            Draw.MolToFile(mol, img_path, size=(300, 300))
            print(f"[OK] 已生成结构图: {img_path}")
        except Exception as e:
            print(f"[失败] {name} 生成图片报错: {e}")


if __name__ == "__main__":
    main()
