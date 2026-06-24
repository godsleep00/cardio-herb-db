import os
import csv
from rdkit import Chem
from rdkit.Chem import AllChem

# ----------------------------
# 配置你的 SDF 文件夹路径
# ----------------------------
SDF_DIR = r"D:\Compounds\sdf_files"
OUTPUT_CSV = r"D:\Compounds_aidicha_results.csv"


def process_sdf(path):
    results = []

    try:
        suppl = Chem.SDMolSupplier(path, removeHs=False)
    except Exception as e:
        print(f"[跳过] {path} - 无法读取 SDF 文件: {e}")
        return results

    if suppl is None:
        print(f"[跳过] {path} - RDKit 无法解析文件")
        return results

    for i, mol in enumerate(suppl):
        if mol is None:
            print(f"[跳过] {path} 分子 {i} - 解析失败(可能是文件损坏或结构错误)")
            continue

        try:
            # 尝试标准化
            mol = Chem.RemoveHs(mol)  # 简单去氢
        except:
            pass

        try:
            smiles = Chem.MolToSmiles(mol)
        except:
            smiles = None

        try:
            inchi = Chem.MolToInchi(mol)
        except:
            inchi = None

        try:
            inchikey = Chem.InchiToInchiKey(inchi) if inchi else None
        except:
            inchikey = None

        results.append({
            "filename": os.path.basename(path),
            "mol_index": i,
            "smiles": smiles,
            "inchi": inchi,
            "inchikey": inchikey
        })

    return results


def main():
    all_results = []

    sdf_files = [f for f in os.listdir(SDF_DIR) if f.lower().endswith(".sdf")]

    print(f"在 {SDF_DIR} 中找到 {len(sdf_files)} 个 SDF 文件\n")

    for file in sdf_files:
        path = os.path.join(SDF_DIR, file)
        print(f"Processing: {file}")

        try:
            results = process_sdf(path)
            all_results.extend(results)

        except Exception as e:
            # 最终兜底：任意未知错误直接跳过
            print(f"[跳过] 处理 {file} 时出现异常：{e}")
            continue

    # 写入 CSV
    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["filename", "mol_index", "smiles", "inchi", "inchikey"])

        for item in all_results:
            writer.writerow([
                item["filename"],
                item["mol_index"],
                item["smiles"],
                item["inchi"],
                item["inchikey"]
            ])

    print(f"\n全部完成！结果已保存到 {OUTPUT_CSV}")


if __name__ == "__main__":
    main()
