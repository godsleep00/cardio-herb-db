import os
from rdkit import Chem

def sanitize_name(name: str):
    replacements = {
        'α': 'a', 'β': 'b', 'γ': 'g', 'δ': 'd',
        ' ': '_', ',': '_', '(': '', ')': '',
        '[': '', ']': '', '/': '_', '\\': '_'
    }
    for old, new in replacements.items():
        name = name.replace(old, new)
    return name

def read_mol_once(path):
    """确保文件句柄释放，不占用文件"""
    try:
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            data = f.read()
        mol = Chem.MolFromMolBlock(data, sanitize=True)
        return mol
    except Exception:
        return None

def rename_all_sdf(root_folder):
    print(f"🔍 开始扫描文件夹：{root_folder}")

    for dirpath, dirs, files in os.walk(root_folder):
        for file in files:
            if not file.lower().endswith(".sdf"):
                continue

            old_path = os.path.join(dirpath, file)

            # 先清理文件名（避免 α、β 等）
            safe_name = sanitize_name(file)
            if safe_name != file:
                safe_path = os.path.join(dirpath, safe_name)
                os.rename(old_path, safe_path)
                old_path = safe_path

            print(f"▶ 读取：{old_path}")

            # 读取 SDF（确保关闭文件）
            mol = read_mol_once(old_path)
            if mol is None:
                print(f"❌ 无法解析：{file}")
                continue

            # 获取 InChIKey
            try:
                inchikey = Chem.MolToInchiKey(mol)
            except Exception:
                print(f"❌ 无法生成 InChIKey：{file}")
                continue

            new_name = inchikey + ".sdf"
            new_path = os.path.join(dirpath, new_name)

            # 执行重命名（此时文件不再被占用）
            try:
                os.rename(old_path, new_path)
                print(f"✅ 重命名成功：{new_name}")
            except Exception as e:
                print(f"❌ 重命名失败：{file} → {e}")

    print("\n🎉 所有文件处理完成！")

if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__))
    rename_all_sdf(script_dir)
