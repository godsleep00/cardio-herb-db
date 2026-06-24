"""
从Excel文件中提取CAS号，并通过PubChem转换为SMILES
"""
import argparse
import csv
import os
import time
from typing import Optional, Sequence, List, Dict, Tuple
import pandas as pd

try:
    import pubchempy as pcp
    PUBCHEMPY_AVAILABLE = True
except ImportError:
    PUBCHEMPY_AVAILABLE = False


DEFAULT_CAS_COLUMNS: Sequence[str] = (
    "CAS号",
    "CAS",
    "CAS Number",
    "cas",
    "cas_number",
    "CAS No",
    "CAS No.",
)

SUPPORTED_EXCEL_EXTENSIONS: Sequence[str] = (".xls", ".xlsx", ".xlsm", ".xlsb")
DEFAULT_EXCEL_DIRECTORY: str = (
    r"D:\共同研究者项目\活血化瘀组分库化学信息表\活血化瘀组分库化学信息表"
)
DEFAULT_SMILES_OUTPUT_DIR: str = r"D:\共同研究者项目\smile"


def _normalize_column_name(name: Optional[str]) -> str:
    if name is None:
        return ""
    name_str = str(name).strip()
    return "".join(ch.lower() for ch in name_str if not ch.isspace() and ch.isalnum())


def _locate_column(
    df: pd.DataFrame,
    preferred: Optional[str],
    candidates: Sequence[str],
    label: str,
) -> Optional[str]:
    """返回存在的列名，按 preferred -> candidates 顺序匹配。"""

    def _match_column(target: str) -> Optional[str]:
        target_norm = _normalize_column_name(target)
        if not target_norm:
            return None

        for column in df.columns:
            column_norm = _normalize_column_name(column)
            if column_norm and column_norm == target_norm:
                return column

        for column in df.columns:
            column_norm = _normalize_column_name(column)
            if column_norm and target_norm in column_norm:
                return column

        return None

    if preferred:
        match = _match_column(preferred)
        if match:
            return match
        print(f"警告: 指定的{label}列 '{preferred}' 不存在，尝试自动识别...")

    for candidate in candidates:
        match = _match_column(candidate)
        if match:
            return match

    # 兜底尝试匹配包含关键字的列
    for column in df.columns:
        column_norm = _normalize_column_name(column)
        if "cas" in column_norm:
            return column

    return None


def _ensure_unique_columns(columns: Sequence) -> List[str]:
    counts: Dict[str, int] = {}
    unique_columns: List[str] = []

    for col in columns:
        col_str = str(col).strip() if col is not None else ""
        if not col_str:
            col_str = "Unnamed"

        count = counts.get(col_str, 0)
        if count == 0:
            unique_columns.append(col_str)
        else:
            unique_columns.append(f"{col_str}_{count}")

        counts[col_str] = count + 1

    return unique_columns


def _clean_dataframe(df: Optional[pd.DataFrame]) -> Optional[pd.DataFrame]:
    if df is None:
        return None

    cleaned_df = df.dropna(how="all")
    cleaned_df = cleaned_df.dropna(axis=1, how="all")
    return cleaned_df.reset_index(drop=True)


def _find_header_row(
    df: pd.DataFrame,
    preferred: Optional[str],
    candidates: Sequence[str],
) -> Optional[int]:
    search_targets = [preferred] if preferred else []
    search_targets.extend(candidates)
    normalized_targets = [
        _normalize_column_name(target) for target in search_targets if target
    ]
    if not normalized_targets:
        normalized_targets.append("cas")

    max_rows = min(len(df), 50)
    for row_idx in range(max_rows):
        row = df.iloc[row_idx]
        for value in row:
            normalized_value = _normalize_column_name(value)
            if not normalized_value:
                continue
            for target in normalized_targets:
                if target and target in normalized_value:
                    return row_idx
            if "cas" in normalized_value:
                return row_idx
    return None


def _load_sheet_with_header_search(
    excel_file: pd.ExcelFile,
    sheet: str,
    preferred: Optional[str],
    candidates: Sequence[str],
) -> Optional[pd.DataFrame]:
    df_raw = excel_file.parse(sheet_name=sheet, header=None)
    df_raw = _clean_dataframe(df_raw)
    if df_raw is None or df_raw.empty:
        return df_raw

    header_row_idx = _find_header_row(df_raw, preferred, candidates)
    if header_row_idx is None:
        return None

    header_row = df_raw.iloc[header_row_idx].tolist()
    inferred_columns = _ensure_unique_columns(header_row)
    data = df_raw.iloc[header_row_idx + 1 :].reset_index(drop=True)
    if data.empty:
        return pd.DataFrame(columns=inferred_columns)

    data.columns = inferred_columns
    return _clean_dataframe(data)


def _read_excel_with_adaptive_header(
    excel_path: str,
    sheet_name: Optional[str],
    preferred_cas_col: Optional[str],
    candidates: Sequence[str],
) -> Tuple[pd.DataFrame, str, str]:
    excel_file = pd.ExcelFile(excel_path)

    if sheet_name and sheet_name not in excel_file.sheet_names:
        raise ValueError(
            f"指定的工作表 '{sheet_name}' 不存在。可用工作表: {excel_file.sheet_names}"
        )

    sheet_candidates = [sheet_name] if sheet_name else excel_file.sheet_names
    last_columns: Optional[List[str]] = None
    empty_sheet_encountered = False

    for sheet in sheet_candidates:
        df_default = excel_file.parse(sheet_name=sheet, header=0)
        df_default = _clean_dataframe(df_default)
        if df_default is not None and not df_default.empty:
            cas_col_found = _locate_column(
                df_default, preferred_cas_col, candidates, "CAS号"
            )
            if cas_col_found:
                return df_default, sheet, cas_col_found
            last_columns = list(df_default.columns)

        df_inferred = _load_sheet_with_header_search(
            excel_file, sheet, preferred_cas_col, candidates
        )
        if df_inferred is None:
            continue
        if df_inferred.empty:
            empty_sheet_encountered = True
            continue

        cas_col_found = _locate_column(
            df_inferred, preferred_cas_col, candidates, "CAS号"
        )
        if cas_col_found:
            return df_inferred, sheet, cas_col_found
        last_columns = list(df_inferred.columns)

    if empty_sheet_encountered and not last_columns:
        raise ValueError("Excel数据为空，无法处理。")

    if last_columns is not None:
        raise ValueError(
            "错误: Excel文件中没有找到CAS号列。\n"
            f"可用列: {last_columns}\n"
            "请使用 --cas-col 参数手动指定列名"
        )

    raise ValueError("Excel数据为空，无法处理。")


def get_smiles_from_pubchem(cas_number: str, verbose: bool = False) -> Optional[str]:
    """
    从PubChem获取SMILES字符串（通过CAS号）
    
    Args:
        cas_number: CAS号
        verbose: 是否输出详细调试信息
        
    Returns:
        SMILES字符串，如果查询失败返回None
    """
    if not PUBCHEMPY_AVAILABLE:
        return None
    
    if not cas_number or not str(cas_number).strip() or str(cas_number) == '/':
        return None
    
    cas_str = str(cas_number).strip()
    
    try:
        # pubchempy中查询CAS号需要使用'name'作为查询类型
        compounds = pcp.get_compounds(cas_str, 'name')
        if compounds and len(compounds) > 0:
            # 优先获取isomeric_smiles（保留立体化学），然后connectivity_smiles（推荐），最后canonical_smiles（已弃用但兼容）
            smiles = (
                compounds[0].isomeric_smiles 
                or getattr(compounds[0], 'connectivity_smiles', None)
                or getattr(compounds[0], 'canonical_smiles', None)
            )
            if smiles:
                if verbose:
                    print(f"  [成功] CAS号查询: {cas_str}")
                return smiles
    except Exception as e:
        if verbose:
            print(f"  [失败] CAS号查询: {cas_str} - {str(e)}")
    
    return None


def convert_cas_to_smiles(
    cas_numbers: List[str],
    delay: float = 0.1,
) -> Dict[str, Optional[str]]:
    """
    批量将CAS号转换为SMILES
    
    Args:
        cas_numbers: CAS号列表
        delay: 每次API调用的延迟（秒），避免请求过快
        
    Returns:
        CAS号到SMILES的字典映射
    """
    if not PUBCHEMPY_AVAILABLE:
        raise ImportError(
            "错误: 需要 pubchempy 库来查询SMILES\n"
            "请运行: pip install pubchempy"
        )
    
    print(f"\n开始通过PubChem将CAS号转换为SMILES...")
    print(f"共 {len(cas_numbers)} 个CAS号")
    print()
    
    cas_to_smiles = {}
    success_count = 0
    failed_count = 0
    
    for idx, cas in enumerate(cas_numbers, 1):
        print(f"[{idx:4d}/{len(cas_numbers)}] 查询 {cas}...", end=' ')
        
        smiles = get_smiles_from_pubchem(cas)
        
        if smiles:
            cas_to_smiles[cas] = smiles
            success_count += 1
            print(f"✓ {smiles[:50]}...")
        else:
            cas_to_smiles[cas] = None
            failed_count += 1
            print("✗ 查询失败（PubChem中未找到）")
        
        # 延迟以避免请求过快
        if delay > 0 and idx < len(cas_numbers):
            time.sleep(delay)
    
    print()
    print("=" * 60)
    print(f"转换完成！")
    print(f"成功获取: {success_count} 个")
    print(f"查询失败: {failed_count} 个")
    print("=" * 60)
    
    return cas_to_smiles


def _is_excel_file(path: str) -> bool:
    """判断路径是否为支持的Excel文件。"""
    if not os.path.isfile(path):
        return False
    _, ext = os.path.splitext(path)
    return ext.lower() in SUPPORTED_EXCEL_EXTENSIONS


def _gather_excel_files(directory: str) -> List[str]:
    """遍历目录并收集所有Excel文件路径（包含子目录）。"""
    excel_files: List[str] = []
    for root, _, files in os.walk(directory):
        for filename in files:
            if filename.startswith("~$"):
                # 跳过Excel的临时锁定文件
                continue
            file_path = os.path.join(root, filename)
            if _is_excel_file(file_path):
                excel_files.append(file_path)
    return sorted(excel_files)


def extract_cas_numbers(
    excel_path: str,
    cas_col: Optional[str] = None,
    sheet_name: Optional[str] = None,
    output_path: Optional[str] = None,
) -> List[str]:
    """
    从Excel文件中提取所有CAS号
    
    Args:
        excel_path: Excel文件路径
        cas_col: CAS号列名，如果为None则自动识别
        sheet_name: 工作表名称，如果为None则读取第一个
        output_path: 输出文件路径（可选），如果指定则保存到文件
        
    Returns:
        提取到的CAS号列表
    """
    df, sheet_used, cas_col_found = _read_excel_with_adaptive_header(
        excel_path=excel_path,
        sheet_name=sheet_name,
        preferred_cas_col=cas_col,
        candidates=DEFAULT_CAS_COLUMNS,
    )

    print(f"读取Excel文件: {excel_path}")
    print(f"使用工作表: {sheet_used}")
    print(f"共 {len(df)} 条记录")
    print(f"找到CAS号列: {cas_col_found}")
    print()
    
    # 提取CAS号
    cas_numbers = []
    valid_count = 0
    empty_count = 0
    
    for idx, row in df.iterrows():
        cas_value = row.get(cas_col_found)
        
        # 跳过空值
        if pd.isna(cas_value) or (isinstance(cas_value, str) and not cas_value.strip()):
            empty_count += 1
            continue
        
        # 转换为字符串并去除空白
        cas_str = str(cas_value).strip()
        
        # 跳过无效值（如 '/' 等占位符）
        if cas_str in ['/', '-', 'N/A', 'n/a', 'NA', 'na']:
            empty_count += 1
            continue
        
        # 添加到列表（去重）
        if cas_str not in cas_numbers:
            cas_numbers.append(cas_str)
            valid_count += 1
    
    print("=" * 60)
    print(f"提取完成！")
    print(f"有效CAS号: {valid_count} 个（去重后）")
    print(f"空值/无效值: {empty_count} 个")
    print(f"总计记录: {len(df)} 条")
    print("=" * 60)
    
    # 如果指定了输出路径，保存到文件
    if output_path:
        os.makedirs(os.path.dirname(output_path) if os.path.dirname(output_path) else '.', exist_ok=True)
        with open(output_path, 'w', encoding='utf-8') as f:
            for cas in cas_numbers:
                f.write(f"{cas}\n")
        print(f"\nCAS号已保存到: {os.path.abspath(output_path)}")
    
    return cas_numbers


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="从Excel文件中提取CAS号，并通过PubChem转换为SMILES"
    )
    parser.add_argument(
        "excel",
        nargs="?",
        default=DEFAULT_EXCEL_DIRECTORY,
        help=(
            "Excel文件路径或目录路径。"
            f"（默认: {DEFAULT_EXCEL_DIRECTORY}，将批量处理该目录下所有Excel）"
        ),
    )
    parser.add_argument(
        "--cas-col",
        help="CAS号列名，默认自动识别",
    )
    parser.add_argument(
        "--sheet",
        help="需要读取的工作表名称，不填则取第一张",
    )
    parser.add_argument(
        "-o",
        "--output",
        help=(
            "输出路径。处理单个文件时可指定CSV文件路径；"
            "批量模式下可指定结果目录，不填则保存到默认smile文件夹"
        ),
    )
    parser.add_argument(
        "--to-smiles",
        action="store_true",
        help="将CAS号转换为SMILES（通过PubChem）",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=0.1,
        help="每次API调用的延迟（秒），默认0.1秒，避免请求过快",
    )
    return parser.parse_args()


def _ensure_output_directory(path: str) -> str:
    """确保输出目录存在，返回绝对路径。"""
    os.makedirs(path, exist_ok=True)
    return os.path.abspath(path)


def _check_existing_output(output_path: str) -> bool:
    """
    检查输出文件是否已存在且包含有效数据
    
    Args:
        output_path: 输出文件路径
        
    Returns:
        如果文件存在且包含有效数据返回True，否则返回False
    """
    if not os.path.exists(output_path):
        return False
    
    try:
        # 读取CSV文件检查是否有数据
        df = pd.read_csv(output_path, encoding='utf-8')
        if df.empty:
            return False
        
        # 检查是否有SMILES列
        if 'SMILES' not in df.columns:
            return False
        
        # 检查是否有至少一条有效的SMILES数据（非空）
        valid_smiles = df['SMILES'].dropna()
        valid_smiles = valid_smiles[valid_smiles.str.strip() != '']
        return len(valid_smiles) > 0
    except Exception:
        # 如果读取失败，认为文件无效
        return False


def process_single_excel(
    excel_path: str,
    cas_col: Optional[str],
    sheet_name: Optional[str],
    to_smiles: bool,
    delay: float,
    output_path: Optional[str] = None,
    default_output_dir: Optional[str] = None,
) -> str:
    """
    处理单个Excel文件，可选择转换为SMILES。
    
    Returns:
        "skipped": 文件已存在，跳过转换
        "success": 成功处理
        "no_smiles": 未转换为SMILES（仅提取CAS号）
    """
    print("\n" + "#" * 80)
    print(f"处理文件: {excel_path}")
    print("#" * 80 + "\n")

    # 如果需要转换为SMILES，先检查输出文件是否已存在
    if to_smiles:
        final_output_path: str
        if output_path:
            output_path = os.path.abspath(output_path)
            if output_path.lower().endswith(".csv") or os.path.splitext(output_path)[1]:
                final_output_path = output_path
            else:
                output_dir = _ensure_output_directory(output_path)
                excel_basename = os.path.splitext(os.path.basename(excel_path))[0]
                final_output_path = os.path.join(output_dir, f"{excel_basename}_cas_smiles.csv")
        else:
            base_dir = default_output_dir or DEFAULT_SMILES_OUTPUT_DIR
            output_dir = _ensure_output_directory(base_dir)
            excel_basename = os.path.splitext(os.path.basename(excel_path))[0]
            final_output_path = os.path.join(output_dir, f"{excel_basename}_cas_smiles.csv")
        
        # 检查输出文件是否已存在
        if _check_existing_output(final_output_path):
            print(f"✓ 输出文件已存在: {final_output_path}")
            try:
                df = pd.read_csv(final_output_path, encoding='utf-8')
                valid_count = df['SMILES'].dropna()
                valid_count = valid_count[valid_count.str.strip() != '']
                print(f"  已包含 {len(valid_count)} 条有效的SMILES记录")
                print(f"  跳过转换，直接使用已有结果\n")
                return "skipped"
            except Exception as e:
                print(f"  读取已有文件时出错: {e}")
                print(f"  将重新转换...\n")

    cas_output_path = output_path if (output_path and not to_smiles) else None
    cas_numbers = extract_cas_numbers(
        excel_path=excel_path,
        cas_col=cas_col,
        sheet_name=sheet_name,
        output_path=cas_output_path,
    )

    if not to_smiles:
        if not cas_output_path and cas_numbers:
            print("\n前10个CAS号预览:")
            for i, cas in enumerate(cas_numbers[:10], 1):
                print(f"  {i}. {cas}")
            if len(cas_numbers) > 10:
                print(f"  ... 还有 {len(cas_numbers) - 10} 个CAS号")
        return "no_smiles"

    cas_to_smiles = convert_cas_to_smiles(cas_numbers, delay=delay)

    # 重新确定输出路径（与上面的逻辑一致）
    final_output_path: str
    if output_path:
        output_path = os.path.abspath(output_path)
        if output_path.lower().endswith(".csv") or os.path.splitext(output_path)[1]:
            final_output_path = output_path
        else:
            output_dir = _ensure_output_directory(output_path)
            excel_basename = os.path.splitext(os.path.basename(excel_path))[0]
            final_output_path = os.path.join(output_dir, f"{excel_basename}_cas_smiles.csv")
    else:
        base_dir = default_output_dir or DEFAULT_SMILES_OUTPUT_DIR
        output_dir = _ensure_output_directory(base_dir)
        excel_basename = os.path.splitext(os.path.basename(excel_path))[0]
        final_output_path = os.path.join(output_dir, f"{excel_basename}_cas_smiles.csv")

    os.makedirs(os.path.dirname(final_output_path), exist_ok=True)
    with open(final_output_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["CAS号", "SMILES"])
        for cas, smiles in cas_to_smiles.items():
            writer.writerow([cas, smiles if smiles else ""])

    print(f"\n结果已保存到: {os.path.abspath(final_output_path)}")
    return "success"


def process_directory(
    directory: str,
    cas_col: Optional[str],
    sheet_name: Optional[str],
    to_smiles: bool,
    delay: float,
    output_dir: Optional[str],
) -> None:
    """批量处理目录下的所有Excel文件。"""
    excel_files = _gather_excel_files(directory)
    if not excel_files:
        print(f"在目录中未找到Excel文件：{directory}")
        return

    print(f"批量处理目录: {directory}")
    print(f"共找到 {len(excel_files)} 个Excel文件\n")

    resolved_output_dir = (
        _ensure_output_directory(os.path.abspath(output_dir))
        if output_dir
        else _ensure_output_directory(DEFAULT_SMILES_OUTPUT_DIR)
    )

    # 统计信息
    total_count = len(excel_files)
    skipped_count = 0
    success_count = 0
    failed_count = 0
    skipped_files = []
    failed_files = []

    for idx, excel_file in enumerate(excel_files, 1):
        print(f"\n[{idx}/{total_count}] 处理文件: {os.path.basename(excel_file)}")
        
        try:
            result = process_single_excel(
                excel_path=excel_file,
                cas_col=cas_col,
                sheet_name=sheet_name,
                to_smiles=to_smiles,
                delay=delay,
                output_path=None,
                default_output_dir=resolved_output_dir,
            )
            if result == "skipped":
                skipped_count += 1
                skipped_files.append(os.path.basename(excel_file))
            elif result == "success":
                success_count += 1
            elif result == "no_smiles":
                success_count += 1  # 也算成功处理
        except Exception as exc:
            failed_count += 1
            failed_files.append((os.path.basename(excel_file), str(exc)))
            print(f"❌ 处理文件失败: {excel_file}")
            print(f"   错误信息: {exc}\n")
    
    # 打印最终统计
    print("\n" + "=" * 80)
    print("批量处理完成！统计信息：")
    print("=" * 80)
    print(f"总文件数:     {total_count}")
    print(f"成功转换:     {success_count}")
    if skipped_count > 0:
        print(f"跳过（已存在）: {skipped_count}")
    if failed_count > 0:
        print(f"处理失败:     {failed_count}")
    print("=" * 80)
    
    if skipped_count > 0:
        print(f"\n跳过的文件（已存在转换结果，共 {skipped_count} 个）：")
        for f in skipped_files:
            print(f"  ✓ {f}")
    
    if failed_count > 0:
        print(f"\n处理失败的文件（共 {failed_count} 个）：")
        for f, error in failed_files:
            print(f"  ❌ {f}")
            print(f"     错误: {error}")


def main() -> None:
    args = parse_args()
    input_path = os.path.abspath(args.excel)

    if not os.path.exists(input_path):
        raise FileNotFoundError(f"输入路径不存在：{input_path}")

    if os.path.isdir(input_path):
        process_directory(
            directory=input_path,
            cas_col=args.cas_col,
            sheet_name=args.sheet,
            to_smiles=args.to_smiles,
            delay=args.delay,
            output_dir=args.output,
        )
        return

    output_path = os.path.abspath(args.output) if args.output else None
    process_single_excel(
        excel_path=input_path,
        cas_col=args.cas_col,
        sheet_name=args.sheet,
        to_smiles=args.to_smiles,
        delay=args.delay,
        output_path=output_path,
        default_output_dir=None,
    )


if __name__ == "__main__":
    main()
# python cas_to_smiles.py --to-smiles

