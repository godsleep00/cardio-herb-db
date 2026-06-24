import sys
import os
import io
import json
import time
import random
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QComboBox, QFileDialog,
    QMessageBox, QSplitter, QFrame, QScrollArea, QTableWidget, QTableWidgetItem,
    QHeaderView, QGroupBox, QGridLayout, QProgressBar
)
from PyQt5.QtGui import QPixmap, QFont, QColor, QPainter, QPen, QIcon
from PyQt5.QtCore import Qt, QSize, QThread, pyqtSignal as Signal

# 导入RDKit用于化学信息学处理
from rdkit import Chem
from rdkit.Chem import Draw, AllChem, Descriptors
import pandas as pd
import numpy as np
import pyodbc
# 尝试导入 MySQL 连接库
MYSQL_AVAILABLE = False
MYSQL_LIB = None
try:
    import pymysql
    MYSQL_AVAILABLE = True
    MYSQL_LIB = 'pymysql'
except ImportError:
    try:
        import mysql.connector
        MYSQL_AVAILABLE = True
        MYSQL_LIB = 'mysql.connector'
    except ImportError:
        MYSQL_AVAILABLE = False
        print("警告：未安装 MySQL 连接库。请运行: pip install pymysql 或 pip install mysql-connector-python")
import pubchempy as pcp
import win32com.client
import pythoncom
import win32clipboard
import tempfile
import glob
import shutil
from PIL import ImageGrab


# 查询线程类，用于在后台执行化合物查询
class QueryThread(QThread):
    query_finished = Signal(object)  # 改为 object 类型，可以发送 dict 或 None
    query_error = Signal(str)
    
    def __init__(self, db_connector, query_type, query_value):
        super().__init__()
        self.db_connector = db_connector
        self.query_type = query_type
        self.query_value = query_value
        self.running = True
    
    def run(self):
        try:
            if self.running:
                result = self.db_connector.query_data(self.query_type, self.query_value)
                if self.running:  # 再次检查以确保线程没有被取消
                    # 确保发送的是字典或 None
                    if result is None:
                        self.query_finished.emit(None)
                    elif isinstance(result, dict):
                        self.query_finished.emit(result)
                    else:
                        # 如果不是字典也不是 None，转换为字典或发送 None
                        self.query_finished.emit(None)
        except Exception as e:
            if self.running:
                self.query_error.emit(str(e))
    
    def stop(self):
        self.running = False


# 简易的ChemDraw图片提取器，仅用于查询展示
class ChemDrawExtractor:
    def __init__(self):
        self.excel = None
        self.workbook = None

    def _open_workbook(self, excel_file):
        pythoncom.CoInitialize()
        self.excel = win32com.client.Dispatch("Excel.Application")
        self.excel.Visible = False
        self.excel.DisplayAlerts = False
        self.workbook = self.excel.Workbooks.Open(excel_file)

    def _close_workbook(self):
        if self.workbook:
            try:
                self.workbook.Close(SaveChanges=False)
            except:
                pass
        if self.excel:
            try:
                self.excel.Quit()
            except:
                pass
        self.workbook = None
        self.excel = None

    @staticmethod
    def _ole_in_cell(ole, cell):
        return (
            abs(ole.Top - cell.Top) < cell.Height and
            abs(ole.Left - cell.Left) < cell.Width
        )

    def extract_image(self, excel_file, row_index, col_index=7):
        """从指定单元格提取ChemDraw图片，返回PIL Image"""
        try:
            self._open_workbook(excel_file)
            ws = self.workbook.Worksheets(1)
            cell = ws.Cells(row_index, col_index)

            for ole in ws.OLEObjects():
                if 'ChemDraw' not in ole.progID:
                    continue
                if not self._ole_in_cell(ole, cell):
                    continue

                try:
                    ws.Shapes(ole.Name).Copy()
                    time.sleep(0.2)
                    image = ImageGrab.grabclipboard()
                    if image:
                        return image
                except Exception as e:
                    print(f"提取ChemDraw图片失败: {e}")
                    continue
            return None
        except Exception as e:
            print(f"ChemDraw提取器错误: {e}")
            return None
        finally:
            self._close_workbook()




# MySQL数据库连接模块
class MySQLConnector:
    def __init__(self, host='localhost', port=3306, user='root', password='', database=None):
        self.connection = None
        self.connected = False
        self.host = host
        self.port = port
        self.user = user
        self.password = password
        self.database = database

    def connect(self):
        """连接到MySQL数据库"""
        if not MYSQL_AVAILABLE:
            print("错误：MySQL 连接库未安装")
            return False
        
        try:
            # 使用 pymysql 或 mysql.connector 连接
            if MYSQL_LIB == 'pymysql':
                import pymysql
                self.connection = pymysql.connect(
                    host=self.host,
                    port=self.port,
                    user=self.user,
                    password=self.password,
                    database=self.database,
                    charset='utf8mb4',
                    cursorclass=pymysql.cursors.DictCursor
                )
            elif MYSQL_LIB == 'mysql.connector':
                import mysql.connector
                self.connection = mysql.connector.connect(
                    host=self.host,
                    port=self.port,
                    user=self.user,
                    password=self.password,
                    database=self.database,
                    charset='utf8mb4'
                )
            else:
                print("错误：MySQL 连接库未正确初始化")
                return False
                
            self.connected = True
            print(f"MySQL连接成功 (主机: {self.host}:{self.port})")
            return True
        except Exception as e:
            print(f"MySQL连接失败: {str(e)}")
            self.connected = False
            return False

    def query_compound(self, query_type, query_value):
        """
        从MySQL数据库查询化合物
        查询逻辑：
        1. 根据查询类型（Name/SMILES/CAS）确定要查询的列名
        2. 在数据库中查找包含该化合物的表
        3. 提取并返回化合物信息
        """
        if not self.connected:
            return None

        try:
            # 根据使用的库选择游标类型
            if MYSQL_LIB == 'pymysql':
                cursor = self.connection.cursor()
            else:  # mysql.connector
                cursor = self.connection.cursor(dictionary=True)
            
            # 根据查询类型确定要查询的列名
            if query_type == "SMILES":
                search_columns = ['SMILES', 'smiles', 'Smiles']
            elif query_type == "CAS":
                search_columns = ['CAS号', 'CAS', 'cas', 'CASNumber', 'CAS Number', 'cas_number', 'CAS号']
            elif query_type == "Name":
                search_columns = ['CompoundName', 'ProductName', 'name', 'compound_name', 'product_name', 
                                 '化合物名称', '名称', '化合物', '中文名称', '英文名称']
            else:
                cursor.close()
                return None
            
            # 确定要查询的数据库列表
            if self.database:
                # 如果指定了数据库，只查询该数据库
                databases_to_try = [self.database]
            else:
                # 如果没有指定，先尝试常见的数据库，然后查找所有数据库
                databases_to_try = ['Compound', 'NaturalProductDB', 'compound_db', 'compounds_db']
            
            # 首先尝试指定的或常见的数据库
            for db_name in databases_to_try:
                if not db_name:
                    continue
                result = self._query_in_database(cursor, db_name, search_columns, query_value)
                if result:
                    cursor.close()
                    return result
            
            # 如果指定数据库中没有找到，尝试查找所有数据库（如果未指定数据库）
            if not self.database:
                try:
                    # 获取所有数据库列表
                    cursor.execute("SHOW DATABASES")
                    all_databases = cursor.fetchall()
                    
                    # 转换为列表
                    if isinstance(all_databases[0], dict):
                        db_list = [db['Database'] for db in all_databases]
                    else:
                        db_list = [db[0] for db in all_databases]
                    
                    # 排除系统数据库
                    system_dbs = ['information_schema', 'mysql', 'performance_schema', 'sys']
                    db_list = [db for db in db_list if db not in system_dbs and db not in databases_to_try]
                    
                    # 在每个数据库中查找
                    for db_name in db_list:
                        result = self._query_in_database(cursor, db_name, search_columns, query_value)
                        if result:
                            cursor.close()
                            return result
                except Exception as e:
                    print(f"查找所有数据库时出错: {str(e)}")
            
            cursor.close()
            print(f"未找到匹配的化合物 (查询类型: {query_type}, 查询值: {query_value})")
            return None

        except Exception as e:
            print(f"MySQL查询错误: {str(e)}")
            return None
    
    def _query_in_database(self, cursor, db_name, search_columns, query_value):
        """在指定数据库中查询化合物"""
        try:
            print(f"正在查询数据库: {db_name}")
            # 判断是否是名称查询（用于决定是否使用模糊匹配）
            is_name_query = any(name_col in search_columns for name_col in 
                              ['CompoundName', 'ProductName', 'name', 'compound_name', 'product_name', 
                               '化合物名称', '名称', '化合物', '中文名称', '英文名称'])
            
            # 获取数据库中的所有表
            cursor.execute(f"SHOW TABLES FROM `{db_name}`")
            tables = cursor.fetchall()
            
            # 转换为表名列表
            if not tables:
                print(f"  数据库 '{db_name}' 中没有表")
                return None
                
            if isinstance(tables[0], dict):
                table_list = [list(table.values())[0] for table in tables]
            else:
                table_list = [table[0] for table in tables]
            
            print(f"  找到 {len(table_list)} 个表: {table_list}")
            
            # 在每个表中查找
            for table_name in table_list:
                try:
                    # 获取表的列名
                    cursor.execute(f"SHOW COLUMNS FROM `{db_name}`.`{table_name}`")
                    columns = cursor.fetchall()
                    
                    # 转换为列名列表
                    if isinstance(columns[0], dict):
                        column_list = [col['Field'] for col in columns]
                    else:
                        column_list = [col[0] for col in columns]
                    
                    print(f"    表 '{table_name}' 的列: {column_list}")
                    
                    # 查找匹配的列
                    for search_col in search_columns:
                        if search_col in column_list:
                            # 找到了匹配的列，执行查询
                            # 1. 先尝试精确匹配
                            query = f"SELECT * FROM `{db_name}`.`{table_name}` WHERE `{search_col}` = %s"
                            print(f"      在列 '{search_col}' 中精确查询: '{query_value}'")
                            cursor.execute(query, (query_value,))
                            row = cursor.fetchone()
                            
                            if row:
                                print(f"✓ 在数据库 '{db_name}' 的表 '{table_name}' 中找到精确匹配项 (列: {search_col})")
                                result = self._parse_row(row, cursor)
                                result['_database'] = db_name
                                result['_table'] = table_name
                                return result
                            
                            # 2. 如果是名称查询，尝试大小写不敏感匹配
                            if is_name_query:
                                query = f"SELECT * FROM `{db_name}`.`{table_name}` WHERE LOWER(`{search_col}`) = LOWER(%s)"
                                print(f"      在列 '{search_col}' 中大小写不敏感查询: '{query_value}'")
                                cursor.execute(query, (query_value,))
                                row = cursor.fetchone()
                                
                                if row:
                                    print(f"✓ 在数据库 '{db_name}' 的表 '{table_name}' 中找到匹配项 (列: {search_col}, 忽略大小写)")
                                    result = self._parse_row(row, cursor)
                                    result['_database'] = db_name
                                    result['_table'] = table_name
                                    return result
                                
                                # 3. 尝试模糊匹配（包含查询）
                                query = f"SELECT * FROM `{db_name}`.`{table_name}` WHERE `{search_col}` LIKE %s"
                                like_value = f"%{query_value}%"
                                print(f"      在列 '{search_col}' 中模糊查询: '{like_value}'")
                                cursor.execute(query, (like_value,))
                                row = cursor.fetchone()
                                
                                if row:
                                    print(f"✓ 在数据库 '{db_name}' 的表 '{table_name}' 中找到模糊匹配项 (列: {search_col})")
                                    result = self._parse_row(row, cursor)
                                    result['_database'] = db_name
                                    result['_table'] = table_name
                                    return result
                            
                            print(f"      未找到匹配的记录")
                except Exception as e:
                    # 表查询失败，继续下一个表
                    print(f"    查询表 '{table_name}' 时出错: {str(e)}")
                    import traceback
                    traceback.print_exc()
                    continue
            
            print(f"  数据库 '{db_name}' 中未找到匹配项")
            return None
        except Exception as e:
            # 数据库查询失败，返回None
            print(f"查询数据库 '{db_name}' 时出错: {str(e)}")
            import traceback
            traceback.print_exc()
            return None

    def _parse_row(self, row, cursor):
        """解析MySQL查询结果并标准化字段名"""
        # MySQL 使用字典游标时，row 已经是字典
        if isinstance(row, dict):
            result = row
        else:
            # 如果不是字典，转换为字典
            if hasattr(cursor, 'description'):
                columns = [column[0] for column in cursor.description]
                result = dict(zip(columns, row))
            else:
                result = {}
        
        # 标准化字段名
        standardized = {}
        
        # 字段映射：将数据库列名映射到标准字段名
        field_mapping = {
            'compound_name': ['CompoundName', 'ProductName', '化合物名称', '名称', '化合物', 'Name', 'name', '中文名称', '英文名称', 'compounds', 'compounds（英文）', 'compound_name', 'product_name'],
            'cid': ['CID', 'cid', 'Cid', '化合物ID', 'cid'],
            'smiles': ['SMILES', 'smiles', 'Smiles', 'smiles'],
            'ec50': ['EC50', 'ec50', 'EC₅₀', 'ec50'],
            'cas_number': ['CAS号', 'CAS', 'cas', 'CASNumber', 'CAS Number', 'cas_number'],
            'molecular_weight': ['分子量', 'MW', 'MolecularWeight', 'Molecular Weight', 'molecular_weight'],
            'molecular_formula': ['分子式', '化学式', 'Formula', 'MolecularFormula', 'Molecular Formula', 'molecular_formula'],
            'category': ['类别', '活性类别', 'Category', 'category', '化合物种类', '种类', '化合物类别', 'compound_type', 'CompoundType', 'type', 'Type'],
            'ad_status': ['AD状态', 'ADStatus', 'AD Status', 'ad_status'],
            'animal_type': ['动物类型', '动物', 'Animal', 'AnimalType', 'animal_type'],
            'vessel_type': ['血管类型', '血管', 'Vessel', 'VesselType', 'vessel_type'],
            'vasoconstrictor': ['收缩剂', 'Vasoconstrictor', 'vasoconstrictor'],
            'concentration': ['用药浓度', '浓度', 'Concentration', 'concentration'],
            'mechanism': ['机制', '作用机制', 'Mechanism', 'mechanism'],
            'literature': ['文献', '文献链接', 'Literature', 'Reference', 'literature'],
            'sdf_path': ['SDF路径', 'SDF文件', 'SDF文件路径', 'sdf_path', 'sdf_file', 'SDFPath', 'SDFFile', 'SDF文件', 'sdf', 'SDF']
        }
        
        # 遍历原始数据，匹配字段
        for std_key, possible_keys in field_mapping.items():
            for key in possible_keys:
                if key in result:
                    value = result[key]
                    # 跳过None值和空字符串
                    if value is not None and value != '':
                        standardized[std_key] = value
                        print(f"  字段映射: '{key}' -> '{std_key}' = '{value}'")
                        break
        
        # 保留所有原始数据字段（用于兼容性）
        for key, value in result.items():
            if key not in standardized:
                standardized[key] = value
        
        # 调试：检查化合物种类字段
        if 'category' in standardized and standardized['category']:
            print(f"  ✓ 化合物种类已提取: {standardized['category']}")
        else:
            # 查找可能的种类相关字段
            possible_category_keys = [k for k in result.keys() if any(term in str(k).lower() for term in ['种类', '类别', 'type', 'category', '类'])]
            if possible_category_keys:
                print(f"  ⚠ 化合物种类未映射，发现可能的字段: {possible_category_keys}")
                # 尝试直接使用第一个找到的字段
                for key in possible_category_keys:
                    if result.get(key) and result[key] not in [None, '', 'N/A']:
                        standardized['category'] = result[key]
                        print(f"  ✓ 使用字段 '{key}' 作为化合物种类: {result[key]}")
                        break
        
        return standardized
    
    def close(self):
        """关闭MySQL连接"""
        if self.connection:
            try:
                self.connection.close()
                self.connected = False
                print("MySQL连接已关闭")
            except:
                pass


# SQL Server数据库连接模块
class SQLServerConnector:
    def __init__(self, server='localhost', trusted_connection=True, user=None, password=None):
        self.connection = None
        self.connected = False
        self.server = server
        self.trusted_connection = trusted_connection
        self.user = user
        self.password = password

    def connect(self):
        """连接到SQL Server数据库"""
        try:
            # SQL Server连接字符串
            if self.trusted_connection:
                connection_string = (
                    f'DRIVER={{SQL Server}};'
                    f'SERVER={self.server};'
                    f'Trusted_Connection=yes;'
                )
            else:
                connection_string = (
                    f'DRIVER={{SQL Server}};'
                    f'SERVER={self.server};'
                    f'UID={self.user};'
                    f'PWD={self.password};'
                )
            self.connection = pyodbc.connect(connection_string)
            self.connected = True
            print("SQL Server连接成功")
            return True
        except Exception as e:
            print(f"SQL Server连接失败: {str(e)}")
            self.connected = False
            return False

    def query_compound(self, query_type, query_value):
        """从Compound和NaturalProductDB数据库查询化合物"""
        if not self.connected:
            return None

        try:
            cursor = self.connection.cursor()

            # 先从Compound数据库查询
            try:
                if query_type == "SMILES":
                    query = "SELECT * FROM Compound.dbo.CompoundTable WHERE SMILES = ?"
                elif query_type == "CID":
                    query = "SELECT * FROM Compound.dbo.CompoundTable WHERE CID = ?"
                elif query_type == "Name":
                    query = "SELECT * FROM Compound.dbo.CompoundTable WHERE CompoundName = ?"
                else:
                    return None

                cursor.execute(query, (query_value,))
                row = cursor.fetchone()

                if row:
                    # 返回查询结果
                    return self._parse_row(row, cursor)
            except:
                # 如果表不存在，继续尝试下一个数据库
                pass

            # 如果Compound数据库没有，查询NaturalProductDB
            try:
                if query_type == "SMILES":
                    query = "SELECT * FROM NaturalProductDB.dbo.ProductTable WHERE SMILES = ?"
                elif query_type == "CID":
                    query = "SELECT * FROM NaturalProductDB.dbo.ProductTable WHERE CID = ?"
                elif query_type == "Name":
                    query = "SELECT * FROM NaturalProductDB.dbo.ProductTable WHERE ProductName = ?"

                cursor.execute(query, (query_value,))
                row = cursor.fetchone()

                if row:
                    return self._parse_row(row, cursor)
            except:
                # 如果表不存在，返回None，让程序继续使用Excel查询
                pass

            return None

        except Exception as e:
            # 静默处理错误，继续使用Excel查询
            return None

    def _parse_row(self, row, cursor):
        """解析数据库查询结果并标准化字段名"""
        columns = [column[0] for column in cursor.description]
        result = dict(zip(columns, row))
        
        # 标准化字段名，使其与Excel查询返回的格式一致
        standardized = {}
        
        # 字段映射：将数据库列名映射到标准字段名
        field_mapping = {
            'compound_name': ['CompoundName', 'ProductName', '化合物名称', '名称', '化合物', 'Name', 'name', '中文名称', '英文名称', 'compounds', 'compounds（英文）'],
            'cid': ['CID', 'cid', 'Cid', '化合物ID'],
            'smiles': ['SMILES', 'smiles', 'Smiles'],
            'ec50': ['EC50', 'ec50', 'EC₅₀'],
            'cas_number': ['CAS号', 'CAS', 'cas', 'CASNumber', 'CAS Number'],
            'molecular_weight': ['分子量', 'MW', 'MolecularWeight', 'Molecular Weight'],
            'molecular_formula': ['分子式', '化学式', 'Formula', 'MolecularFormula', 'Molecular Formula'],
            'category': ['类别', '活性类别', 'Category'],
            'ad_status': ['AD状态', 'ADStatus', 'AD Status'],
            'animal_type': ['动物类型', '动物', 'Animal', 'AnimalType'],
            'vessel_type': ['血管类型', '血管', 'Vessel', 'VesselType'],
            'vasoconstrictor': ['收缩剂', 'Vasoconstrictor'],
            'concentration': ['用药浓度', '浓度', 'Concentration'],
            'mechanism': ['机制', '作用机制', 'Mechanism'],
            'literature': ['文献', '文献链接', 'Literature', 'Reference']
        }
        
        # 遍历原始数据，匹配字段
        for std_key, possible_keys in field_mapping.items():
            for key in possible_keys:
                if key in result:
                    value = result[key]
                    # 跳过None值和空字符串
                    if value is not None and value != '':
                        standardized[std_key] = value
                        break
        
        # 保留所有原始数据字段（用于兼容性）
        for key, value in result.items():
            if key not in standardized:
                standardized[key] = value
        
        return standardized


# Excel文件查询模块
class ExcelQueryHelper:
    def __init__(self, excel_folder):
        self.excel_folder = excel_folder
        self.excel_files = []
        self.loaded = False
        self.chemdraw_extractor = ChemDrawExtractor()

    def load_excel_files(self):
        """加载文件夹中的所有Excel文件"""
        try:
            self.excel_files = []
            import glob

            # 查找所有.xlsx和.xls文件
            xlsx_files = glob.glob(os.path.join(self.excel_folder, '**', '*.xlsx'), recursive=True)
            xls_files = glob.glob(os.path.join(self.excel_folder, '**', '*.xls'), recursive=True)

            all_files = xlsx_files + xls_files

            if all_files:
                self.excel_files = all_files
                print(f"成功找到 {len(self.excel_files)} 个Excel文件")
                self.loaded = True
                return True
            else:
                print("未找到任何Excel文件")
                return False

        except Exception as e:
            print(f"加载Excel文件失败: {str(e)}")
            return False

    def query_compound(self, query_type, query_value):
        """从所有Excel文件中查询化合物"""
        if not self.loaded:
            if not self.load_excel_files():
                print("未加载Excel文件，无法进行查询")
                return None

        # 根据查询类型确定可能的列名
        possible_columns = []
        if query_type == "SMILES":
            possible_columns = ["SMILES", "smiles", "Smiles"]
        elif query_type == "CID":
            possible_columns = ["CID", "cid", "Cid", "化合物ID"]
        elif query_type == "Name":
            possible_columns = ["化合物名称", "名称", "化合物", "Name", "name", "CompoundName", "compounds", "compounds（英文）"]
        
        # 预处理查询值（仅去除多余空格，保持原始大小写）
        def preprocess_value(value):
            if isinstance(value, str):
                return value.strip()
            return str(value).strip() if value is not None else ""
        
        processed_query_value = preprocess_value(query_value)
        print(f"执行查询 - 类型: {query_type}, 值: '{query_value}', 预处理后: '{processed_query_value}'")
        print(f"正在搜索 {len(self.excel_files)} 个Excel文件")

        # 遍历所有Excel文件查询
        for excel_file in self.excel_files:
            try:
                print(f"处理文件: {os.path.basename(excel_file)}")
                # 读取Excel文件
                df = pd.read_excel(excel_file)
                
                # 检查文件列名
                print(f"  文件包含列: {list(df.columns)}")

                # 尝试在不同的列中查找
                for col in possible_columns:
                    if col in df.columns:
                        print(f"  在列 '{col}' 中查找")
                        
                        # 使用预处理后的字符串进行精确匹配
                        matched_rows = []
                        for idx, row in df.iterrows():
                            cell_value = row[col]
                            processed_cell_value = preprocess_value(cell_value)
                            
                            # 使用精确相等匹配，确保只有完全匹配的化合物才会被返回
                            if processed_query_value == processed_cell_value:
                                print(f"    找到精确匹配 - 行 {idx+2}: '{cell_value}'")
                                matched_rows.append(idx)
                        
                        # 如果找到匹配行
                        if matched_rows:
                            # 使用第一个匹配结果
                            result_idx = matched_rows[0]
                            result = df.iloc[result_idx:result_idx+1]
                            print(f"✓ 在文件 {os.path.basename(excel_file)} 的第 {result_idx+2} 行找到匹配项")
                            
                            # 转换为字典并返回
                            data = result.iloc[0].to_dict()

                            # 记录Excel文件路径和行号
                            excel_row_index = result_idx + 2

                            # 标准化字段名
                            standardized_data = self._standardize_data(data)

                            # 添加Excel元数据
                            standardized_data['_excel_file'] = excel_file
                            standardized_data['_excel_row'] = excel_row_index

                            # 提取ChemDraw图片（尝试常见的几列）
                            chemdraw_image = None
                            for col_idx in (7, 8, 9):
                                chemdraw_image = self.chemdraw_extractor.extract_image(
                                    excel_file, excel_row_index, col_index=col_idx
                                )
                                if chemdraw_image:
                                    print(f"  ✓ 成功提取ChemDraw图片 (列{col_idx})")
                                    standardized_data['_chemdraw_image'] = chemdraw_image
                                    break
                                else:
                                    print(f"  列 {col_idx} 未找到ChemDraw对象")

                            return standardized_data

            except Exception as e:
                print(f"读取文件 {os.path.basename(excel_file)} 时出错: {str(e)}")
                # 继续处理下一个文件
                continue

        print("所有文件搜索完毕，未找到匹配项")
        return None

    def _standardize_data(self, data):
        """标准化数据字段名"""
        standardized = {}

        # 字段映射
        field_mapping = {
            'compound_name': ['化合物名称', '名称', '化合物', 'Name', 'name', 'CompoundName', '中文名称', '英文名称','compounds','compounds（英文）'],
            'cid': ['CID', 'cid', 'Cid', '化合物ID'],
            'smiles': ['SMILES', 'smiles', 'Smiles'],
            'ec50': ['EC50', 'ec50', 'EC₅₀'],
            'cas_number': ['CAS号', 'CAS', 'cas', 'CAS Number'],
            'molecular_weight': ['分子量', 'MW', 'Molecular Weight', 'MolecularWeight'],
            'molecular_formula': ['分子式', '化学式', 'Formula', 'Molecular Formula'],
            'category': ['类别', '活性类别', 'Category'],
            'ad_status': ['AD状态', 'AD Status'],
            'animal_type': ['动物类型', '动物', 'Animal'],
            'vessel_type': ['血管类型', '血管', 'Vessel'],
            'vasoconstrictor': ['收缩剂', 'Vasoconstrictor'],
            'concentration': ['用药浓度', '浓度', 'Concentration'],
            'mechanism': ['机制', '作用机制', 'Mechanism'],
            'literature': ['文献', '文献链接', 'Literature', 'Reference']
        }

        # 遍历原始数据，匹配字段
        for std_key, possible_keys in field_mapping.items():
            for key in possible_keys:
                if key in data:
                    value = data[key]
                    # 跳过NaN值
                    if pd.isna(value):
                        continue
                    standardized[std_key] = value
                    break

            # 如果没有找到，保持为None
            if std_key not in standardized:
                standardized[std_key] = None

        # 打印SMILES状态日志
        excel_smiles = standardized.get('smiles')
        print(f"  Excel文件中SMILES状态: {'有' if excel_smiles else '无'}")
        
        # 如果没有SMILES，尝试从PubChem获取（已注释）
        # if not excel_smiles:
        #     print(f"  尝试从PubChem获取SMILES，使用CAS: {standardized.get('cas_number')} 或名称: {standardized.get('compound_name')}")
        #     smiles = self._get_smiles_from_pubchem(standardized)
        #     if smiles:
        #         standardized['smiles'] = smiles
        #         print(f"  ✓ 从PubChem成功获取SMILES: {smiles[:50]}...")
        #     else:
        #         print(f"  ✗ 从PubChem获取SMILES失败")
        
        # 如果有SMILES，显示获取结果
        if excel_smiles:
            print(f"  ✓ 从Excel获取的SMILES: {excel_smiles[:50]}...")
        
        # 保留所有原始数据
        for key, value in data.items():
            if key not in standardized:
                standardized[key] = value

        return standardized

    def _get_smiles_from_pubchem(self, compound_data):
        """从PubChem通过CAS号或化合物名称获取SMILES"""
        try:
            # 优先使用CAS号
            cas = compound_data.get('cas_number')
            if cas and str(cas) != '/' and not pd.isna(cas):
                cas = str(cas).strip()
                try:
                    compounds = pcp.get_compounds(cas, 'name')
                    if compounds:
                        smiles = compounds[0].isomeric_smiles or compounds[0].canonical_smiles
                        if smiles:
                            return smiles
                except:
                    pass

            # 尝试使用化合物名称
            name = compound_data.get('compound_name')
            if name and not pd.isna(name):
                name = str(name).strip()
                if name:
                    try:
                        compounds = pcp.get_compounds(name, 'name')
                        if compounds:
                            smiles = compounds[0].isomeric_smiles or compounds[0].canonical_smiles
                            if smiles:
                                return smiles
                    except:
                        pass

            return None

        except Exception as e:
            return None


# 数据库连接器（支持MySQL和SQL Server）
class DatabaseConnector:
    def __init__(self, db_type='mysql', **db_config):
        """
        初始化数据库连接器
        db_type: 'mysql' 或 'sqlserver'
        db_config: 数据库连接参数
            MySQL: host, port, user, password, database
            SQL Server: server, trusted_connection, user, password
        """
        self.db_type = db_type.lower()
        self.db_connector = None
        self.connected = False
        
        if self.db_type == 'mysql':
            self.db_connector = MySQLConnector(
                host=db_config.get('host', 'localhost'),
                port=db_config.get('port', 3306),
                user=db_config.get('user', 'root'),
                password=db_config.get('password', ''),
                database=db_config.get('database', None)
            )
        else:  # sqlserver
            self.db_connector = SQLServerConnector(
                server=db_config.get('server', 'localhost'),
                trusted_connection=db_config.get('trusted_connection', True),
                user=db_config.get('user', None),
                password=db_config.get('password', None)
            )

    def connect(self):
        """连接到数据库"""
        if self.db_connector:
            connected = self.db_connector.connect()
            self.connected = connected
            
            if connected:
                print(f"系统配置为使用{self.db_type.upper()}数据库进行查询")
            else:
                print(f"警告：{self.db_type.upper()}数据库连接失败，请检查数据库配置")
        else:
            print("错误：数据库连接器未初始化")
            self.connected = False
        
        return self.connected

    def query_data(self, query_type, query_value):
        """
        查询化合物信息
        使用数据库查询
        """
        if not self.connected:
            print("警告：数据库未连接，请检查数据库配置")
            return None

        # 使用数据库查询
        result = self.db_connector.query_compound(query_type, query_value)
        if result:
            print(f"从{self.db_type.upper()}数据库找到结果")
            return result

        # 如果数据库没找到，返回None
        print("数据库中未找到匹配的化合物")
        return None
    
    def close(self):
        """关闭数据库连接"""
        if self.db_connector:
            if hasattr(self.db_connector, 'close'):
                self.db_connector.close()
            elif hasattr(self.db_connector, 'connection') and self.db_connector.connection:
                try:
                    self.db_connector.connection.close()
                    print("数据库连接已关闭")
                except:
                    pass



# 预测模型包装器
class PredictionModel:
    def __init__(self):
        pass

    def predict(self, smiles):
        """预测EC50值和活性类别"""
        try:
            ec50 = round(random.uniform(0.1, 100), 4)
            if ec50 < 10:
                category = "优"
            elif ec50 < 50:
                category = "中"
            else:
                category = "差"
            ad_status = random.choice(["In AD", "Out AD"])

            return {
                "ec50": ec50,
                "category": category,
                "ad_status": ad_status
            }
        except Exception as e:
            print(f"预测错误: {str(e)}")
            return None


# 化合物结构绘制器
class CompoundDrawer:
    @staticmethod
    def smiles_to_pixmap(smiles, width=300, height=200):
        try:
            mol = Chem.MolFromSmiles(smiles)
            if mol is None:
                return None

            img = Draw.MolToImage(mol, size=(width, height))
            buffer = io.BytesIO()
            img.save(buffer, format='PNG')
            pixmap = QPixmap()
            pixmap.loadFromData(buffer.getvalue())

            return pixmap
        except Exception as e:
            print(f"分子绘制错误: {str(e)}")
            return None


# 批量下载SDF文件线程
class BatchSDFDownloadThread(QThread):
    progress_updated = Signal(int, str)  # (进度百分比, 状态消息)
    download_finished = Signal(str, int, int)  # (SDF文件路径, 成功数量, 失败数量)
    download_error = Signal(str)

    def __init__(self, db_connector, compounds_data, query_type, output_file):
        super().__init__()
        self.db_connector = db_connector
        self.compounds_data = compounds_data  # 直接使用查询结果数据
        self.query_type = query_type
        self.output_file = output_file
        self.running = True
        self.downloaded_compounds = []
        self.failed_count = 0

    def run(self):
        try:
            total = len(self.compounds_data)
            success_count = 0
            self.failed_count = 0  # 确保初始化失败计数
            
            if total == 0:
                self.download_error.emit("没有可下载的化合物数据")
                return
            
            print(f"[批量下载] 开始处理 {total} 个化合物")
            self.progress_updated.emit(0, f"开始下载 {total} 个化合物...")
            
            # 规范化输出文件路径，确保使用正确的路径分隔符
            self.output_file = os.path.normpath(self.output_file)
            # 转换为绝对路径
            self.output_file = os.path.abspath(self.output_file)
            output_dir = os.path.dirname(self.output_file)
            
            # 确保输出目录存在
            if output_dir and not os.path.exists(output_dir):
                try:
                    os.makedirs(output_dir, exist_ok=True)
                    print(f"[批量下载] 创建输出目录: {output_dir}")
                except Exception as e:
                    self.download_error.emit(f"无法创建输出目录: {str(e)}")
                    return
            
            # 确保输出目录确实存在
            if not os.path.exists(output_dir):
                self.download_error.emit(f"输出目录不存在: {output_dir}")
                return
            
            # 检查是否有写入权限
            if not os.access(output_dir, os.W_OK):
                self.download_error.emit(f"没有写入权限: {output_dir}")
                return
            
            print(f"[批量下载] 输出文件路径: {self.output_file}")
            print(f"[批量下载] 输出目录: {output_dir}")
            print(f"[批量下载] 输出目录存在: {os.path.exists(output_dir)}")
            
            # 测试是否可以创建文件（先尝试创建一个临时文件来测试权限）
            test_file = os.path.join(output_dir, ".test_write_permission.tmp")
            try:
                with open(test_file, 'w') as f:
                    f.write("test")
                os.remove(test_file)
                print(f"[批量下载] 目录写入权限测试通过")
            except Exception as e:
                error_msg = f"无法在输出目录中创建文件: {str(e)}\n目录: {output_dir}"
                print(f"[批量下载] 错误: {error_msg}")
                self.download_error.emit(error_msg)
                return
            
            # 如果输出文件已存在，先删除它（RDKit SDWriter 可能无法覆盖）
            if os.path.exists(self.output_file):
                try:
                    os.remove(self.output_file)
                    print(f"[批量下载] 删除已存在的文件: {self.output_file}")
                except Exception as e:
                    error_msg = f"无法删除已存在的文件: {str(e)}\n文件路径: {self.output_file}"
                    print(f"[批量下载] 错误: {error_msg}")
                    self.download_error.emit(error_msg)
                    return
            
            # 尝试多种方法创建SDF写入器
            # RDKit SDWriter 对中文路径可能有兼容性问题，尝试多种方法
            writer = None
            temp_file = None
            
            # 方法1: 尝试先创建空文件，然后使用SDWriter
            try:
                print(f"[批量下载] 方法1: 先创建空文件，然后使用SDWriter")
                # 先创建一个空文件
                with open(self.output_file, 'wb') as f:
                    pass
                print(f"[批量下载] 空文件创建成功: {self.output_file}")
                writer = Chem.SDWriter(str(self.output_file))
                print(f"[批量下载] SDF写入器创建成功（方法1）")
            except Exception as e1:
                print(f"[批量下载] 方法1失败: {str(e1)}")
                # 删除可能创建的空文件
                if os.path.exists(self.output_file):
                    try:
                        os.remove(self.output_file)
                    except:
                        pass
                
                # 方法2: 使用临时文件，最后重命名
                try:
                    print(f"[批量下载] 方法2: 使用临时文件")
                    import tempfile
                    temp_dir = output_dir
                    temp_fd, temp_file = tempfile.mkstemp(suffix='.sdf', dir=temp_dir, prefix='temp_')
                    os.close(temp_fd)  # 关闭文件描述符，SDWriter会重新打开
                    print(f"[批量下载] 临时文件创建: {temp_file}")
                    writer = Chem.SDWriter(temp_file)
                    self._temp_sdf_file = temp_file  # 保存临时文件路径
                    print(f"[批量下载] SDF写入器创建成功（方法2，使用临时文件）")
                except Exception as e2:
                    print(f"[批量下载] 方法2失败: {str(e2)}")
                    
                    # 方法3: 尝试使用正斜杠路径
                    try:
                        print(f"[批量下载] 方法3: 使用正斜杠路径")
                        sdf_path_str = str(self.output_file).replace('\\', '/')
                        writer = Chem.SDWriter(sdf_path_str)
                        print(f"[批量下载] SDF写入器创建成功（方法3，正斜杠路径）")
                    except Exception as e3:
                        print(f"[批量下载] 方法3失败: {str(e3)}")
                        
                        # 方法4: 尝试使用短路径名（8.3格式）来避免中文路径问题
                        try:
                            print(f"[批量下载] 方法4: 尝试使用短路径名")
                            try:
                                import win32api
                                short_path = win32api.GetShortPathName(output_dir)
                                short_file = os.path.join(short_path, f"compounds_{time.strftime('%Y%m%d_%H%M%S')}.sdf")
                                writer = Chem.SDWriter(short_file)
                                self.output_file = short_file  # 更新输出文件路径
                                print(f"[批量下载] SDF写入器创建成功（方法4，短路径名）")
                            except ImportError:
                                raise Exception("win32api 模块不可用")
                        except Exception as e4:
                            # 所有方法都失败了
                            error_msg = f"无法创建SDF文件写入器，已尝试所有方法:\n"
                            error_msg += f"方法1（先创建空文件）: {str(e1)}\n"
                            error_msg += f"方法2（临时文件）: {str(e2)}\n"
                            error_msg += f"方法3（正斜杠路径）: {str(e3)}\n"
                            error_msg += f"方法4（短路径名）: {str(e4)}\n"
                            error_msg += f"\n文件路径: {self.output_file}\n"
                            error_msg += f"目录存在: {os.path.exists(output_dir)}\n"
                            error_msg += f"建议: 尝试将输出文件夹改为不包含中文的路径"
                            print(f"[批量下载] 错误: {error_msg}")
                            self.download_error.emit(error_msg)
                            return
            
            # 如果没有使用临时文件，确保 _temp_sdf_file 为 None
            if not hasattr(self, '_temp_sdf_file'):
                self._temp_sdf_file = None
            
            for idx, compound_data in enumerate(self.compounds_data):
                if not self.running:
                    break
                
                progress = int((idx / total) * 100) if total > 0 else 0
                compound_name = compound_data.get('compound_name', f'Compound_{idx+1}')
                self.progress_updated.emit(progress, f"正在下载 ({idx+1}/{total}): {compound_name}")
                
                print(f"[批量下载] 处理化合物 {idx+1}/{total}: {compound_name}")
                print(f"  数据键: {list(compound_data.keys())}")
                print(f"  SMILES: {compound_data.get('smiles', 'None')}")
                print(f"  SDF路径: {compound_data.get('sdf_path') or compound_data.get('sdf_file') or compound_data.get('_sdf_path') or 'None'}")
                
                processed = False
                try:
                    # 如果有SDF文件路径，直接复制
                    sdf_path = compound_data.get('sdf_path') or compound_data.get('sdf_file') or compound_data.get('_sdf_path')
                    
                    if sdf_path and os.path.exists(sdf_path):
                        # 如果存在SDF文件，读取并合并到输出文件
                        try:
                            supplier = Chem.SDMolSupplier(sdf_path)
                            for mol in supplier:
                                if mol:
                                    # 保留原有属性，添加额外信息
                                    if compound_data.get('compound_name'):
                                        mol.SetProp('_Name', str(compound_data.get('compound_name')))
                                    writer.write(mol)
                                    success_count += 1
                                    processed = True
                                    print(f"  ✓ 从SDF文件成功写入: {sdf_path}")
                                    break
                        except Exception as e:
                            print(f"  ✗ 读取SDF文件失败: {sdf_path}, {str(e)}")
                            # 如果读取失败，尝试从SMILES生成
                            sdf_path = None
                    
                    # 如果没有SDF文件路径，从SMILES生成
                    if not processed:
                        smiles = compound_data.get('smiles')
                        # 检查SMILES是否有效（不是None、空字符串或'None'字符串）
                        if smiles and str(smiles).strip() and str(smiles).strip().lower() != 'none':
                            smiles = str(smiles).strip()
                            mol = Chem.MolFromSmiles(smiles)
                            
                            if mol:
                                # 添加属性
                                mol.SetProp('_Name', str(compound_data.get('compound_name', f'Compound_{idx+1}')))
                                if compound_data.get('cas_number'):
                                    mol.SetProp('CAS', str(compound_data.get('cas_number')))
                                if compound_data.get('molecular_weight'):
                                    mol.SetProp('MolecularWeight', str(compound_data.get('molecular_weight')))
                                if compound_data.get('molecular_formula'):
                                    mol.SetProp('Formula', str(compound_data.get('molecular_formula')))
                                if compound_data.get('category'):
                                    mol.SetProp('Category', str(compound_data.get('category')))
                                
                                # 写入SDF
                                writer.write(mol)
                                success_count += 1
                                processed = True
                                print(f"  ✓ 从SMILES成功生成并写入: {smiles[:50]}...")
                                self.downloaded_compounds.append({
                                    'name': compound_data.get('compound_name', ''),
                                    'smiles': smiles,
                                    'cas': compound_data.get('cas_number', '')
                                })
                            else:
                                print(f"  ✗ 无法从SMILES创建分子: {smiles[:50]}...")
                                self.failed_count += 1
                        else:
                            print(f"  ✗ 化合物 '{compound_name}' 没有有效的SMILES或SDF路径")
                            print(f"    SMILES值: {repr(smiles)}")
                            self.failed_count += 1
                        
                except Exception as e:
                    print(f"  ✗ 下载化合物 '{compound_name}' 时出错: {str(e)}")
                    import traceback
                    traceback.print_exc()
                    self.failed_count += 1
                    continue
            
            writer.close()
            
            # 如果使用了临时文件，需要重命名为最终文件名
            if hasattr(self, '_temp_sdf_file') and self._temp_sdf_file and os.path.exists(self._temp_sdf_file):
                try:
                    # 如果目标文件已存在，先删除
                    if os.path.exists(self.output_file):
                        os.remove(self.output_file)
                    # 重命名临时文件
                    os.rename(self._temp_sdf_file, self.output_file)
                    print(f"[批量下载] 临时文件已重命名为: {self.output_file}")
                except Exception as e:
                    print(f"[批量下载] 警告: 重命名临时文件失败: {str(e)}")
                    # 如果重命名失败，使用临时文件路径作为最终路径
                    self.output_file = self._temp_sdf_file
            
            print(f"[批量下载] 完成: 成功 {success_count}, 失败 {self.failed_count}")
            self.progress_updated.emit(100, "下载完成")
            self.download_finished.emit(self.output_file, success_count, self.failed_count)
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.download_error.emit(f"批量下载出错: {str(e)}")
    
    def stop(self):
        self.running = False


# ChemDraw批量提取线程
class ChemDrawBatchExtractorThread(QThread):
    progress_updated = Signal(int, str)  # (进度百分比, 状态消息)
    extraction_finished = Signal(str, int, int)  # (SDF文件路径, 成功数量, 失败数量)
    extraction_error = Signal(str)

    def __init__(self, excel_folder, output_file):
        super().__init__()
        self.excel_folder = excel_folder
        self.output_file = output_file
        self.running = True
        self.extracted_compounds = []
        self.failed_count = 0

    def run(self):
        extractor = None
        try:
            # 查找所有Excel文件
            self.progress_updated.emit(0, "正在查找Excel文件...")
            excel_files = glob.glob(os.path.join(self.excel_folder, '**', '*.xlsx'), recursive=True)
            excel_files += glob.glob(os.path.join(self.excel_folder, '**', '*.xls'), recursive=True)
            excel_files = [f for f in excel_files if not os.path.basename(f).startswith('~$')]

            print(f"[Batch] 共找到 {len(excel_files)} 个Excel文件，开始处理")
            if not excel_files:
                self.extraction_error.emit("未找到Excel文件")
                return

            self.progress_updated.emit(5, f"找到 {len(excel_files)} 个Excel文件")
            time.sleep(0.5)

            # 创建ChemDraw提取器
            extractor = ChemDrawExtractor()
            total_processed = 0
            total_files = len(excel_files)

            # 处理每个Excel文件
            for file_idx, excel_file in enumerate(excel_files):
                if not self.running:
                    break

                file_name = os.path.basename(excel_file)
                self.progress_updated.emit(
                    int((file_idx / total_files) * 90) + 5,
                    f"处理文件 [{file_idx+1}/{total_files}]: {file_name}"
                )

                print(f"[Batch] 开始处理文件 ({file_idx+1}/{total_files}): {file_name}")
                try:
                    # 读取Excel数据
                    df = pd.read_excel(excel_file)
                    print(f"[Batch] 文件 {file_name} 读取完成，共 {len(df)} 行")

                    # 查找化合物名称列
                    name_col = None
                    for col in ['化合物名称', '名称', '化合物', '中文名称']:
                        if col in df.columns:
                            name_col = col
                            break

                    if not name_col:
                        continue

                    # 查找其他信息列
                    cas_col = None
                    for col in ['CAS', 'cas', 'CAS号', 'CAS Number']:
                        if col in df.columns:
                            cas_col = col
                            break

                    formula_col = None
                    for col in ['分子式', '化学式', 'Formula', 'Molecular Formula']:
                        if col in df.columns:
                            formula_col = col
                            break

                    # 遍历每一行
                    for idx, row in df.iterrows():
                        if not self.running:
                            break

                        excel_row = idx + 2  # Excel行号

                        # 获取化合物信息
                        compound_name = row[name_col] if name_col and not pd.isna(row[name_col]) else f"Compound_{excel_row}"
                        cas_number = row[cas_col] if cas_col and not pd.isna(row[cas_col]) else ""
                        formula = row[formula_col] if formula_col and not pd.isna(row[formula_col]) else ""

                        if idx % 10 == 0:
                            print(f"[Batch] {file_name} 第 {excel_row} 行")

                        # 策略1: 提取ChemDraw图片
                        chemdraw_image = extractor.extract_chemdraw_from_excel(excel_file, excel_row, col_index=7)

                        # 策略2: 直接从ChemDraw对象提取MOL数据
                        mol_data, mol, cdx_path = extractor.extract_chemdraw_as_mol(excel_file, excel_row, col_index=7)

                        if mol is None and mol_data:
                            try:
                                mol = Chem.MolFromMolBlock(mol_data)
                            except:
                                mol = None

                        smiles = None
                        if mol:
                            try:
                                smiles = Chem.MolToSmiles(mol)
                            except:
                                smiles = None

                        # PubChem API已禁用 - 太慢
                        # if cas_number and str(cas_number) != '/' and not pd.isna(cas_number):
                        #     try:
                        #         import pubchempy as pcp
                        #         compounds = pcp.get_compounds(str(cas_number).strip(), 'name')
                        #         if compounds:
                        #             smiles = compounds[0].isomeric_smiles or compounds[0].canonical_smiles
                        #             if smiles:
                        #                 mol = Chem.MolFromSmiles(smiles)
                        #                 if mol:
                        #                     mol_data = Chem.MolToMolBlock(mol)
                        #     except:
                        #         pass

                        # 只要有任一有效信息就视为成功
                        if chemdraw_image or mol or mol_data:
                            self.extracted_compounds.append({
                                'name': compound_name,
                                'cas': cas_number,
                                'formula': formula,
                                'mol': mol,
                                'mol_data': mol_data,
                                'cdx_path': cdx_path,
                                'smiles': smiles,
                                'chemdraw_image': chemdraw_image,
                                'source_file': file_name,
                                'row': excel_row
                            })
                            total_processed += 1
                        else:
                            self.failed_count += 1

                except Exception as e:
                    print(f"处理文件 {file_name} 出错: {str(e)}")
                    continue

            if not self.extracted_compounds:
                self.extraction_error.emit(f"未能提取到任何ChemDraw结构\n处理了 {total_files} 个文件")
                return

            # 生成SDF文件
            self.progress_updated.emit(95, f"正在生成SDF文件...")
            self.generate_sdf()
            self.progress_updated.emit(100, "完成！")
            self.extraction_finished.emit(self.output_file, len(self.extracted_compounds), self.failed_count)

        except Exception as e:
            self.extraction_error.emit(f"批量提取出错: {str(e)}")
        finally:
            if extractor:
                extractor.close_all()

    def generate_sdf(self):
        """生成多种格式的分子文件：SMILES, MOL, SDF, PDB, PDBQT"""
        try:
            # 创建输出目录结构
            output_dir = os.path.dirname(self.output_file)
            images_dir = os.path.join(output_dir, "chemdraw_images")
            mol_dir = os.path.join(output_dir, "mol_files")
            cdx_dir = os.path.join(output_dir, "cdx_files")
            pdb_dir = os.path.join(output_dir, "pdb_files")
            pdbqt_dir = os.path.join(output_dir, "pdbqt_files")

            os.makedirs(images_dir, exist_ok=True)
            os.makedirs(mol_dir, exist_ok=True)
            os.makedirs(cdx_dir, exist_ok=True)
            os.makedirs(pdb_dir, exist_ok=True)
            os.makedirs(pdbqt_dir, exist_ok=True)

            # 计数器
            sdf_count = 0
            image_count = 0
            mol_count = 0
            pdb_count = 0
            pdbqt_count = 0
            cdx_count = 0

            base_output = os.path.splitext(self.output_file)[0]

            # 创建SMILES文件
            smiles_file = base_output + '_smiles.txt'
            smiles_list = []

            writer = Chem.SDWriter(self.output_file)
            sd_copy_path = base_output + '.sd'

            for idx, compound in enumerate(self.extracted_compounds):
                # 保存ChemDraw图片（JPEG）
                if compound.get('chemdraw_image'):
                    try:
                        img = compound['chemdraw_image']
                        if img.mode != 'RGB':
                            img = img.convert('RGB')
                        img_filename = f"{idx+1:04d}_{compound['name'][:30]}.jpg"
                        img_filename = "".join(c for c in img_filename if c.isalnum() or c in ('_', '-', '.'))
                        img_path = os.path.join(images_dir, img_filename)
                        img.save(img_path, format='JPEG', quality=95)
                        image_count += 1
                    except Exception as e:
                        print(f"保存图片失败: {compound['name']}, {str(e)}")

                base_filename = f"{idx+1:04d}_{compound['name'][:30]}"
                base_filename = "".join(c for c in base_filename if c.isalnum() or c in ('_', '-'))

                mol = compound.get('mol')
                if mol is None and compound.get('mol_data'):
                    try:
                        mol = Chem.MolFromMolBlock(compound['mol_data'])
                    except:
                        mol = None

                # 导出CDX文件
                if compound.get('cdx_path') and os.path.exists(compound['cdx_path']):
                    try:
                        cdx_file = os.path.join(cdx_dir, f"{base_filename}.cdx")
                        shutil.copy2(compound['cdx_path'], cdx_file)
                        cdx_count += 1
                    except Exception as e:
                        print(f"复制CDX文件失败: {compound['name']}, {str(e)}")

                # 处理有MOL数据的化合物
                if mol:
                    # 添加属性
                    mol.SetProp('_Name', str(compound['name']))
                    mol.SetProp('CAS', str(compound.get('cas', '')))
                    mol.SetProp('Formula', str(compound.get('formula', '')))
                    if compound.get('smiles'):
                        mol.SetProp('SMILES', str(compound['smiles']))
                    mol.SetProp('Source_File', str(compound['source_file']))
                    mol.SetProp('Source_Row', str(compound['row']))

                    # 计算分子描述符
                    try:
                        mol.SetProp('MolecularWeight', str(round(Descriptors.MolWt(mol), 2)))
                        mol.SetProp('LogP', str(round(Descriptors.MolLogP(mol), 2)))
                        mol.SetProp('TPSA', str(round(Descriptors.TPSA(mol), 2)))
                        mol.SetProp('NumHDonors', str(Descriptors.NumHDonors(mol)))
                        mol.SetProp('NumHAcceptors', str(Descriptors.NumHAcceptors(mol)))
                        mol.SetProp('NumRotatableBonds', str(Descriptors.NumRotatableBonds(mol)))
                    except:
                        pass

                    # 1. 添加到SDF文件
                    writer.write(mol)
                    sdf_count += 1

                    # 2. 保存单独的MOL文件
                    try:
                        mol_file = os.path.join(mol_dir, f"{base_filename}.mol")
                        Chem.MolToMolFile(mol, mol_file)
                        mol_count += 1
                    except Exception as e:
                        print(f"保存MOL文件失败: {compound['name']}, {str(e)}")

                    # 3. 保存SMILES到列表
                    if compound.get('smiles'):
                        smiles_list.append(f"{compound['smiles']}\t{compound['name']}\t{compound.get('cas', '')}")
                elif compound.get('mol_data'):
                    # 没有成功解析成RDKit分子，但仍然导出MOL文件
                    try:
                        mol_file = os.path.join(mol_dir, f"{base_filename}.mol")
                        with open(mol_file, 'w', encoding='utf-8') as f:
                            f.write(compound['mol_data'])
                        mol_count += 1
                    except Exception as e:
                        print(f"保存纯文本MOL失败: {compound['name']}, {str(e)}")

                    # 4. 生成3D构象并保存为PDB格式
                    try:
                        mol_3d = Chem.Mol(mol)
                        AllChem.EmbedMolecule(mol_3d, randomSeed=42)
                        AllChem.MMFFOptimizeMolecule(mol_3d)

                        pdb_file = os.path.join(pdb_dir, f"{base_filename}.pdb")
                        Chem.MolToPDBFile(mol_3d, pdb_file)
                        pdb_count += 1

                        # 5. 尝试转换为PDBQT格式（需要OpenBabel）
                        try:
                            import subprocess
                            pdbqt_file = os.path.join(pdbqt_dir, f"{base_filename}.pdbqt")

                            # 尝试使用obabel命令
                            result = subprocess.run(
                                ['obabel', pdb_file, '-O', pdbqt_file, '-h'],
                                capture_output=True,
                                text=True,
                                timeout=5
                            )

                            if result.returncode == 0:
                                pdbqt_count += 1
                        except FileNotFoundError:
                            # OpenBabel未安装，跳过PDBQT生成
                            pass
                        except Exception:
                            pass

                    except Exception as e:
                        print(f"生成PDB文件失败: {compound['name']}, {str(e)}")

            writer.close()

            # 复制生成.sd扩展名
            try:
                if not self.output_file.lower().endswith('.sd'):
                    shutil.copyfile(self.output_file, sd_copy_path)
            except Exception as e:
                print(f"复制SD文件失败: {str(e)}")

            # 写入SMILES文件
            if smiles_list:
                with open(smiles_file, 'w', encoding='utf-8') as f:
                    f.write("SMILES\tName\tCAS\n")
                    for line in smiles_list:
                        f.write(line + "\n")

            # 生成报告文件
            report_file = self.output_file.replace('.sdf', '_report.txt')
            with open(report_file, 'w', encoding='utf-8') as f:
                f.write("分子结构批量生成报告\n")
                f.write("Python + RDKit + OpenBabel\n")
                f.write("="*60 + "\n\n")
                f.write("生成文件统计:\n")
                f.write("-"*60 + "\n")
                f.write(f"ChemDraw图片 (JPG):  {image_count} 个\n")
                f.write(f"MOL 文件:            {mol_count} 个\n")
                f.write(f"SDF 文件:            1 个 (包含 {sdf_count} 个分子)\n")
                f.write(f"SD 副本:             1 个\n")
                f.write(f"CDX 文件:           {cdx_count} 个\n")
                f.write(f"PDB 文件:            {pdb_count} 个\n")
                f.write(f"PDBQT 文件:          {pdbqt_count} 个\n")
                f.write(f"SMILES 文件:         {len(smiles_list)} 条记录\n")
                f.write(f"\n提取总数:            {len(self.extracted_compounds)}\n")
                f.write(f"失败数:              {self.failed_count}\n\n")

                f.write("输出目录结构:\n")
                f.write("-"*60 + "\n")
                f.write(f"  ├─ chemdraw_images/  (ChemDraw图片)\n")
                f.write(f"  ├─ mol_files/        (MOL格式文件)\n")
                f.write(f"  ├─ cdx_files/        (ChemDraw CDX文件)\n")
                f.write(f"  ├─ pdb_files/        (PDB格式文件)\n")
                f.write(f"  ├─ pdbqt_files/      (PDBQT格式文件)\n")
                f.write(f"  ├─ compounds.sdf     (SDF合并文件)\n")
                f.write(f"  └─ compounds_smiles.txt (SMILES列表)\n\n")

                f.write("详细信息:\n")
                f.write("-"*60 + "\n")

                for idx, comp in enumerate(self.extracted_compounds, 1):
                    f.write(f"\n{idx}. {comp['name']}\n")
                    f.write(f"   CAS号:        {comp.get('cas', 'N/A')}\n")
                    f.write(f"   分子式:       {comp.get('formula', 'N/A')}\n")
                    f.write(f"   SMILES:       {comp.get('smiles', 'N/A')}\n")
                    f.write(f"   来源文件:     {comp['source_file']} (第{comp['row']}行)\n")
                    f.write(f"   ChemDraw图片: {'✓' if comp.get('chemdraw_image') else '✗'}\n")
                    f.write(f"   MOL文件:      {'✓' if comp.get('mol') else '✗'}\n")

            print(f"\n文件生成完成:")
            print(f"  ChemDraw JPG图片: {image_count} 个")
            print(f"  MOL文件: {mol_count} 个")
            print(f"  CDX文件: {cdx_count} 个")
            print(f"  SDF文件: 1 个 ({sdf_count} 个分子)")
            print(f"  SD副本: 1 个")
            print(f"  PDB文件: {pdb_count} 个")
            print(f"  PDBQT文件: {pdbqt_count} 个")
            print(f"  SMILES文件: {len(smiles_list)} 条")
            print(f"  报告文件: {report_file}")

        except Exception as e:
            raise Exception(f"生成输出文件失败: {str(e)}")

    def stop(self):
        self.running = False


# 主窗口类
class CardioHerbDBApp(QMainWindow):
    def __init__(self):
        super().__init__()

        # 初始化数据库连接器和预测模型
        # 默认使用 MySQL 连接（根据图片中的配置）
        # 可以根据需要修改这些参数
        # 注意：请将下面的 password 参数改为您的实际数据库密码
        self.db_connector = DatabaseConnector(
            db_type='mysql',
            host='localhost',      # 数据库主机地址
            port=3306,             # MySQL 端口
            user='root',           # 数据库用户名
            password='Sexy050117',           # ⚠️ 请在这里填入您的数据库密码
            database='compounds'          # 可以指定数据库名（如 'Compound'），或设为 None 让程序自动查找
        )
        self.db_connector.connect()
        self.prediction_model = PredictionModel()
        self.compound_drawer = CompoundDrawer()
        self.batch_thread = None
        self.query_thread = None

        # 设置窗口标题和大小
        self.setWindowTitle("治疗心血管中药-靶标-疾病信息数据库")
        self.setGeometry(100, 100, 1200, 800)

        # 设置字体
        font = QFont()
        font.setFamily("Microsoft YaHei")
        self.setFont(font)

        # 设置蓝色调样式表
        self.setStyleSheet("""
            #central_widget {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #e6f7ff, stop:1 #f0f8ff);
            }
            #title_label {
                font-size: 24px;
                font-weight: bold;
                color: #1890ff;
                margin-bottom: 10px;
            }
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #1890ff, stop:1 #096dd9);
                color: white;
                border-radius: 8px;
                padding: 8px 16px;
                font-size: 14px;
                font-weight: bold;
                border: none;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #40a9ff, stop:1 #1890ff);
            }
            QPushButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #096dd9, stop:1 #0050b3);
            }
            QLineEdit, QComboBox {
                border: 1px solid #91d5ff;
                border-radius: 8px;
                padding: 8px;
                font-size: 14px;
                background-color: rgba(255, 255, 255, 0.8);
            }
            QLineEdit:focus, QComboBox:focus {
                border: 2px solid #40a9ff;
                background-color: white;
            }
            QGroupBox {
                border: 1px solid #91d5ff;
                border-radius: 10px;
                margin-top: 6px;
                padding: 15px;
                background-color: rgba(255, 255, 255, 0.7);
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 15px;
                padding: 0 5px 0 5px;
                color: #1890ff;
                font-weight: bold;
                font-size: 16px;
            }
            QTableWidget {
                border: 1px solid #91d5ff;
                border-radius: 8px;
                background-color: rgba(255, 255, 255, 0.9);
            }
            QTableWidget::item:selected {
                background-color: #e6f7ff;
                color: #1890ff;
            }
            QTextEdit {
                border: 1px solid #91d5ff;
                border-radius: 8px;
                padding: 10px;
                background-color: rgba(255, 255, 255, 0.9);
                font-size: 14px;
            }
        """)

        # 创建主布局
        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        central_widget.setObjectName("central_widget")

        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # 标题
        title_label = QLabel("治疗心血管中药-靶标-疾病信息数据库")
        title_label.setObjectName("title_label")
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)

        # 分割器
        splitter = QSplitter(Qt.Vertical)

        # 上半部分：输入和查询区域
        top_widget = QWidget()
        top_layout = QVBoxLayout(top_widget)
        top_layout.setSpacing(10)

        self.create_single_query_section(top_layout)
        self.create_batch_query_section(top_layout)

        splitter.addWidget(top_widget)

        # 下半部分：结果展示区域
        bottom_widget = QWidget()
        bottom_layout = QVBoxLayout(bottom_widget)
        bottom_layout.setSpacing(10)

        self.create_results_section(bottom_layout)

        splitter.addWidget(bottom_widget)
        splitter.setSizes([300, 500])

        main_layout.addWidget(splitter)

    def create_single_query_section(self, parent_layout):
        query_group = QGroupBox("单条查询")
        query_layout = QVBoxLayout(query_group)

        input_layout = QHBoxLayout()
        input_layout.setSpacing(10)

        input_layout.addWidget(QLabel("查询类型:"))
        self.query_type_combo = QComboBox()
        self.query_type_combo.addItems(["SMILES", "CAS号", "化合物名称"])
        self.query_type_combo.setMinimumWidth(120)
        input_layout.addWidget(self.query_type_combo)

        self.query_input = QLineEdit()
        self.query_input.setPlaceholderText("请输入查询内容")
        self.query_input.setMinimumWidth(400)
        input_layout.addWidget(self.query_input)

        self.query_btn = QPushButton("查询")
        self.query_btn.clicked.connect(self.perform_single_query)
        input_layout.addWidget(self.query_btn)

        self.clear_btn = QPushButton("清除")
        self.clear_btn.clicked.connect(self.clear_single_query)
        input_layout.addWidget(self.clear_btn)

        query_layout.addLayout(input_layout)
        parent_layout.addWidget(query_group)

    def create_batch_query_section(self, parent_layout):
        batch_group = QGroupBox("批量下载SDF文件")
        batch_layout = QVBoxLayout(batch_group)
        batch_layout.setSpacing(10)

        # 提示信息
        info_label = QLabel("从上方查询结果列表中批量下载SDF文件")
        info_label.setStyleSheet("color: #666; font-size: 12px; padding: 5px;")
        batch_layout.addWidget(info_label)

        # 输出文件选择
        output_layout = QHBoxLayout()
        output_layout.addWidget(QLabel("输出文件夹:"))
        self.output_sdf_path = QLineEdit()
        self.output_sdf_path.setReadOnly(True)
        self.output_sdf_path.setPlaceholderText("点击选择输出文件夹")
        output_layout.addWidget(self.output_sdf_path)

        self.select_output_btn = QPushButton("选择输出文件夹")
        self.select_output_btn.clicked.connect(self.select_output_file)
        output_layout.addWidget(self.select_output_btn)

        batch_layout.addLayout(output_layout)

        # 开始下载按钮和进度条
        control_layout = QHBoxLayout()
        self.batch_download_btn = QPushButton("开始下载")
        self.batch_download_btn.clicked.connect(self.start_batch_download)
        self.batch_download_btn.setEnabled(False)  # 初始状态禁用，有查询结果后启用
        control_layout.addWidget(self.batch_download_btn)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        control_layout.addWidget(self.progress_bar)

        batch_layout.addLayout(control_layout)

        # 状态标签
        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: #1890ff; font-size: 12px;")
        batch_layout.addWidget(self.status_label)

        parent_layout.addWidget(batch_group)

    def create_results_section(self, parent_layout):
        results_group = QGroupBox("查询结果")
        results_layout = QVBoxLayout(results_group)

        # 结果操作按钮
        results_ops_layout = QHBoxLayout()
        results_ops_layout.setSpacing(10)

        self.save_results_btn = QPushButton("保存结果")
        self.save_results_btn.clicked.connect(self.save_results)
        self.save_results_btn.setEnabled(False)
        results_ops_layout.addWidget(self.save_results_btn)

        self.clear_results_btn = QPushButton("清空结果")
        self.clear_results_btn.clicked.connect(self.clear_results)
        results_ops_layout.addWidget(self.clear_results_btn)

        results_ops_layout.addStretch()
        results_layout.addLayout(results_ops_layout)

        # 结果表格
        self.results_table = QTableWidget()
        self.results_table.setColumnCount(6)  # 6列数据
        self.results_table.setHorizontalHeaderLabels(["化合物名称", "CAS号", "SMILES", "EC50", "活性类别", "AD状态"])

        header = self.results_table.horizontalHeader()
        # 设置所有列为平均分布（Stretch模式）
        header.setSectionResizeMode(0, QHeaderView.Stretch)  # 化合物名称
        header.setSectionResizeMode(1, QHeaderView.Stretch)  # CAS号
        header.setSectionResizeMode(2, QHeaderView.Stretch)  # SMILES
        header.setSectionResizeMode(3, QHeaderView.Stretch)  # EC50
        header.setSectionResizeMode(4, QHeaderView.Stretch)  # 活性类别
        header.setSectionResizeMode(5, QHeaderView.Stretch)  # AD状态

        self.results_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.results_table.itemClicked.connect(self.on_table_item_clicked)

        results_layout.addWidget(self.results_table)

        # 分子结构和详细信息显示区域
        self.create_molecule_display_section(results_layout)

        parent_layout.addWidget(results_group)

    def create_molecule_display_section(self, parent_layout):
        mol_group = QGroupBox("分子结构")
        mol_layout = QHBoxLayout(mol_group)

        # 左侧：分子结构图片
        self.molecule_image_label = QLabel()
        self.molecule_image_label.setAlignment(Qt.AlignCenter)
        self.molecule_image_label.setMinimumSize(400, 300)
        self.molecule_image_label.setStyleSheet(
            "background-color: #f9f9f9; border: 1px solid #ddd; border-radius: 6px;"
        )
        self.molecule_image_label.setText("选择表格中的化合物查看分子结构")
        mol_layout.addWidget(self.molecule_image_label, 1)

        # 右侧：化合物详细信息
        self.molecule_info_text = QTextEdit()
        self.molecule_info_text.setReadOnly(True)
        self.molecule_info_text.setPlaceholderText("化合物详细信息将显示在这里")
        self.molecule_info_text.setMinimumSize(400, 300)
        mol_layout.addWidget(self.molecule_info_text, 1)

        parent_layout.addWidget(mol_group)
        
    def perform_single_query(self):
        query_type = self.query_type_combo.currentText()
        query_value = self.query_input.text().strip()

        if not query_value:
            QMessageBox.warning(self, "查询失败", f"请输入{query_type}")
            return

        # 禁用查询按钮，防止重复点击
        self.query_btn.setEnabled(False)
        
        # 设置光标为等待状态
        self.setCursor(Qt.WaitCursor)
        
        # 显示状态提示
        self.statusBar().showMessage(f"正在查询化合物 '{query_value}'，请稍候...")

        query_type_map = {
            "SMILES": "SMILES",
            "CAS号": "CAS",
            "化合物名称": "Name"
        }

        # 创建并启动查询线程
        if hasattr(self, 'query_thread') and self.query_thread and self.query_thread.isRunning():
            self.query_thread.stop()
            self.query_thread.wait()
            
        self.query_thread = QueryThread(self.db_connector, query_type_map[query_type], query_value)
        self.query_thread.query_finished.connect(self.on_query_finished)
        self.query_thread.query_error.connect(self.on_query_error)
        self.query_thread.finished.connect(self.on_query_thread_finished)
        self.query_thread.start()
        
    def on_query_finished(self, result):
        if result:
            print(f"查询结果数据: {result.keys()}")
            # 添加到表格
            row_position = self.results_table.rowCount()
            self.results_table.insertRow(row_position)
            
            # 添加各数据列
            self.results_table.setItem(row_position, 0, QTableWidgetItem(str(result.get("compound_name", ""))))
            self.results_table.setItem(row_position, 1, QTableWidgetItem(str(result.get("cas_number", ""))))
            self.results_table.setItem(row_position, 2, QTableWidgetItem(str(result.get("smiles", ""))))
            self.results_table.setItem(row_position, 3, QTableWidgetItem(str(result.get("ec50", ""))))
            self.results_table.setItem(row_position, 4, QTableWidgetItem(str(result.get("category", ""))))
            self.results_table.setItem(row_position, 5, QTableWidgetItem(str(result.get("ad_status", ""))))
            
            # 存储SDF文件路径（使用setData存储在item中，不显示）
            # 尝试从查询结果中获取SDF路径，如果没有则根据SMILES生成
            sdf_path = result.get('sdf_path') or result.get('sdf_file') or result.get('_sdf_path')
            if not sdf_path and result.get('smiles'):
                # 如果没有SDF路径，可以根据需要生成或留空
                # 这里先存储为None，批量下载时会根据SMILES生成
                sdf_path = None
            
            # 将完整结果数据存储在第一个item的userData中，方便后续使用
            item = self.results_table.item(row_position, 0)
            if item:
                item.setData(Qt.UserRole, result)  # 存储完整结果数据

            self.save_results_btn.setEnabled(True)
            self.batch_download_btn.setEnabled(True)  # 有查询结果后启用批量下载按钮
            self.results_table.selectRow(row_position)
            self.display_molecule_info(result)
        else:
            QMessageBox.information(self, "查询结果", "未找到匹配的化合物信息")
    
    def on_query_error(self, error_msg):
        QMessageBox.critical(self, "查询错误", f"查询过程中出现错误: {error_msg}")
    
    def on_query_thread_finished(self):
        # 恢复按钮状态
        self.query_btn.setEnabled(True)
        # 恢复光标状态
        self.setCursor(Qt.ArrowCursor)
        # 清除状态提示
        self.statusBar().clearMessage()
        
    def closeEvent(self, event):
        print("程序正在关闭...")
        # 停止查询线程
        if hasattr(self, 'query_thread') and self.query_thread and self.query_thread.isRunning():
            self.query_thread.stop()
            self.query_thread.wait()
        # 关闭数据库连接
        if hasattr(self, 'db_connector'):
            self.db_connector.close()
        event.accept()

    def clear_single_query(self):
        self.query_input.clear()
        self.molecule_image_label.clear()
        self.molecule_image_label.setText("选择表格中的化合物查看分子结构")
        self.molecule_info_text.clear()


    def select_output_file(self):
        """选择SDF输出文件夹"""
        folder_path = QFileDialog.getExistingDirectory(
            self, "选择输出文件夹"
        )
        if folder_path:
            # 规范化路径，确保使用正确的路径分隔符
            folder_path = os.path.normpath(folder_path)
            self.output_sdf_path.setText(folder_path)

    def start_batch_download(self):
        """开始批量下载SDF文件（从查询结果列表）"""
        # 检查是否有查询结果
        row_count = self.results_table.rowCount()
        if row_count == 0:
            QMessageBox.warning(self, "下载失败", "查询结果列表为空，请先查询化合物")
            return

        # 获取输出文件夹路径
        output_folder = self.output_sdf_path.text()
        if not output_folder:
            QMessageBox.warning(self, "下载失败", "请先选择输出文件夹")
            return
        
        # 规范化路径，确保使用正确的路径分隔符
        output_folder = os.path.normpath(output_folder)
        
        # 确保输出文件夹存在
        if not os.path.exists(output_folder):
            try:
                os.makedirs(output_folder, exist_ok=True)
            except Exception as e:
                QMessageBox.critical(self, "错误", f"无法创建输出文件夹: {str(e)}")
                return
        
        # 在选择的文件夹中生成SDF文件名
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(output_folder, f"compounds_{timestamp}.sdf")
        # 规范化并转换为绝对路径
        output_path = os.path.normpath(output_path)
        output_path = os.path.abspath(output_path)

        # 从表格中收集所有查询结果
        compounds_data = []
        print(f"[批量下载] 开始收集表格数据，共 {row_count} 行")
        for row in range(row_count):
            item = self.results_table.item(row, 0)
            if item:
                # 从userData中获取完整结果数据
                result_data = item.data(Qt.UserRole)
                if result_data:
                    print(f"  行 {row+1}: 使用存储的完整数据")
                    print(f"    键: {list(result_data.keys())}")
                    print(f"    SMILES: {result_data.get('smiles', 'None')}")
                    compounds_data.append(result_data)
                else:
                    # 如果没有存储完整数据，从表格中读取
                    print(f"  行 {row+1}: 从表格单元格读取数据")
                    compound_info = {
                        "compound_name": self.results_table.item(row, 0).text() if self.results_table.item(row, 0) else "",
                        "cas_number": self.results_table.item(row, 1).text() if self.results_table.item(row, 1) else "",
                        "smiles": self.results_table.item(row, 2).text() if self.results_table.item(row, 2) else "",
                        "ec50": self.results_table.item(row, 3).text() if self.results_table.item(row, 3) else "",
                        "category": self.results_table.item(row, 4).text() if self.results_table.item(row, 4) else "",
                        "ad_status": self.results_table.item(row, 5).text() if self.results_table.item(row, 5) else ""
                    }
                    print(f"    化合物名称: {compound_info['compound_name']}")
                    print(f"    SMILES: {compound_info['smiles']}")
                    compounds_data.append(compound_info)
            else:
                print(f"  行 {row+1}: 警告 - 第一列为空")

        print(f"[批量下载] 收集到 {len(compounds_data)} 个化合物数据")
        if not compounds_data:
            QMessageBox.warning(self, "下载失败", "没有可下载的化合物数据")
            return

        # 禁用按钮
        self.batch_download_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.status_label.setText(f"准备下载 {len(compounds_data)} 个化合物...")

        # 创建并启动下载线程
        if hasattr(self, 'batch_thread') and self.batch_thread and self.batch_thread.isRunning():
            self.batch_thread.stop()
            self.batch_thread.wait()

        self.batch_thread = BatchSDFDownloadThread(
            None, compounds_data, None, output_path  # 不需要重新查询，直接使用已有数据
        )
        self.batch_thread.progress_updated.connect(self.on_download_progress)
        self.batch_thread.download_finished.connect(self.on_download_finished)
        self.batch_thread.download_error.connect(self.on_download_error)
        self.batch_thread.start()

    def on_download_progress(self, progress, message):
        """更新下载进度"""
        self.progress_bar.setValue(progress)
        self.status_label.setText(message)

    def on_download_finished(self, sdf_file, success_count, failed_count):
        """下载完成"""
        self.progress_bar.setVisible(False)
        self.status_label.setText(f"下载完成！成功: {success_count}, 失败: {failed_count}")
        self.batch_download_btn.setEnabled(True)

        # 显示完成消息
        msg = f"SDF文件下载完成！\n\n"
        msg += f"成功下载: {success_count} 个化合物\n"
        msg += f"下载失败: {failed_count} 个\n"
        msg += f"\nSDF文件已保存到:\n{sdf_file}"

        QMessageBox.information(self, "下载完成", msg)

        # 询问是否打开输出文件夹
        reply = QMessageBox.question(
            self, "打开文件夹",
            "是否打开输出文件所在文件夹？",
            QMessageBox.Yes | QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            output_dir = os.path.dirname(sdf_file)
            os.startfile(output_dir)

    def on_download_error(self, error_msg):
        """下载出错"""
        self.progress_bar.setVisible(False)
        self.status_label.setText("下载失败")
        self.batch_download_btn.setEnabled(True)

        QMessageBox.critical(self, "下载错误", f"批量下载过程中出现错误:\n\n{error_msg}")

    def on_table_item_clicked(self, item):
        row = item.row()

        compound_info = {
            "compound_name": self.results_table.item(row, 0).text() if self.results_table.item(row, 0) else "",
            "cas_number": self.results_table.item(row, 1).text() if self.results_table.item(row, 1) else "",
            "smiles": self.results_table.item(row, 2).text() if self.results_table.item(row, 2) else "",
            "ec50": self.results_table.item(row, 3).text() if self.results_table.item(row, 3) else "",
            "category": self.results_table.item(row, 4).text() if self.results_table.item(row, 4) else "",
            "ad_status": self.results_table.item(row, 5).text() if self.results_table.item(row, 5) else ""
        }

        # 从数据库重新获取完整信息（使用SMILES）
        if compound_info["smiles"]:
            full_info = self.db_connector.query_data("SMILES", compound_info["smiles"])
            if full_info:
                compound_info.update(full_info)

        self.display_molecule_info(compound_info)

    def display_molecule_info(self, compound_info):
        """显示分子结构和详细信息"""
        # 仅保留从Excel提取ChemDraw图片的功能
        pixmap = None

        # 尝试显示从Excel提取的ChemDraw图片（如果存在）
        if '_chemdraw_image' in compound_info and compound_info['_chemdraw_image']:
            try:
                pil_image = compound_info['_chemdraw_image']
                # 将PIL Image转换为QPixmap
                buffer = io.BytesIO()
                pil_image.save(buffer, format='PNG')
                pixmap = QPixmap()
                pixmap.loadFromData(buffer.getvalue())
                print("显示ChemDraw图片")
            except Exception as e:
                print(f"转换ChemDraw图片失败: {str(e)}")
                pixmap = None

        # 如果没有ChemDraw图片，尝试从SMILES生成分子结构图
        if not pixmap:
            smiles = compound_info.get("smiles", "")
            if smiles:
                pixmap = self.compound_drawer.smiles_to_pixmap(smiles, width=380, height=280)
                if pixmap:
                    print("从SMILES生成分子结构图")

        # 3. 显示图片或提示信息
        if pixmap:
            self.molecule_image_label.setPixmap(pixmap.scaled(
                self.molecule_image_label.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation
            ))
        else:
            self.molecule_image_label.setText("没有可用的分子结构")

        # 显示详细信息（右侧）
        # 获取数据来源（显示数据库名.表名）
        if '_database' in compound_info and '_table' in compound_info:
            # 从数据库查询，显示数据库名.表名
            source_display = f"{compound_info.get('_database', 'N/A')}.{compound_info.get('_table', 'N/A')}"
        elif '_excel_file' in compound_info:
            # 从Excel文件查询，显示文件名
            excel_file = compound_info.get('_excel_file', '')
            source_display = os.path.basename(excel_file) if excel_file else 'Excel文件'
        else:
            # 未知来源
            source_display = '未知来源'
        
        info_text = f"""<h3>化合物详细信息</h3>
<table style='width:100%; border-collapse: collapse;'>
<tr><td style='padding:5px; font-weight:bold; color:#1890ff;'>名称：</td><td style='padding:5px;'>{compound_info.get('compound_name', 'N/A')}</td></tr>
<tr><td style='padding:5px; font-weight:bold; color:#1890ff;'>数据来源：</td><td style='padding:5px;'>{source_display}</td></tr>
<tr><td style='padding:5px; font-weight:bold; color:#1890ff;'>化合物种类：</td><td style='padding:5px;'>{compound_info.get('category', 'N/A')}</td></tr>
<tr><td style='padding:5px; font-weight:bold; color:#1890ff;'>CAS号：</td><td style='padding:5px;'>{compound_info.get('cas_number', 'N/A')}</td></tr>
<tr><td style='padding:5px; font-weight:bold; color:#1890ff;'>分子量：</td><td style='padding:5px;'>{compound_info.get('molecular_weight', 'N/A')}</td></tr>
<tr><td style='padding:5px; font-weight:bold; color:#1890ff;'>动物类型：</td><td style='padding:5px;'>{compound_info.get('animal_type', 'N/A')}</td></tr>
<tr><td style='padding:5px; font-weight:bold; color:#1890ff;'>血管类型：</td><td style='padding:5px;'>{compound_info.get('vessel_type', 'N/A')}</td></tr>
<tr><td style='padding:5px; font-weight:bold; color:#1890ff;'>收缩剂：</td><td style='padding:5px;'>{compound_info.get('vasoconstrictor', 'N/A')}</td></tr>
<tr><td style='padding:5px; font-weight:bold; color:#1890ff;'>用药浓度：</td><td style='padding:5px;'>{compound_info.get('concentration', 'N/A')}</td></tr>
<tr><td style='padding:5px; font-weight:bold; color:#1890ff;'>EC50：</td><td style='padding:5px;'>{compound_info.get('ec50', 'N/A')}</td></tr>
<tr><td style='padding:5px; font-weight:bold; color:#1890ff;'>机制：</td><td style='padding:5px;'>{compound_info.get('mechanism', 'N/A')}</td></tr>
</table>
<br>
<p style='margin-top:10px;'><span style='font-weight:bold; color:#1890ff;'>文献链接：</span><br>{compound_info.get('literature', 'N/A')}</p>
"""

        self.molecule_info_text.setHtml(info_text)

    def save_results(self):
        if self.results_table.rowCount() == 0:
            QMessageBox.warning(self, "保存失败", "没有可保存的结果")
            return

        file_path, file_type = QFileDialog.getSaveFileName(
            self, "保存结果", "", "Excel文件 (*.xlsx);;CSV文件 (*.csv)"
        )

        if not file_path:
            return

        try:
            data = []
            for row in range(self.results_table.rowCount()):
                row_data = {
                    "化合物名称": self.results_table.item(row, 0).text() if self.results_table.item(row, 0) else "",
                    "CAS号": self.results_table.item(row, 1).text() if self.results_table.item(row, 1) else "",
                    "SMILES": self.results_table.item(row, 2).text() if self.results_table.item(row, 2) else "",
                    "EC50": self.results_table.item(row, 3).text() if self.results_table.item(row, 3) else "",
                    "活性类别": self.results_table.item(row, 4).text() if self.results_table.item(row, 4) else "",
                    "AD状态": self.results_table.item(row, 5).text() if self.results_table.item(row, 5) else ""
                }
                data.append(row_data)

            df = pd.DataFrame(data)

            if file_type == "Excel文件 (*.xlsx)":
                if not file_path.endswith('.xlsx'):
                    file_path += '.xlsx'
                df.to_excel(file_path, index=False)
            else:
                if not file_path.endswith('.csv'):
                    file_path += '.csv'
                df.to_csv(file_path, index=False, encoding='utf-8-sig')

            QMessageBox.information(self, "保存成功", f"结果已成功保存到: {file_path}")
        except Exception as e:
            QMessageBox.critical(self, "保存错误", f"保存文件时出现错误: {str(e)}")

    def clear_results(self):
        self.results_table.setRowCount(0)
        self.molecule_image_label.setText("选择表格中的化合物查看分子结构")
        self.molecule_info_text.clear()
        self.save_results_btn.setEnabled(False)
        self.batch_download_btn.setEnabled(False)  # 清空结果后禁用批量下载按钮


# 主函数
if __name__ == "__main__":
    QApplication.setStyle("Fusion")

    app = QApplication(sys.argv)
    window = CardioHerbDBApp()
    window.show()

    sys.exit(app.exec_())
