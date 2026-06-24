# Cardio Herb DB

This repository provides the compound data and supporting Python programs used in our cardiovascular traditional Chinese medicine database project, titled **治疗心血管中药-靶标-疾病信息数据库**.

We organized chemical constituent information for cardiovascular-related traditional Chinese medicines into herb-level data folders, and prepared structure information such as SMILES, InChI, InChIKey, SDF files, and molecular structure images where available. The repository is intended to make the compound data and database construction workflow accessible as supplementary material for our manuscript.

## Repository Structure

```text
.
├── Compounds/            # Curated herb-compound data and structure files
├── cardio_herb_db.py     # Desktop database query and visualization program
├── cas_to_smiles.py      # CAS-to-SMILES processing utility
├── LICENSE
└── README.md
```

## Data

The `Compounds/` directory contains herb-specific folders. Each folder stores the compound table for one herb and, when available, associated structure images, SDF files, source spreadsheets, and processing scripts.

Most folders follow a structure similar to:

```text
Compounds/<herb-name>化学成分信息表/
├── compounds.csv
├── images/
├── sdf/
└── source or processing files
```

The compound tables were compiled from herb-specific source files, so field names are not fully identical across all herbs. The main types of information included are:

- compound identifier
- Chinese and/or English compound name
- CAS number
- molecular formula and molecular weight
- compound category
- SMILES
- InChI
- InChIKey
- herb source

The repository currently includes compound data folders for:

三七、三棱、丝瓜络、丹参、九里香、乳香、云南红景天、儿茶、冬凌草、凌霄花、北刘寄奴、千金子、半枝莲、卷柏、合欢皮、四季青、土牛膝、地锦草、夏天无、大蓟、大血藤、大黄、天仙藤、姜黄、小叶莲、山楂、山楂叶、山香圆叶、川牛膝、川芎、广枣、延胡索、当归、急性子、朱砂根、桃仁、桃枝、水红花子、没药、泽兰、洪连、滇鸡血藤、牛膝、牡丹皮、独活、王不留行、珠子参、瓜子金、甜瓜子、益母草、瞿麦、矮地茶、穿山龙、红花、续断、肿节风、苏木、茺蔚子、蒲黄、蓍草、虎杖、败酱草、赤芍、赶黄草、连钱草、郁金、金荞麦、金铁锁、银杏叶、阿魏、降香、预知子、鬼箭羽、鸡血藤、龙血竭.

## Database Program

`cardio_herb_db.py` is the desktop program we used for compound record query, display, and structure handling. It is built with PyQt5 and RDKit, and supports:

- querying compounds by SMILES, CAS number, or compound name
- displaying query results in a table
- rendering two-dimensional molecular structures from SMILES
- exporting query results to Excel or CSV
- generating or downloading SDF files from selected records
- connecting to MySQL or SQL Server databases
- extracting ChemDraw/Excel structure images in the Windows desktop workflow

The default database configuration in the script uses MySQL on `localhost:3306` with database name `compounds`. Database connection settings are defined in `CardioHerbDBApp.__init__`.

## CAS-to-SMILES Utility

`cas_to_smiles.py` was used as an auxiliary data-processing script. It extracts CAS numbers from Excel files, identifies CAS columns automatically where possible, and can query PubChem through `pubchempy` to obtain SMILES strings.

## Environment

Main dependencies:

- Python 3.7 or later
- PyQt5
- RDKit
- pandas
- numpy
- pymysql or mysql-connector-python
- pyodbc
- pubchempy
- pywin32
- Pillow

Example installation:

```bash
pip install PyQt5 rdkit-pypi pandas numpy pymysql pyodbc pubchempy pywin32 Pillow
```

The full desktop workflow is intended for Windows because the ChemDraw/Excel extraction module uses Windows COM interfaces.

## Usage

Run the database program after installing the dependencies and configuring the database connection:

```bash
python cardio_herb_db.py
```

View the CAS-to-SMILES utility options:

```bash
python cas_to_smiles.py --help
```

## Version for Manuscript Submission

For manuscript review and supplementary material submission, this repository is linked to the following fixed database upload commit:

```text
752ae0f Add cardio herb compound database
```

## Citation

Citation information will be updated after the manuscript information is finalized.

## License

This repository is released under the MIT License. See `LICENSE` for details.
