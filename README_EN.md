EBC Activation Recommendation

This script is designed to automatically evaluate whether products targeted for EBC (Enhanced Brand Content) activation have complete data, meaning photos and annotations.

This allows you to quickly identify which products are fully prepared and which require additional data before activation.

Features

The current version supports:

✅ Loading the main product database from good.xlsx

Works with a sheet containing at minimum:

OID_zbozi (unique product identifier)

If Fotografie or Anotace columns are missing, the script automatically creates them as empty, preventing crashes.

✅ Processing multiple input files with product lists

For each .xlsx file, the script evaluates products where at least one of the following columns contains the value ano:

EBC požadovaný stav

EBC 2 požadovaný stav

✅ Merging input product data with information from good.xlsx

✅ Evaluating data completeness for each product:

Fotografie + Anotace – product has both

Pouze fotografie – missing annotation

Pouze anotace – missing photo

Chybí obojí – missing both

✅ Automatically ignoring temporary Excel lock files (~$*.xlsx)

✅ Generating an output file doporučení_aktivace_EBC.xlsx with the complete evaluation

✅ Displaying a summary of data completeness directly in the terminal

✅ Graceful error handling:

Open file locked by Excel (PermissionError)

Corrupted file (BadZipFile)

Empty sheets or missing required columns

Required Files

All files should be placed in the same folder as the script (doporuceni_aktivace_EBC.py).

1. good.xlsx

Main database of all products.
It must contain at least the following:

Column	Required	Description
OID_zbozi	YES	Unique product identifier
Fotografie	NO (auto-created if missing)	Photo filename, URL, or any reference to product images
Anotace	NO (auto-created if missing)	Text annotation or product description

Tip:
If Fotografie or Anotace are missing, the script will automatically create empty columns, ensuring the script doesn't crash.

2. Input product list files

Each input file must be in .xlsx format and contain at least:

Column	Required	Description
(first column)	YES	Product identifier (must match OID_zbozi from good.xlsx)
EBC požadovaný stav	YES	Value ano / ne
EBC 2 požadovaný stav	YES	Value ano / ne

Products where either column contains ano are included in the evaluation as candidates for EBC activation.

Output

When the script finishes, it generates the following file:

doporučení_aktivace_EBC.xlsx


Each row contains:

Column	Description
Zdrojový soubor	Name of the input file where the product came from
OID_zbozi	Unique product identifier
Fotografie	Original value from good.xlsx
Anotace	Original value from good.xlsx
Stav dat	Result of the evaluation (see descriptions below)
... other columns from the original input file	
Example Output
Zdrojový soubor	OID_zbozi	Fotografie	Anotace	Stav dat
2025_04_svítidla_RAJL_JIKU.xlsx	1001	photo1.jpg	Desc A	Fotografie + Anotace
2025_04_svítidla_RAJL_JIKU.xlsx	1002	photo2.jpg	(empty)	Pouze fotografie
2025_06_svítidla_RAJL_JIKU.xlsx	1003	(empty)	Desc B	Pouze anotace
2025_06_svítidla_RAJL_JIKU.xlsx	1004	(empty)	(empty)	Chybí obojí
How to Run

Place all required files in one folder:

doporuceni_aktivace_EBC.py (the script)

good.xlsx (product database)

All input .xlsx files with EBC product lists

Close all Excel files before running the script.

Open Command Prompt in the folder:

cd "C:\Práce na datech\Doporučení aktivace EBC"


Run the script:

python doporuceni_aktivace_EBC.py


Once finished, the output file doporučení_aktivace_EBC.xlsx will be created, and a summary will be displayed in the terminal.

Notes

Files starting with ~$ are automatically ignored (Excel temporary lock files).

If a file is open in Excel during processing, the script will skip it and display a warning.

If OID_zbozi is missing in good.xlsx, the script terminates with an error since product matching is impossible without it.

If Fotografie or Anotace are missing, they are automatically added as empty columns.

A summary of the results is displayed directly in the terminal after processing.

Example Terminal Summary
📊 Data completeness summary (all files):
  - Fotografie + Anotace: 157
  - Pouze fotografie: 21
  - Pouze anotace: 12
  - Chybí obojí: 3

✅ DONE! Output saved as:
C:\Práce na datech\Doporučení aktivace EBC\doporučení_aktivace_EBC.xlsx

Requirements

Python 3.9+ (recommended: 3.13)

Required packages:

pandas
openpyxl


Install with:

pip install pandas openpyxl

Summary

This script simplifies the process of determining which products are ready for EBC activation by checking photo and annotation availability, merging the data with a central product database, and generating a comprehensive output Excel file with clear indicators of what is missing.
