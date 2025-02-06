# ETL Script for Password-Protected XLSX

This script processes a **password-protected `.xlsx` file**, extracts **ICD-10 codes**, checks for matches in a separate **CSV file**, and writes the filtered data to a new **CSV file**.

## Description

A healthcare company was receiving **Excel files** in an **unsuitable format**.  
- **Problem:** All **patient diagnosis codes (ICD-10)** were stored in a **single cell**, making filtering difficult in Excel.  
- **Solution:** This script **extracts and separates** the ICD-10 codes, **matches** them against a client-specified **list of codes**, and **outputs** only the relevant patient data to a new CSV file.  

---

## Getting Started

### Dependencies

This script is written in **Python** and requires the following libraries:  

- `pandas` - Used for reading, processing, and manipulating `.xlsx` files  
- `msoffcrypto-tool` - Used to decrypt password-protected Excel files  
- `csv` - Used to read and write CSV files  
- `getpass` - Used to securely capture password input from the user  
- `io.BytesIO` - Used to handle in-memory decrypted file streams  
- `datetime` - Used to generate timestamps for naming output files  

Each dependency is listed in the `requirements.txt` file. To install them, follow the installation steps below.

---

### Installing

A **client-provided `.xlsx` file** must be **saved in the root project directory** before running the script.  
- The **table schema** in the `.xlsx` file **must match the expected format** used by this program.  

---

### Executing the Program

1. Install dependencies  
```
pip install -r requirements.txt
```
2. Import the `.xlsx` file into the project directory  
3. Run the script  
```
python match_codes.py
```
4. Enter the filename (including `.xlsx` extension)  
5. Enter the document password when prompted  
6. Retrieve the output CSV file from the project directory  

---

## Authors

**Mike Bibeau**  
[GitHub Profile](https://github.com/MR-8264)  
