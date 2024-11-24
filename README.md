# Script-to-convert-XLSX-to-CSV-using-Powershell
This repo is responsible to Automate Excel to CSV Conversion Using PowerShell
This repository provides PowerShell scripts to automate the conversion of Excel (.xlsx) files to CSV files. The scripts are designed for users who need to process multiple Excel files efficiently without relying on Microsoft Office or external tools.

Key Features:<br />
Converts all Excel files in a specified directory to CSV format.<br />
Supports single-sheet and multi-sheet Excel files.<br />
Outputs CSV files with custom delimiters and UTF-8 encoding.<br />
Includes automatic module management to handle Excel import/export operations.<br />
Lightweight and efficient, suitable for integration into larger ETL pipelines or scheduled tasks.<br />

When to Use Each Script?<br />
First Script:<br />

Use when processing Excel files with multiple sheets.<br />
Outputs a separate CSV for each sheet.<br />

Second Script:<br />

Use when processing Excel files with a single sheet or only one output file per Excel file is needed.<br />


Some Code Explanation <br />
This script handles Excel files with multiple sheets, generating separate CSV files for each sheet.<br />

Install ImportExcel Module:<br />
Checks if the ImportExcel module is installed; if not, it installs it for the current user.<br />
ImportExcel allows Excel file operations without Office installed.<br />

Define Input and Output Directories:<br />
$indir: Directory containing Excel files.<br />
$outdir: Directory where CSV files will be saved.<br />

Process Each File: Using foreach looping to Loops through each Excel file in the input directory. <br /> 

Handle Multi-Sheet Excel Files: Opens the Excel file and extracts sheet names using Open-ExcelPackage and Iterates over each sheet, converts it to CSV, and saves it with the sheet name appended. <br />

Cleanup:<br />
Removes the ImportExcel module after the process is complete.<br />




# Excel to CSV Conversion Using PowerShell  

This repository contains PowerShell scripts to automate the conversion of Excel files (.xlsx) to CSV files. These scripts are lightweight, efficient, and designed for environments where Microsoft Office is not installed. Perfect for integration into ETL pipelines or scheduled tasks, the scripts handle single-sheet and multi-sheet Excel files seamlessly.  

---

## Features  
- **Supports Multi-Sheet Excel Files:** Converts each sheet into a separate CSV file.  
- **Batch Processing:** Processes all `.xlsx` files in a specified directory.  
- **Customizable Output:** Allows setting custom delimiters and UTF-8 encoding for CSV files.  
- **No Office Dependency:** Uses the `ImportExcel` module to handle Excel files without requiring Microsoft Office.  
- **Modular Cleanup:** Automatically removes the `ImportExcel` module after execution to maintain a clean environment.  

---

## Requirements  
- PowerShell 5.0 or later.  
- The `ImportExcel` PowerShell module (installed automatically by the script if not available).  

---

## Usage  

### 1. Script for Multi-Sheet Excel Files  
This script processes each Excel file in the input directory and generates a separate CSV file for each sheet.  

#### How It Works:  
1. Installs the `ImportExcel` module if not already installed.  
2. Loops through all `.xlsx` files in the input directory.  
3. Extracts all sheet names from each Excel file.  
4. Converts each sheet into a separate CSV file with the sheet name appended to the filename.  

#### Example:  
Input file: `Sample.xlsx` (with sheets `Sheet1` and `Sheet2`).  
Output:  
- `Sample_Sheet1.csv`  
- `Sample_Sheet2.csv`  

#### Code:  
Refer to [`script1.ps1`](path-to-script1) in this repository.  

---

### 2. Script for Single-Sheet Excel Files  
This script processes each Excel file in the input directory, generating one CSV file per Excel file.  

#### How It Works:  
1. Installs the `ImportExcel` module if not already installed.  
2. Loops through all `.xlsx` files in the input directory.  
3. Exports the content of each file (assuming a single sheet) to a CSV file with the same name as the Excel file.  

#### Example:  
Input file: `Sample.xlsx` (single sheet).  
Output:  
- `Sample.csv`  

#### Code:  
Refer to [`script2.ps1`](path-to-script2) in this repository.  

---

## How to Schedule the Scripts  

You can automate the execution of these scripts using **SQL Server Agent** or any other task scheduling tool:  

1. Open **SQL Server Management Studio**.  
2. Go to **SQL Server Agent** > **New Job**.  
3. Set up the job name and description.  
4. Add a new step:  
   - Choose **PowerShell Script** as the type.  
   - Paste the script content into the command field.  
5. Configure the schedule for the job (e.g., daily, hourly).  

---

## Repository Structure  
```plaintext
.
├── script1.ps1  # Script for multi-sheet Excel files
├── script2.ps1  # Script for single-sheet Excel files
├── README.md    # Documentation







