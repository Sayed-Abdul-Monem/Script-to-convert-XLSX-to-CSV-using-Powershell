# Script-to-convert-XLSX-to-CSV-using-Powershell
This repo is responsible to Automate Excel to CSV Conversion Using PowerShell
This repository provides PowerShell scripts to automate the conversion of Excel (.xlsx) files to CSV files. The scripts are designed for users who need to process multiple Excel files efficiently without relying on Microsoft Office or external tools.

## Key Features:<br />
Converts all Excel files in a specified directory to CSV format.<br />
Supports single-sheet and multi-sheet Excel files.<br />
Outputs CSV files with custom delimiters and UTF-8 encoding.<br />
Includes automatic module management to handle Excel import/export operations.<br />
Lightweight and efficient, suitable for integration into larger ETL pipelines or scheduled tasks.<br />

When to Use Each Script?<br />
### First Script:<br />

Use when processing Excel files with multiple sheets.<br />
Outputs a separate CSV for each sheet.<br />

### Second Script:<br />

Use when processing Excel files with a single sheet or only one output file per Excel file is needed.<br />


### Some Code Explanation <br />
This script handles Excel files with multiple sheets, generating separate CSV files for each sheet.<br />

#### Install ImportExcel Module:<br />
Checks if the ImportExcel module is installed; if not, it installs it for the current user.<br />
ImportExcel allows Excel file operations without Office installed.<br />

#### Define Input and Output Directories:<br />
$indir: Directory containing Excel files.<br />
$outdir: Directory where CSV files will be saved.<br />

#### Process Each File: Using foreach looping to Loops through each Excel file in the input directory. <br /> 

#### Handle Multi-Sheet Excel Files: Opens the Excel file and extracts sheet names using Open-ExcelPackage and Iterates over each sheet, converts it to CSV, and saves it with the sheet name appended. <br />

#### Cleanup:<br />
Removes the ImportExcel module after the process is complete.<br />











