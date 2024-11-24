# Script-to-convert-XLSX-to-CSV-using-Powershell
This repo is responsible to Automate Excel to CSV Conversion Using PowerShell
This repository provides PowerShell scripts to automate the conversion of Excel (.xlsx) files to CSV files. The scripts are designed for users who need to process multiple Excel files efficiently without relying on Microsoft Office or external tools.

Key Features:<br />
Converts all Excel files in a specified directory to CSV format.<br />
Supports single-sheet and multi-sheet Excel files.<br />
Outputs CSV files with custom delimiters and UTF-8 encoding.<br />
Includes automatic module management to handle Excel import/export operations.<br />
Lightweight and efficient, suitable for integration into larger ETL pipelines or scheduled tasks.<br />

Some Code Explanation <br />
This script handles Excel files with multiple sheets, generating separate CSV files for each sheet.<br />

Install ImportExcel Module:
Checks if the ImportExcel module is installed; if not, it installs it for the current user.
ImportExcel allows Excel file operations without Office installed.
