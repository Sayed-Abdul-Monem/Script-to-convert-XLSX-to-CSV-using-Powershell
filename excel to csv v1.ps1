# Install NuGet PackageProvider if not installed (without specifying version)
if (-not (Get-PackageProvider -Name NuGet)) {
    Install-PackageProvider -Name NuGet -Force
}

# Install ImportExcel Module if not installed 
if (-not (Get-Module -ListAvailable -Name ImportExcel)) 
    { Install-Module -Name ImportExcel -Force -Scope CurrentUser}

# Create variable for directory where you save your excel files 
$indir = "C:\Users\Sayed Abdul-Monem\Desktop\Excel to csv powershell\Excel files\"  

# Create variable for directory where you want to save the csv output files
$outdir = "C:\Users\Sayed Abdul-Monem\Desktop\Excel to csv powershell\Csv Output\"   
$infiles = Get-ChildItem $indir -Filter "*.xlsx"  # Only get .xlsx files

foreach($infile in $infiles)
{
    $file = $infile.FullName  # Full path of the current file
    $filenameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($infile.Name)  # Get file name without extension
    
    $outfile = "{0}{1}.csv" -f $outdir, $filenameWithoutExtension  # Set output file path
    
    # Import the Excel file (since there's only one sheet, no need to specify the worksheet name) add -NoHeader if no header found in the file
    Import-Excel -Path $file  |  
        Export-Csv -Path $outfile -NoTypeInformation -Delimiter ";" -Encoding UTF8
}

# Remove the imported module if you need 
Remove-Module -Name ImportExcel
Exit
