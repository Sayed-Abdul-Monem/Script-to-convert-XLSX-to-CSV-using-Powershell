# Install ImportExcel Module if not installed 
if (-not (Get-Module -ListAvailable -Name ImportExcel)) 
    { Install-Module -Name ImportExcel -Scope CurrentUser -Force }
 
# Create variable for directory where you save your excel files 
$indir = "C:\Users\Sayed Abdul-Monem\Desktop\Excel to csv powershell\Excel files\"  

# Create variable for directory where you want to save the csv output files
$outdir = "C:\Users\Sayed Abdul-Monem\Desktop\Excel to csv powershell\Csv Output\"  
$infiles = (Get-ChildItem $indir).Name
 
foreach($infile in $infiles)
 
{$file = "$indir$infile"
(Open-ExcelPackage -Path $file).psobject.properties |  
    Where-Object {$_.MemberType -eq 'ScriptProperty'} |  
        Select-Object -ExpandProperty Name |  
            ForEach-Object{  
                $outfile = "{0}{1}.csv" -f $outdir,$_  
                Import-Excel -Path $file -WorksheetName $_ |  
                    Export-Csv $outfile -NoTypeInformation -Delimiter ";" -Encoding UTF8 } }
# Remove the imported module if you need  
Remove-Module -Name ImportExcel
Exit