# Delete all Files in C:\temp older than 30 day(s)
$Path = "C:\Users\erdkr\OneDrive\Desktop\powershell_excel_conversion\csv_files"
$Path1 = "C:\Users\erdkr\OneDrive\Desktop\powershell_excel_conversion\xls_files"
$Daysback = "-15"
 
$CurrentDate = Get-Date
$DatetoDelete = $CurrentDate.AddDays($Daysback)
Get-ChildItem $Path | Where-Object { $_.LastWriteTime -lt $DatetoDelete } | Remove-Item
Get-ChildItem $Path1 | Where-Object { $_.LastWriteTime -lt $DatetoDelete } | Remove-Item

Clear-RecycleBin -Force

Write-Host "Deleted all the files (xls and csv both) which all are older than 15 days" -ForegroundColor Green