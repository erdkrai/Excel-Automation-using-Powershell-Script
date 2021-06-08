do{
$excel = New-Object -ComObject excel.application

#Converting Excel File
Get-ChildItem -Path C:\Users\erdkr\OneDrive\Desktop\powershell_excel_conversion -Filter "*.xls" | ForEach-Object{
    $Workbook = $excel.Workbooks.Open($_.Fullname) 
    $newName = ($_.Fullname).Replace($_.Extension,".csv")
    $Workbook.SaveAs($newName)
    $Workbook.Close($true)
    $excel.Quit()	
}


#Moving Files to their Specified Folders
Move-Item -Path C:\Users\erdkr\OneDrive\Desktop\powershell_excel_conversion\*.xls -Destination C:\Users\erdkr\OneDrive\Desktop\powershell_excel_conversion\xls_files
Move-Item -Path C:\Users\erdkr\OneDrive\Desktop\powershell_excel_conversion\*.csv -Destination C:\Users\erdkr\OneDrive\Desktop\powershell_excel_conversion\csv_files

Write-Host "Converted and now waiting for next 10 Seconds" -ForegroundColor Green

start-sleep -Seconds 10

}until ($infinity)