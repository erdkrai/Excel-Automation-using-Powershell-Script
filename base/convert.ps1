do{
$excel = New-Object -ComObject excel.application
$ext = ".xls"
$random = Get-Random
$get_date = Get-Date
$date_string = ([DateTime]($get_date)).addMinutes(15).ToString('mm/dd/yyyy')
$dest = "C:\Users\erdkr\OneDrive\Desktop\powershell_excel_conversion\$random-file_$date_string$ext"
$excel.visible = $True

$objWorkbook = $excel.workbooks.Open("C:\Users\erdkr\OneDrive\Desktop\powershell_excel_conversion\main_excel\excel1.xls")

start-sleep -Seconds 10

$objWorkbook.Close($True)
$excel.Quit()

Copy-Item "C:\Users\erdkr\OneDrive\Desktop\powershell_excel_conversion\main_excel\excel1.xls" -Destination $dest

Write-Host "SAVED ==> Converting your Excel File from XLS to CSV" -ForegroundColor Green

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