do{
$excel = New-Object -ComObject excel.application
$ext = ".xls"
$random = Get-Random
$get_date = Get-Date -UFormat "%m-%d-%Y"
#$date_string = ([DateTime]($get_date)).addMinutes(15).ToString('mm/dd/yyyy')
$dest = "C:\Users\erdkr\OneDrive\Desktop\Excel-Automation-using-Powershell-Script\$random-file_$get_date$ext"
$excel.visible = $True

$objWorkbook = $excel.workbooks.Open("C:\Users\erdkr\OneDrive\Desktop\Excel-Automation-using-Powershell-Script\main_excel\excel1.xls")

start-sleep -Seconds 10

$objWorkbook.Close($True)
$excel.Quit()

$SourceWorkbook = 'C:\Users\erdkr\OneDrive\Desktop\Excel-Automation-using-Powershell-Script\main_excel\excel1.xls' 
$TargetWorkbook = "C:\Users\erdkr\OneDrive\Desktop\Excel-Automation-using-Powershell-Script\main_excel\excel_file.xls"

#Creating a blank excel file in which the data is to be pasted
$excel.visible = $False
$outputpath = 'C:\Users\erdkr\OneDrive\Desktop\Excel-Automation-using-Powershell-Script\main_excel\excel_file.xls'
$workbook = $excel.Workbooks.Add()
$workbook.SaveAs($outputpath) 
$excel.Quit()

#Copying Data of cells A1:G1 from 'excel1.xls' to 'excel_file.xls'
$excel.displayAlerts = $false # don't prompt the user
$excel.visible = $false
$excel.displayAlerts=$false
#$xlPasteValues = $wb1.Range('A1:G1').EntireColumn
$wb1=$excel.workbooks.open($SourceWorkbook, $null, $true)
$wb2=$excel.workbooks.open($TargetWorkbook)

$targetRange=$wb2.Worksheets.Item(1).Range('A1:G1').EntireColumn
$wb1.Worksheets.Item(1).Range('A1:G1').EntireColumn.copy()
$wb2.Worksheets.Item(1).Activate()
$targetRange.PasteSpecial(-4163)

$wb1.close($True)
$wb2.close($True)
$excel.quit()

#Copying the file from main_excel folder to parent folder and deleting the file from main_excel folder (we can also use Move-Item in this case)
Copy-Item "C:\Users\erdkr\OneDrive\Desktop\Excel-Automation-using-Powershell-Script\main_excel\excel_file.xls" -Destination $dest
Remove-Item 'C:\Users\erdkr\OneDrive\Desktop\Excel-Automation-using-Powershell-Script\main_excel\excel_file.xls'

Write-Host "SAVED ==> Converting your Excel File from XLS to CSV" -ForegroundColor Green

#Converting Excel File
Get-ChildItem -Path C:\Users\erdkr\OneDrive\Desktop\Excel-Automation-using-Powershell-Script -Filter "*.xls" | ForEach-Object{
    $Workbook = $excel.Workbooks.Open($_.Fullname) 
    $newName = ($_.Fullname).Replace($_.Extension,".csv")
    $Workbook.SaveAs($newName)
    $Workbook.Close($true)
    $excel.Quit()
    Stop-Process -n excel	
}


#Moving Files to their Specified Folders
Move-Item -Path C:\Users\erdkr\OneDrive\Desktop\Excel-Automation-using-Powershell-Script\*.xls -Destination C:\Users\erdkr\OneDrive\Desktop\Excel-Automation-using-Powershell-Script\xls_files
Move-Item -Path C:\Users\erdkr\OneDrive\Desktop\Excel-Automation-using-Powershell-Script\*.csv -Destination C:\Users\erdkr\OneDrive\Desktop\Excel-Automation-using-Powershell-Script\csv_files

Write-Host "Converted and now waiting for next 10 Seconds" -ForegroundColor Green

start-sleep -Seconds 10

}until ($infinity)