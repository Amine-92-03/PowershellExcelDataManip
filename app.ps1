$wbname = 'C:\Users\MohamedElAmine.MEFTI\Desktop\Work Space\exploitation DC DEVRED\applications exploitation\PowershellExcelDataManip\inputExcel.xlsx'
$xlsx = New-Object -comobject Excel.Application
$xlsx.DisplayAlerts = $False
$wb = $xlsx.Workbooks.open($wbname)
$sheet = $wb.Sheets.Item(1)
#suppression de la colone A
$sheet.Range("A:A").EntireColumn.Delete()
#Supression des lignes vides
for ($i = 0; $i -lt 6; $i++) {
    $sheet.Cells.Item($i,1).EntireRow.Delete()
}

$xlsx.DisplayAlerts = $False

$wb.Save()
$wn.close()