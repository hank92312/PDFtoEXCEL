$e = New-Object -ComObject Excel.Application
$e.Visible = $false
$e.DisplayAlerts = $false
try {
  $w = $e.Workbooks.Open("D:\PDFtoEXCEL\scripts\test_batch_8_and_9.xlsx")
  Write-Host "OPEN MULTI-MONTH BATCH SUCCESS!"

  # 員工01 在第 6 個工作表
  $ws = $w.Worksheets.Item(6)
  Write-Host "Worksheet Name: $($ws.Name)"
  Write-Host "B2 (Name): $($ws.Range('B2').Text)"

  # 檢查各列的月份日期與保費金額
  for ($r = 4; $r -le 16; $r++) {
    $dateCol = $ws.Range("A$r").Text
    $healthCo = $ws.Range("E$r").Text
    $laborCo = $ws.Range("F$r").Text
    $pensionCo = $ws.Range("G$r").Text
    $healthSelf = $ws.Range("H$r").Text
    $laborSelf = $ws.Range("I$r").Text
    if ($dateCol) {
      Write-Host "Row $r | Month: $dateCol | HealthCo: $healthCo | LaborCo: $laborCo | PensionCo: $pensionCo | HealthSelf: $healthSelf | LaborSelf: $laborSelf"
    }
  }

  $w.Close($false)
} catch {
  Write-Host "FAILED: $($_.Exception.Message)"
}
$e.Quit()
