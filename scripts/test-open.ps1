$e = New-Object -ComObject Excel.Application
$e.Visible = $false
$e.DisplayAlerts = $false
try {
  $w = $e.Workbooks.Open("D:\PDFtoEXCEL\scripts\test_all_fixed.xlsx")
  Write-Host "OPEN SUCCESS!"
  for ($i = 1; $i -le $w.Worksheets.Count; $i++) {
    $ws = $w.Worksheets.Item($i)
    Write-Host "Sheet $i : $($ws.Name)"
  }
  # 讀取員工01
  $wsEmp1 = $w.Worksheets.Item(6)
  Write-Host "員工01 B2: $($wsEmp1.Range('B2').Text)"
  Write-Host "員工01 E12 (健保公司): $($wsEmp1.Range('E12').Text)"
  Write-Host "員工01 F12 (勞保公司): $($wsEmp1.Range('F12').Text)"
  Write-Host "員工01 G12 (勞退提繳): $($wsEmp1.Range('G12').Text)"
  Write-Host "員工01 I12 (勞保自付): $($wsEmp1.Range('I12').Text)"

  # 讀取員工05 (陳萱蘋)
  $wsEmp5 = $w.Worksheets.Item(10)
  Write-Host "員工05 B2: $($wsEmp5.Range('B2').Text)"
  Write-Host "員工05 E12 (健保公司): $($wsEmp5.Range('E12').Text)"
  Write-Host "員工05 H12 (健保自付): $($wsEmp5.Range('H12').Text)"

  $w.Close($false)
} catch {
  Write-Host "FAILED: $($_.Exception.Message)"
}
$e.Quit()
