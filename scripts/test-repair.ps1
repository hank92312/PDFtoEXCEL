$e = New-Object -ComObject Excel.Application
$e.Visible = $false
$e.DisplayAlerts = $false

# CorruptLoad: 2 = xlRepairFile (嘗試修復)
try {
  $w = $e.Workbooks.Open("D:\PDFtoEXCEL\scripts\test_1_pure_resave.xlsx", 0, $true, 5, "", "", $true, 2, $null, $false, $false, 0, $false, $true, 2)
  Write-Host "REPAIR_OPEN_SUCCESS!"
  Write-Host "Worksheets count: $($w.Worksheets.Count)"
  $w.Close($false)
} catch {
  Write-Host "REPAIR_FAILED: $($_.Exception.Message)"
}
$e.Quit()
