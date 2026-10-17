$e = New-Object -ComObject Excel.Application
$e.Visible = $false
$e.DisplayAlerts = $false
try {
  $p = (Resolve-Path "cost.xlsx").Path
  $w = $e.Workbooks.Open($p)
  Write-Host "OPEN COST.XLSX SUCCESS!"
  Write-Host "Total Worksheets: $($w.Worksheets.Count)"
  for ($i = 1; $i -le $w.Worksheets.Count; $i++) {
    $ws = $w.Worksheets.Item($i)
    Write-Host "Sheet $i : $($ws.Name)"
  }
  $w.Close($false)
} catch {
  Write-Host "FAILED: $($_.Exception.Message)"
}
$e.Quit()
