$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Open('D:\PDFtoEXCEL\cost_test_renamed.xlsx')
    Write-Host "Total sheets: $($wb.Sheets.Count)"
    for ($i = 1; $i -le $wb.Sheets.Count; $i++) {
        $s = $wb.Sheets.Item($i)
        Write-Host "Sheet $i : $($s.Name)"
    }
    $wb.Close($false)
    Write-Host "SUCCESS: Opened without any issues!"
} catch {
    Write-Host "ERROR: $($_.Exception.Message)"
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
