$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $filePath = "D:\PDFtoEXCEL\cost_september_result.xlsx"
    $wb = $excel.Workbooks.Open($filePath)
    $excel.CalculateFull()

    Write-Host "Worksheet Count: $($wb.Sheets.Count)"

    # 全員工彙總 (Sheet 6)
    $summarySheet = $wb.Sheets.Item(6)
    Write-Host "Summary Sheet Name: $($summarySheet.Name)"
    Write-Host "Aug (Row 12) -> Health Co (E12): $($summarySheet.Range('E12').Value2), Labor Co (F12): $($summarySheet.Range('F12').Value2), Health Self (H12): $($summarySheet.Range('H12').Value2)"
    Write-Host "Sep (Row 13) -> Health Co (E13): $($summarySheet.Range('E13').Value2), Labor Co (F13): $($summarySheet.Range('F13').Value2), Health Self (H13): $($summarySheet.Range('H13').Value2)"

    # 林小美 (Sheet 8, E02)
    $linSheet = $wb.Sheets.Item(8)
    Write-Host "Lin Sheet Name: $($linSheet.Name)"
    Write-Host "Lin Aug Health Self (H12): $($linSheet.Range('H12').Value2)"
    Write-Host "Lin Sep Health Self (H13): $($linSheet.Range('H13').Value2)"

    # 郭台銘 (Sheet 14, E08)
    $guoSheet = $wb.Sheets.Item(14)
    Write-Host "Guo Sheet Name: $($guoSheet.Name)"
    Write-Host "Guo Aug Health Self (H12): $($guoSheet.Range('H12').Value2)"
    Write-Host "Guo Sep Health Self (H13): $($guoSheet.Range('H13').Value2)"

    $wb.Close($false)
    Write-Host "CALCULATION CHECK COMPLETE!"
} catch {
    Write-Host "ERROR: $($_.Exception.Message)"
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
