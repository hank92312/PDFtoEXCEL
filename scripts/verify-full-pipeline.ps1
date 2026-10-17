$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $filePath = "D:\PDFtoEXCEL\cost_september_result.xlsx"
    Write-Host "Opening file with Excel COM..."
    $wb = $excel.Workbooks.Open($filePath)
    
    $cnt = $wb.Sheets.Count
    Write-Host "Total Sheets: $cnt"
    for ($i = 1; $i -le $cnt; $i++) {
        $s = $wb.Sheets.Item($i)
        $n = $s.Name
        Write-Host "Sheet $i : $n"
    }

    # Sheet 6 is 全員工彙總
    $allSheet = $wb.Sheets.Item(6)
    $tB2 = $allSheet.Range("B2").Text
    Write-Host "Sheet 6 (All Employees Summary) B2: $tB2"

    # August row 12
    $augD = $allSheet.Range("D12").Text
    $augE = $allSheet.Range("E12").Text
    $augF = $allSheet.Range("F12").Text
    $augG = $allSheet.Range("G12").Text
    $augH = $allSheet.Range("H12").Text
    Write-Host "Aug Row 12 -> Health Co: $augD, Labor Co: $augE, Pension Co: $augF, Health Personal: $augG, Labor Personal: $augH"

    # September row 13
    $sepD = $allSheet.Range("D13").Text
    $sepE = $allSheet.Range("E13").Text
    $sepF = $allSheet.Range("F13").Text
    $sepG = $allSheet.Range("G13").Text
    $sepH = $allSheet.Range("H13").Text
    Write-Host "Sep Row 13 -> Health Co: $sepD, Labor Co: $sepE, Pension Co: $sepF, Health Personal: $sepG, Labor Personal: $sepH"

    # Sheet 8 is Employee 2 (林小美)
    $linSheet = $wb.Sheets.Item(8)
    $linName = $linSheet.Name
    $linB2 = $linSheet.Range("B2").Text
    $linAugG = $linSheet.Range("G12").Text
    $linSepG = $linSheet.Range("G13").Text
    Write-Host "Sheet 8 ($linName) B2: $linB2, Aug Health: $linAugG, Sep Health: $linSepG"

    # Sheet 14 is Employee 8 (郭台銘)
    $guoSheet = $wb.Sheets.Item(14)
    $guoName = $guoSheet.Name
    $guoB2 = $guoSheet.Range("B2").Text
    $guoAugG = $guoSheet.Range("G12").Text
    $guoSepG = $guoSheet.Range("G13").Text
    Write-Host "Sheet 14 ($guoName) B2: $guoB2, Aug Health: $guoAugG, Sep Health: $guoSepG"

    $wb.Close($false)
    Write-Host "ALL CHECKS COMPLETED SUCCESSFULLY!"
} catch {
    Write-Host "ERROR: $($_.Exception.Message)"
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
