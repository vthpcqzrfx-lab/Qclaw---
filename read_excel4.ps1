$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$output = ""

# Get first 5 xlsx files from Desktop sorted by size
$files = Get-ChildItem -Path "C:\Users\zhaoyang26\Desktop" -Filter "*.xlsx" | Sort-Object Length -Descending | Select-Object -First 5

foreach ($file in $files) {
    $output = $output + "============================================================`n"
    $output = $output + "FILE: " + $file.Name + "`n"
    $output = $output + "Size: " + $file.Length + "`n"
    $output = $output + "============================================================`n"

    try {
        $wb = $excel.Workbooks.Open($file.FullName)
        $sheetNames = ""
        foreach ($s in $wb.Sheets) {
            if ($sheetNames -ne "") { $sheetNames = $sheetNames + ", " }
            $sheetNames = $sheetNames + $s.Name
        }
        $output = $output + "Sheets: " + $sheetNames + "`n"

        foreach ($sheet in $wb.Sheets) {
            if ($sheet.Index -gt 3) { break }
            $output = $output + "`n"
            $output = $output + "--- Sheet: " + $sheet.Name + " ---`n"

            $usedRange = $sheet.UsedRange
            $rowCount = [Math]::Min($usedRange.Rows.Count, 20)

            for ($r = 1; $r -le $rowCount; $r++) {
                $rowData = ""
                $colCount = [Math]::Min($usedRange.Columns.Count, 20)
                for ($c = 1; $c -le $colCount; $c++) {
                    $cell = $sheet.Cells.Item($r, $c)
                    if ($cell.Value2 -ne $null) {
                        if ($rowData -ne "") { $rowData = $rowData + " | " }
                        $rowData = $rowData + $cell.Value2.ToString()
                    }
                }
                $output = $output + $rowData + "`n"
            }
        }
        $wb.Close($false)
    } catch {
        $output = $output + "ERROR: " + $_.Exception.Message + "`n"
    }
    $output = $output + "`n"
}

$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

$output | Out-File -FilePath "C:\Users\zhaoyang26\.qclaw\workspace\excel_summary.txt" -Encoding UTF8
Write-Host "Done"
