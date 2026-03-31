$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$output = @()

# Get first 5 xlsx files from Desktop sorted by size
$files = Get-ChildItem -Path "C:\Users\zhaoyang26\Desktop" -Filter "*.xlsx" | Sort-Object Length -Descending | Select-Object -First 5

foreach ($file in $files) {
    $output += "=" * 60
    $output += "FILE: " + $file.Name
    $output += "Size: " + $file.Length
    $output += "=" * 60

    try {
        $wb = $excel.Workbooks.Open($file.FullName)
        $output += "Sheets: " + ($wb.Sheets | ForEach-Object { $_.Name } -join ", ")

        foreach ($sheet in $wb.Sheets) {
            if ($sheet.Index -gt 3) { break }
            $output += ""
            $output += "--- Sheet: " + $sheet.Name + " ---"

            $usedRange = $sheet.UsedRange
            $rowCount = [Math]::Min($usedRange.Rows.Count, 20)

            for ($r = 1; $r -le $rowCount; $r++) {
                $rowData = @()
                $colCount = [Math]::Min($usedRange.Columns.Count, 20)
                for ($c = 1; $c -le $colCount; $c++) {
                    $cell = $sheet.Cells.Item($r, $c)
                    if ($cell.Value2 -ne $null) {
                        $rowData += $cell.Value2.ToString()
                    } else {
                        $rowData += ""
                    }
                }
                $output += ($rowData -join " | ")
            }
        }
        $wb.Close($false)
    } catch {
        $output += "ERROR: " + $_.Exception.Message
    }
    $output += ""
}

$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

$output | Out-File -FilePath "C:\Users\zhaoyang26\.qclaw\workspace\excel_summary.txt" -Encoding UTF8
Write-Host "Done"
