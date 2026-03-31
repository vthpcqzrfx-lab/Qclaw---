$output = @()

$files = @(
    "C:\Users\zhaoyang26\Desktop\运营岗-实习转正答辩PPT.xlsx",
    "C:\Users\zhaoyang26\Desktop\关键岗位.xlsx",
    "C:\Users\zhaoyang26\Desktop\信息流投放-头条&抖音-2版&3版.xlsx",
    "C:\Users\zhaoyang26\Desktop\city_level_5_251122.xlsx",
    "C:\Users\zhaoyang26\Desktop\新建 DOCX 文档.docx"
)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

foreach ($f in $files) {
    $output += "=" * 60
    $output += "FILE: " + (Split-Path $f -Leaf)
    $output += "=" * 60

    try {
        $wb = $excel.Workbooks.Open($f)
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
