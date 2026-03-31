$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$targetFiles = @(
    @{Size=23856; Name="运营岗-实习转正答辩PPT.xlsx"},
    @{Size=7542760; Name="关键岗位.xlsx"},
    @{Size=28979474; Name="信息流投放-头条&抖音-2版&3版.xlsx"},
    @{Size=32308; Name="city_level_5_251122.xlsx"},
    @{Size=17940; Name="新建 DOCX 文档.docx"}
)

$output = @()

foreach ($target in $targetFiles) {
    $file = Get-ChildItem -Path "C:\Users\zhaoyang26\Desktop" -Filter "*$($target.Name)*" | Where-Object { $_.Length -eq $target.Size } | Select-Object -First 1

    if ($file) {
        $output += "=" * 60
        $output += "FILE: " + $file.Name + " (Size: " + $file.Length + ")"
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
    } else {
        $output += "FILE NOT FOUND: " + $target.Name + " (size: " + $target.Size + ")"
    }
    $output += ""
}

$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

$output | Out-File -FilePath "C:\Users\zhaoyang26\.qclaw\workspace\excel_summary.txt" -Encoding UTF8
Write-Host "Done"
