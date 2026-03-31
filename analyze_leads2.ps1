$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false

# Try different lead data files
$files = @(
    "C:\Users\zhaoyang26\Desktop\线索明细数据 (13).xlsx",
    "C:\Users\zhaoyang26\Desktop\线索明细数据 0205使用.xlsx",
    "C:\Users\zhaoyang26\Desktop\线索明细数据to新雨.xlsx",
    "C:\Users\zhaoyang26\Desktop\线索明细数据 (0211期).xlsx"
)

$targetFile = $null
foreach($path in $files) {
    if(Test-Path $path) {
        $targetFile = $path
        Write-Host "找到文件: $path"
        break
    }
}

if(-not $targetFile) {
    Write-Host "未找到任何线索明细文件"
    $xl.Quit()
    exit
}

try {
    $wb = $xl.Workbooks.Open($targetFile)
    Start-Sleep -Seconds 2

    Write-Host "=== 文件概览 ==="
    Write-Host "工作表数量: $($wb.Sheets.Count)"

    if($wb.Sheets.Count -gt 0) {
        $ws = $wb.Sheets.Item(1)
        $usedRange = $ws.UsedRange
        $rowCount = $usedRange.Rows.Count
        $colCount = $usedRange.Columns.Count

        Write-Host "总行数: $rowCount"
        Write-Host "总列数: $colCount"

        Write-Host "`n=== 表头 (前20列) ==="
        for($col=1; $col -le [Math]::Min(20, $colCount); $col++) {
            $val = $ws.Cells.Item(1,$col).Text
            Write-Host "  $col : $val"
        }

        Write-Host "`n=== 前5行数据样本 (前10列) ==="
        for($row=2; $row -le [Math]::Min(6, $rowCount); $row++) {
            $r = @()
            for($col=1; $col -le [Math]::Min(10, $colCount); $col++) {
                $val = $ws.Cells.Item($row,$col).Text
                $r += if($val) { $val } else { "-" }
            }
            Write-Host "Row$($row): $($r -join ' | ')"
        }
    }

    $wb.Close($false)
} catch {
    Write-Host "Error: $_"
    if($wb) { $wb.Close($false) }
}

$xl.Quit()
