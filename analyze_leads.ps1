$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false

$path = "C:\Users\zhaoyang26\Desktop\线索明细数据.xlsx"
$wb = $xl.Workbooks.Open($path)

Write-Host "=== 文件概览 ==="
Write-Host "工作表数量: $($wb.Sheets.Count)"
Write-Host "工作表名称: "
foreach($sheet in $wb.Sheets) {
    Write-Host "  - $($sheet.Name)"
}

# 读取第一个sheet的数据
$ws = $wb.Sheets.Item(1)
$usedRange = $ws.UsedRange
$rowCount = $usedRange.Rows.Count
$colCount = $usedRange.Columns.Count

Write-Host ""
Write-Host "=== 数据规模 ==="
Write-Host "总行数: $rowCount"
Write-Host "总列数: $colCount"

# 读取表头
Write-Host ""
Write-Host "=== 表头 ==="
$headers = @()
for($col=1; $col -le $colCount; $col++) {
    $val = $ws.Cells.Item(1,$col).Text
    $headers += $val
    if($col -le 20) {
        Write-Host "  Col$col : $val"
    }
}

# 读取前10行数据
Write-Host ""
Write-Host "=== 前10行数据样本 ==="
for($row=2; $row -le [Math]::Min(11, $rowCount); $row++) {
    $r = @()
    for($col=1; $col -le [Math]::Min(15, $colCount); $col++) {
        $val = $ws.Cells.Item($row,$col).Text
        $r += if($val) { $val } else { "-" }
    }
    Write-Host "Row$($row): $($r -join ' | ')"
}

# 简单统计 - 尝试识别关键数值列
Write-Host ""
Write-Host "=== 关键列数据分析 ==="

# 找出数值类型的列并统计
$numericCols = @()
for($col=1; $col -le $colCount; $col++) {
    $sampleSum = 0
    $sampleCount = 0
    for($row=2; $row -le [Math]::Min(100, $rowCount); $row++) {
        $val = $ws.Cells.Item($row,$col).Text
        if($val -match '^[\d,\.\-%]+$' -and $val -notmatch '^[A-Za-z]') {
            $numVal = $val -replace ',',''
            if($numVal -match '^\d+(\.\d+)?$') {
                $sampleSum += [double]$numVal
                $sampleCount++
            }
        }
    }
    if($sampleCount -gt 50) {
        $header = $headers[$col-1]
        $avg = $sampleSum / $sampleCount
        Write-Host "  Column $col [$header]: 有效数值行=$sampleCount, 平均值=$( [Math]::Round($avg, 2) )"
    }
}

$wb.Close($false)
$xl.Quit()
