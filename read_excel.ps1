$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false

$files = @(
    "C:\Users\zhaoyang26\Desktop\city_level_5_251122.xlsx",
    "C:\Users\zhaoyang26\Desktop\信息流分析-头条&广点通-2月&3月.xlsx",
    "C:\Users\zhaoyang26\Desktop\线索明细数据.xlsx",
    "C:\Users\zhaoyang26\Desktop\星光项目部2026年收款及效率目标_执行版0311.xlsx"
)

foreach($path in $files) {
    $fname = [System.IO.Path]::GetFileName($path)
    Write-Host "=== $fname ==="
    try {
        $wb = $xl.Workbooks.Open($path)
        $ws = $wb.Sheets.Item(1)
        $header = @()
        for($col=1; $col -le 10; $col++) {
            $val = $ws.Cells.Item(1,$col).Text
            $header += if($val) { $val } else { "(blank)" }
        }
        Write-Host "Headers: $($header -join ' | ')"
        for($row=2; $row -le 4; $row++) {
            $r = @()
            for($col=1; $col -le 10; $col++) {
                $val = $ws.Cells.Item($row,$col).Text
                $r += if($val) { $val } else { "-" }
            }
            Write-Host "Row$($row): $($r -join ' | ')"
        }
        $wb.Close($false)
    } catch {
        Write-Host "Error: $_"
    }
    Write-Host ""
}

$xl.Quit()
