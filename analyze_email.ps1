# Outlook邮件分析脚本
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$inbox = $namespace.GetDefaultFolder(6)  # 6 = inbox

Write-Host "="*60
Write-Host "📧 邮箱详细分析"
Write-Host "="*60
Write-Host ""

# 获取所有邮件（只取前5000封作为样本）
$items = $inbox.Items
$totalCount = $items.Count
Write-Host "总邮件数: $totalCount"

# 按发件人统计
Write-Host ""
Write-Host "=== 按发件人域名统计 (Top 30) ==="
$senders = @{}
for ($i = 1; $i -le [Math]::Min($totalCount, 5000); $i++) {
    try {
        $item = $items.Item($i)
        if ($item.SenderEmailAddress) {
            $email = $item.SenderEmailAddress
            # 提取域名
            if ($email -match "@(.+)$") {
                $domain = $matches[1]
                $senders[$domain] = ($senders[$domain] ?? 0) + 1
            }
        }
    } catch {
        # 跳过错误
    }
    if ($i % 500 -eq 0) {
        Write-Host "已分析 $i 封..."
    }
}

# 排序并输出
$senders.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 30 | ForEach-Object {
    Write-Host "  $($_.Key): $($_.Value) 封"
}

Write-Host ""
Write-Host "=== 按年份分布 ==="

$yearStats = @{}
for ($i = 1; $i -le [Math]::Min($totalCount, 3000); $i++) {
    try {
        $item = $items.Item($i)
        if ($item.ReceivedTime) {
            $year = (Get-Date $item.ReceivedTime).Year
            $yearStats[$year] = ($yearStats[$year] ?? 0) + 1
        }
    } catch {}
}

$yearStats.GetEnumerator() | Sort-Object Name | ForEach-Object {
    Write-Host "  $($_.Key)年: $($_.Value) 封"
}

Write-Host ""
Write-Host "=== 搜索关键词统计 ==="

# 常见关键词
$keywords = @("BOSS", "招聘", "简历", "京东", "淘宝", "天猫", "顺丰", "圆通", "中通", "韵达", "银行", "中国移动", "中国联通", "腾讯", "阿里", "字节", "抖音", "高途", "学而思", "作业帮")

foreach ($kw in $keywords) {
    $count = 0
    for ($i = 1; $i -le [Math]::Min($totalCount, 3000); $i++) {
        try {
            $item = $items.Item($i)
            if ($item.Subject -match $kw) {
                $count++
            }
        } catch {}
    }
    if ($count -gt 0) {
        Write-Host "  关键词[$kw]: $count 封"
    }
}

Write-Host ""
Write-Host "分析完成!"
