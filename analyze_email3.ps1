# Outlook Email Deep Analysis
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$inbox = $namespace.GetDefaultFolder(6)

Write-Host "=================================================="
Write-Host "Email Deep Analysis - More Dimensions"
Write-Host "=================================================="
Write-Host ""

$items = $inbox.Items
$totalCount = $items.Count
Write-Host "Total emails in analysis: $totalCount"

# 1. 按月份分布
Write-Host ""
Write-Host "=== Monthly Distribution (2025-2026) ==="
$monthStats = @{}
for ($i = 1; $i -le $totalCount; $i++) {
    try {
        $item = $items.Item($i)
        if ($item.ReceivedTime) {
            $date = Get-Date $item.ReceivedTime
            $monthKey = $date.ToString("yyyy-MM")
            if ($monthStats.ContainsKey($monthKey)) {
                $monthStats[$monthKey] = $monthStats[$monthKey] + 1
            } else {
                $monthStats[$monthKey] = 1
            }
        }
    } catch {}
}

$monthStats.GetEnumerator() | Sort-Object Name | ForEach-Object {
    Write-Host $_.Key ": " $_.Value " emails"
}

# 2. 招聘相关详细
Write-Host ""
Write-Host "=== Recruitment Emails Detail ==="
$recruitKeywords = @("BOSS", "zhipin", "mokahr", "简历", "面试", "候选人", "招聘", "猎头", "招聘网")
$recruitCount = @{}
foreach ($kw in $recruitKeywords) {
    $count = 0
    for ($i = 1; $i -le $totalCount; $i++) {
        try {
            $item = $items.Item($i)
            $subject = $item.Subject
            if ($subject -and $subject -match $kw) {
                $count++
            }
        } catch {}
    }
    if ($count -gt 0) {
        Write-Host "  [$kw]: $count"
    }
}

# 3. 电商/物流
Write-Host ""
Write-Host "=== E-commerce & Logistics ==="
$ecomKeywords = @("京东", "淘宝", "天猫", "拼多多", "顺丰", "圆通", "中通", "韵达", "申通", "美团", "饿了么", "JD", "SF Express", "物流", "快递", "订单", "发货", "签收")
$ecomCount = 0
for ($i = 1; $i -le $totalCount; $i++) {
    try {
        $item = $items.Item($i)
        $subject = $item.Subject
        if ($subject) {
            foreach ($kw in $ecomKeywords) {
                if ($subject -match $kw) {
                    $ecomCount++
                    break
                }
            }
        }
    } catch {}
}
Write-Host "  Total e-commerce/logistics: $ecomCount"

# 4. 金融/银行
Write-Host ""
Write-Host "=== Finance & Banking ==="
$financeKeywords = @("银行", "信用卡", "还款", "理财", "转账", "支付", "支付宝", "微信支付", "花呗", "借呗", "保险", "证券", "基金")
$financeCount = 0
for ($i = 1; $i -le $totalCount; $i++) {
    try {
        $item = $items.Item($i)
        $subject = $item.Subject
        if ($subject) {
            foreach ($kw in $financeKeywords) {
                if ($subject -match $kw) {
                    $financeCount++
                    break
                }
            }
        }
    } catch {}
}
Write-Host "  Total finance/banking: $financeCount"

# 5. 运营商
Write-Host ""
Write-Host "=== Telecom ==="
$telecomKeywords = @("10086", "10010", "中国移动", "中国联通", "中国电信", "话单", "流量", "账单")
$telecomCount = 0
for ($i = 1; $i -le $totalCount; $i++) {
    try {
        $item = $items.Item($i)
        $subject = $item.Subject
        if ($subject) {
            foreach ($kw in $telecomKeywords) {
                if ($subject -match $kw) {
                    $telecomCount++
                    break
                }
            }
        }
    } catch {}
}
Write-Host "  Total telecom: $telecomCount"

# 6. 内部邮件
Write-Host ""
Write-Host "=== Internal (Gaotu) Emails ==="
$internalKeywords = @("高途", "gaotu", "GaoTu", "市场部", "运营", "绩效", "汇报", "会议", "通知")
$internalCount = 0
for ($i = 1; $i -le $totalCount; $i++) {
    try {
        $item = $items.Item($i)
        $subject = $item.Subject
        $sender = $item.SenderEmailAddress
        if ($subject -or $sender) {
            foreach ($kw in $internalKeywords) {
                $match = $false
                if ($subject -and $subject -match $kw) { $match = $true }
                if ($sender -and $sender -match $kw) { $match = $true }
                if ($match) { $internalCount++; break }
            }
        }
    } catch {}
}
Write-Host "  Total internal: $internalCount"

# 7. 订阅/营销
Write-Host ""
Write-Host "=== Subscriptions & Marketing ==="
$subKeywords = @("退订", "订阅", "Newsletter", "每周", "每月", "资讯", "简报", "推送")
$subCount = 0
for ($i = 1; $i -le $totalCount; $i++) {
    try {
        $item = $items.Item($i)
        $subject = $item.Subject
        if ($subject) {
            foreach ($kw in $subKeywords) {
                if ($subject -match $kw) {
                    $subCount++
                    break
                }
            }
        }
    } catch {}
}
Write-Host "  Total subscriptions: $subCount"

Write-Host ""
Write-Host "Analysis Complete!"
