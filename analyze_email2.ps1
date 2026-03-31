# Outlook Email Analysis
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$inbox = $namespace.GetDefaultFolder(6)

Write-Host "================================================"
Write-Host "Email Analysis - Total mailbox analysis"
Write-Host "================================================"
Write-Host ""

$items = $inbox.Items
$totalCount = $items.Count
Write-Host "Total emails: $totalCount"
Write-Host ""

# Domain statistics
Write-Host "=== Top Sender Domains ==="
$senders = @{}
for ($i = 1; $i -le [Math]::Min($totalCount, 3000); $i++) {
    try {
        $item = $items.Item($i)
        if ($item.SenderEmailAddress) {
            $email = $item.SenderEmailAddress
            if ($email -match "@(.+)$") {
                $domain = $matches[1]
                if ($senders.ContainsKey($domain)) {
                    $senders[$domain] = $senders[$domain] + 1
                } else {
                    $senders[$domain] = 1
                }
            }
        }
    } catch {}
}

$senders.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 25 | ForEach-Object {
    Write-Host $_.Key ":" $_.Value
}

Write-Host ""
Write-Host "=== Year Distribution ==="
$yearStats = @{}
for ($i = 1; $i -le [Math]::Min($totalCount, 2000); $i++) {
    try {
        $item = $items.Item($i)
        if ($item.ReceivedTime) {
            $year = (Get-Date $item.ReceivedTime).Year
            if ($yearStats.ContainsKey($year)) {
                $yearStats[$year] = $yearStats[$year] + 1
            } else {
                $yearStats[$year] = 1
            }
        }
    } catch {}
}

$yearStats.GetEnumerator() | Sort-Object Name | ForEach-Object {
    Write-Host $_.Key "year:" $_.Value "emails"
}

Write-Host ""
Write-Host "=== Keyword Search ==="
$keywords = @("BOSS", "zhipin", "JD", "jingdong", "taobao", "tmall", "SF", "shunfeng", "bank", "10086", "tencent", "alibaba", "bytedance", "douyin", "gaotu", "xueersi", "zuoyebang")

foreach ($kw in $keywords) {
    $count = 0
    for ($i = 1; $i -le [Math]::Min($totalCount, 2000); $i++) {
        try {
            $item = $items.Item($i)
            $subject = $item.Subject
            if ($subject -and $subject -match $kw) {
                $count++
            }
        } catch {}
    }
    if ($count -gt 0) {
        Write-Host $kw ": " $count
    }
}

Write-Host ""
Write-Host "Done!"
