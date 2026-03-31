# Outlook Email Deep Analysis - English Only
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$inbox = $namespace.GetDefaultFolder(6)

Write-Host "=================================================="
Write-Host "Email Deep Analysis"
Write-Host "=================================================="
Write-Host ""

$items = $inbox.Items
$totalCount = $items.Count
Write-Host "Total emails: $totalCount"

# 1. Monthly Distribution
Write-Host ""
Write-Host "=== Monthly Distribution ==="
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
    Write-Host $_.Key ": " $_.Value
}

# 2. More Keywords
Write-Host ""
Write-Host "=== More Keyword Analysis ==="
$keywords = @(
    "BOSS",
    "zhipin", 
    "mokahr",
    "feishu",
    "meituan",
    "SF",
    "shunfeng",
    "bank",
    "10086",
    "gaotu",
    "xueersi",
    "douyin",
    "resume",
    "interview"
)

foreach ($kw in $keywords) {
    $count = 0
    for ($i = 1; $i -le $totalCount; $i++) {
        try {
            $item = $items.Item($i)
            $subject = $item.Subject
            $sender = $item.SenderEmailAddress
            $text = $subject.ToString() + " " + $sender.ToString()
            if ($text -match $kw) {
                $count++
            }
        } catch {}
    }
    if ($count -gt 0) {
        Write-Host "[" $kw "]: " $count
    }
}

# 3. Summary Statistics
Write-Host ""
Write-Host "=== Summary by Category ==="

# Recruitment
$recruitCount = 0
for ($i = 1; $i -le $totalCount; $i++) {
    try {
        $item = $items.Item($i)
        $text = $item.Subject.ToString() + " " + $item.SenderEmailAddress.ToString()
        if ($text -match "BOSS|zhipin|mokahr|resume|candidate|interview") {
            $recruitCount++
        }
    } catch {}
}
Write-Host "Recruitment: " $recruitCount

# Logistics
$logisticsCount = 0
for ($i = 1; $i -le $totalCount; $i++) {
    try {
        $item = $items.Item($i)
        $text = $item.Subject.ToString()
        if ($text -match "SF|shunfeng|JD|jingdong|logistics|delivery|package") {
            $logisticsCount++
        }
    } catch {}
}
Write-Host "Logistics: " $logisticsCount

# Telecom
$telecomCount = 0
for ($i = 1; $i -le $totalCount; $i++) {
    try {
        $item = $items.Item($i)
        $text = $item.Subject.ToString()
        if ($text -match "10086|10010|mobile|carrier|bill") {
            $telecomCount++
        }
    } catch {}
}
Write-Host "Telecom: " $telecomCount

Write-Host ""
Write-Host "Done!"
