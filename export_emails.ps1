# Email Full Export - All Emails Organized by Category
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$inbox = $namespace.GetDefaultFolder(6)

Write-Host "Starting full email export..."

$items = $inbox.Items
$totalCount = $items.Count

# Categories to track
$categories = @{
    "BOSS直聘招聘" = @()
    "飞书通知" = @()
    "脉脉Moka" = @()
    "美团外卖" = @()
    "中国移动" = @()
    "高途内部" = @()
    "其他" = @()
}

# Keywords for categorization
$bossKeywords = @("BOSS", "zhipin", "简历", "候选人", "面试", "招聘", "应聘")
$feishuKeywords = @("feishu", "飞书", "会议", "日历", "审批")
$mokaKeywords = @("mokahr", "脉脉")
$meituanKeywords = @("meituan", "美团", "外卖", "订单")
$mobileKeywords = @("10086", "10010", "中国移动", "流量", "话单")
$gaotuKeywords = @("gaotu", "高途", "市场部", "运营", "绩效")

$allEmails = @()

Write-Host "Processing $totalCount emails..."

for ($i = 1; $i -le $totalCount; $i++) {
    try {
        $item = $items.Item($i)
        
        $email = @{
            Index = $i
            Subject = ""
            Sender = ""
            Date = ""
            Category = "其他"
        }
        
        if ($item.Subject) {
            $email.Subject = $item.Subject.ToString()
        }
        if ($item.SenderEmailAddress) {
            $email.Sender = $item.SenderEmailAddress.ToString()
        }
        if ($item.ReceivedTime) {
            $email.Date = (Get-Date $item.ReceivedTime).ToString("yyyy-MM-dd HH:mm")
        }
        
        $text = $email.Subject + " " + $email.Sender
        
        # Categorize
        $category = "其他"
        
        foreach ($kw in $bossKeywords) {
            if ($text -match $kw) { $category = "BOSS直聘招聘"; break }
        }
        if ($category -eq "其他") {
            foreach ($kw in $feishuKeywords) {
                if ($text -match $kw) { $category = "飞书通知"; break }
            }
        }
        if ($category -eq "其他") {
            foreach ($kw in $mokaKeywords) {
                if ($text -match $kw) { $category = "脉脉Moka"; break }
            }
        }
        if ($category -eq "其他") {
            foreach ($kw in $meituanKeywords) {
                if ($text -match $kw) { $category = "美团外卖"; break }
            }
        }
        if ($category -eq "其他") {
            foreach ($kw in $mobileKeywords) {
                if ($text -match $kw) { $category = "中国移动"; break }
            }
        }
        if ($category -eq "其他") {
            foreach ($kw in $gaotuKeywords) {
                if ($text -match $kw) { $category = "高途内部"; break }
            }
        }
        
        $email.Category = $category
        $allEmails += $email
        
        if ($i % 500 -eq 0) {
            Write-Host "Processed $i emails..."
        }
        
    } catch {}
}

Write-Host "Export complete! Total: $($allEmails.Count)"

# Output by category
foreach ($cat in $categories.Keys) {
    $catEmails = $allEmails | Where-Object { $_.Category -eq $cat }
    Write-Host ""
    Write-Host "=== $cat ($($catEmails.Count) emails) ==="
    
    $catEmails | ForEach-Object {
        $date = if ($_.Date) { $_.Date } else { "N/A" }
        $subject = if ($_.Subject) { $_.Subject } else { "(无主题)" }
        Write-Host "[$date] $subject"
    }
}

Write-Host ""
Write-Host "Export finished!"
