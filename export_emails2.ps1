# Email Full Export
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$inbox = $namespace.GetDefaultFolder(6)

Write-Host "========================================="
Write-Host "Email Full Export - By Category"
Write-Host "========================================="

$items = $inbox.Items
$totalCount = $items.Count

$allEmails = @()

Write-Host "Processing $totalCount emails..."

for ($i = 1; $i -le $totalCount; $i++) {
    try {
        $item = $items.Item($i)
        
        $subject = ""
        $sender = ""
        $date = ""
        
        if ($item.Subject) { $subject = $item.Subject.ToString() }
        if ($item.SenderEmailAddress) { $sender = $item.SenderEmailAddress.ToString() }
        if ($item.ReceivedTime) { $date = (Get-Date $item.ReceivedTime).ToString("yyyy-MM-dd HH:mm") }
        
        $text = $subject + " " + $sender
        
        # Categorize
        $cat = "Other"
        
        if ($text -match "BOSS|zhipin|resume|candidate|interview") { $cat = "BOSS_Zhipin" }
        elseif ($text -match "feishu|calendar") { $cat = "Feishu" }
        elseif ($text -match "mokahr") { $cat = "Moka" }
        elseif ($text -match "meituan") { $cat = "Meituan" }
        elseif ($text -match "10086|10010|mobile") { $cat = "Mobile" }
        elseif ($text -match "gaotu") { $cat = "Gaotu" }
        
        $email = @{
            Index = $i
            Subject = $subject
            Sender = $sender
            Date = $date
            Category = $cat
        }
        
        $allEmails += $email
        
        if ($i % 500 -eq 0) { Write-Host "Processed $i..." }
        
    } catch {}
}

Write-Host ""
Write-Host "Total processed: $($allEmails.Count)"
Write-Host ""

# Output by category
$categories = @("BOSS_Zhipin", "Feishu", "Moka", "Meituan", "Mobile", "Gaotu", "Other")

foreach ($cat in $categories) {
    $catEmails = $allEmails | Where-Object { $_.Category -eq $cat }
    $count = $catEmails.Count
    
    Write-Host ""
    Write-Host "========================================="
    Write-Host "$cat ($count emails)"
    Write-Host "========================================="
    
    $catEmails | ForEach-Object {
        $d = $_.Date
        $s = $_.Subject
        if (-not $s) { $s = "(No Subject)" }
        Write-Host "[$d] $s"
    }
}

Write-Host ""
Write-Host "Export Complete!"
