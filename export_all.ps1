# Full Email Export - All Email Contents
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$inbox = $namespace.GetDefaultFolder(6)

$items = $inbox.Items
$totalCount = $items.Count

$output = @()

for ($i = 1; $i -le $totalCount; $i++) {
    try {
        $item = $items.Item($i)
        
        $subject = ""
        $sender = ""
        $date = ""
        
        try { if ($item.Subject) { $subject = $item.Subject } } catch {}
        try { if ($item.SenderEmailAddress) { $sender = $item.SenderEmailAddress } } catch {}
        try { if ($item.ReceivedTime) { $date = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm") } } catch {}
        
        if (-not $subject) { $subject = "(No Subject)" }
        if (-not $sender) { $sender = "(Unknown)" }
        
        $line = "$date | $sender | $subject"
        $output += $line
        
    } catch {}
}

$output | Out-File -FilePath "C:\Users\zhaoyang26\Desktop\所有邮件列表.txt" -Encoding UTF8

Write-Host "Exported to: C:\Users\zhaoyang26\Desktop\所有邮件列表.txt"
Write-Host "Total emails: $($output.Count)"
