# Full Email Export - UTF8
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
        
        try { if ($item.Subject) { $subject = [string]$item.Subject } } catch {}
        try { if ($item.SenderEmailAddress) { $sender = [string]$item.SenderEmailAddress } } catch {}
        try { if ($item.ReceivedTime) { $date = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm") } } catch {}
        
        if (-not $subject) { $subject = "(No Subject)" }
        if (-not $sender) { $sender = "(Unknown)" }
        
        $line = "$date | $sender | $subject"
        $output += $line
        
    } catch {}
}

# Use UTF-8 with BOM
$utf8Bom = New-Object System.Text.UTF8Encoding $true
[System.IO.File]::WriteAllLines("C:\Users\zhaoyang26\Desktop\all_emails.txt", $output, $utf8Bom)

Write-Host "Done"
Write-Host "Total: $($output.Count)"
