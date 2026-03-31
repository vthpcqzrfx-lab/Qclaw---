# Full Email Export with Body
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
        $body = ""
        
        try { if ($item.Subject) { $subject = [string]$item.Subject } } catch {}
        try { if ($item.SenderEmailAddress) { $sender = [string]$item.SenderEmailAddress } } catch {}
        try { if ($item.ReceivedTime) { $date = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm") } } catch {}
        
        try { 
            if ($item.Body) { 
                $body = [string]$item.Body 
                # Limit body length
                if ($body.Length -gt 2000) { $body = $body.Substring(0, 2000) + "..." }
            } 
        } catch {}
        
        if (-not $subject) { $subject = "(No Subject)" }
        if (-not $sender) { $sender = "(Unknown)" }
        if (-not $body) { $body = "(No Body)" }
        
        $entry = @{
            Index = $i
            Date = $date
            Sender = $sender
            Subject = $subject
            Body = $body
        }
        
        $output += $entry
        
        if ($i % 100 -eq 0) {
            Write-Host "Processed $i emails..."
        }
        
    } catch {}
}

# Convert to JSON
$json = $output | ConvertTo-Json -Depth 3

$utf8Bom = New-Object System.Text.UTF8Encoding $true
[System.IO.File]::WriteAllText("C:\Users\zhaoyang26\Desktop\all_emails_with_body.json", $json, $utf8Bom)

Write-Host "Done!"
Write-Host "Total: $($output.Count)"
