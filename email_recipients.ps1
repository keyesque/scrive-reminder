$Outlook = New-Object -ComObject Outlook.Application
$recipients = Get-Content "recipients.json" -Encoding UTF8 | ConvertFrom-Json

foreach ($rcpt in $recipients) {
    $Mail = $Outlook.CreateItem(0)
    $Mail.To = $rcpt.Email
    $Mail.Subject = "Signering av avtal för $($rcpt.Name)"
    $Mail.HTMLBody = @"
<html lang="sv">
<head>
    <meta charset="utf-8">
</head>
<body>
    <p>Hej!</p>

    <p>Detta är en påminnelse om att du behöver signera avtalet för din Ragn-Sells-dator via Scrive.</p>

    <p><strong>ID för din order:</strong> $($rcpt.ID)</p>

    <p>Tack på förhand!</p>

    <p>Mvh,<br>
    Karl-Axel</p>
</body>
</html>
"@

    $Mail.Send()
}

$Outlook.Quit()
