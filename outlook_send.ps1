$params = Get-Content "C:\Users\tls2\.claude\projects\H--\_email_params.json" -Raw | ConvertFrom-Json
try {
    $outlook = New-Object -ComObject Outlook.Application
    $mail = $outlook.CreateItem(0)
    $mail.To = $params.To
    $mail.Subject = $params.Subject
    $mail.HTMLBody = $params.Body
    if ($params.From) {
        $sendAccount = $null
        foreach ($account in $outlook.Session.Accounts) {
            if ($account.SmtpAddress -ieq $params.From) {
                $sendAccount = $account
                break
            }
        }
        if ($sendAccount) {
            $mail.SendUsingAccount = $sendAccount
        } else {
            $mail.SentOnBehalfOfName = $params.From
        }
    }
    $mail.Send()
    "Sent." | Out-File "C:\Users\tls2\.claude\projects\H--\_send_result.txt"
} catch {
    "FAILED: $_" | Out-File "C:\Users\tls2\.claude\projects\H--\_send_result.txt"
}
