$params = Get-Content "C:\Users\tls2\.claude\projects\H--\_email_params.json" -Raw | ConvertFrom-Json
$resultFile = "C:\Users\tls2\.claude\projects\H--\_send_result.txt"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $ns = $outlook.GetNamespace("MAPI")
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

    # $mail.Send() only QUEUES the item to the Outbox and returns immediately;
    # the actual transmission happens asynchronously on Outlook's next sync. This
    # process is the COM host for the (often cold-started) classic OUTLOOK.EXE, so
    # if we exit right after Send() the instance can be torn down before the Outbox
    # flushes -> "Sent." logged but nothing delivered. The user's open Outlook is
    # the New Outlook (olk.exe), which has NO COM interface, so COM always spins up
    # a separate throwaway classic instance rather than reusing a warm one.
    # FIX: nudge a send/receive, then WAIT until the default Outbox actually drains
    # before reporting success. Only write "Sent." once the Outbox is empty.
    $outbox = $ns.GetDefaultFolder(4)   # olFolderOutbox

    # Kick every configured send/receive group so we don't wait on the idle timer.
    try {
        $syncs = $ns.SyncObjects
        for ($i = 1; $i -le $syncs.Count; $i++) { $syncs.Item($i).Start() }
    } catch {}

    $deadline = (Get-Date).AddSeconds(150)
    $flushed = $false
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds 2
        $pending = $outbox.Items.Count
        if ($pending -eq 0) { $flushed = $true; break }
        # re-nudge periodically in case the first sync didn't pick it up
        try {
            $syncs = $ns.SyncObjects
            for ($i = 1; $i -le $syncs.Count; $i++) { $syncs.Item($i).Start() }
        } catch {}
    }

    if ($flushed) {
        "Sent." | Out-File $resultFile
    } else {
        # Left the Outbox non-empty; do NOT report a clean "Sent." so the caller
        # won't mark the notification as delivered and will retry next run.
        "PENDING: submitted but Outbox did not drain within timeout" | Out-File $resultFile
    }
} catch {
    "FAILED: $_" | Out-File $resultFile
}
