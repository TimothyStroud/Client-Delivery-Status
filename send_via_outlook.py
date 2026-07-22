"""
Send email via Outlook COM through an interactive scheduled task,
working around window station isolation.
"""
import subprocess, time, os, json

BASE = r'C:\Users\tls2\.claude\projects\H--'
PARAMS_FILE  = os.path.join(BASE, '_email_params.json')
RESULT_FILE  = os.path.join(BASE, '_send_result.txt')
SEND_SCRIPT  = os.path.join(BASE, 'outlook_send.ps1')
TASK_NAME    = '_OutlookSend'

def send(to, subject, body, from_address=None):
    params = {'To': to, 'Subject': subject, 'Body': body}
    if from_address:
        params['From'] = from_address
    with open(PARAMS_FILE, 'w', encoding='utf-8') as f:
        json.dump(params, f)

    if os.path.exists(RESULT_FILE):
        os.remove(RESULT_FILE)

    # Register an interactive scheduled task that fires in 3 seconds
    ps = f"""
$action = New-ScheduledTaskAction -Execute "powershell.exe" `
    -Argument '-NoProfile -NonInteractive -File "{SEND_SCRIPT}"'
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddSeconds(3)
$trigger.EndBoundary = $null
$principal = New-ScheduledTaskPrincipal -UserId $Env:USERNAME -LogonType Interactive
$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit "00:05:00"
Register-ScheduledTask -TaskName "{TASK_NAME}" -Action $action -Trigger $trigger `
    -Principal $principal -Settings $settings -Force | Out-Null
"""
    subprocess.run(['powershell', '-NoProfile', '-NonInteractive', '-Command', ps],
                   capture_output=True)

    # Wait for the result. The helper task fires at +3s, launches Outlook COM,
    # sends, then WAITS for the Outbox to actually drain (up to 150s) before
    # writing "Sent." -- so the result can legitimately take a few minutes,
    # especially on the first unattended send of the day when classic OUTLOOK.EXE
    # is cold-started (the user's open app is New Outlook / olk.exe, which has no
    # COM interface, so COM always spins up its own throwaway classic instance).
    # Poll longer than the helper's 5-min ExecutionTimeLimit floor.
    for _ in range(210):
        time.sleep(1)
        if os.path.exists(RESULT_FILE):
            result = open(RESULT_FILE, encoding='utf-16').read().strip()
            subprocess.run(
                ['powershell', '-NoProfile', '-Command',
                 f'Unregister-ScheduledTask -TaskName "{TASK_NAME}" -Confirm:$false -ErrorAction SilentlyContinue'],
                capture_output=True)
            return result
    return 'TIMEOUT'


if __name__ == '__main__':
    result = send(
        to='DataOperations@machinify.com',
        subject='RAMP Alert Test',
        body='This is a test email from the RAMP unconfigured files alert script.'
    )
    print(result)
