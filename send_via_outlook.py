"""
Send email via Outlook COM through an interactive scheduled task,
working around window station isolation.

CONCURRENCY (fixed 2026-08-07): every send used to share ONE set of files
(_email_params.json / _send_result.txt) and ONE helper task name
(_OutlookSend). Two jobs sending at the same minute -- e.g. the 8:00am
"RAMP Unconfigured Email Alerts" and "Load Completion SLA Check" tasks --
clobbered each other's params, so a single helper run sent whichever
params landed last while BOTH callers read the same "Sent." result and
recorded success. That silently dropped the 8/3 Wellmark and 8/5
Caresource SLA emails. Now each call gets unique params/result files and
a unique helper task name, and calls are serialised behind a lock file
(the Outbox-drain check is global state, so two concurrent sends would
also confuse each other's "did it flush?" test).
"""
import subprocess, time, os, json, itertools

BASE = r'C:\Users\tls2\.claude\projects\H--'
SEND_SCRIPT  = os.path.join(BASE, 'outlook_send.ps1')
LOCK_FILE    = os.path.join(BASE, '_outlook_send.lock')

# Legacy fixed paths, kept only as the PS script's defaults.
PARAMS_FILE  = os.path.join(BASE, '_email_params.json')
RESULT_FILE  = os.path.join(BASE, '_send_result.txt')
TASK_NAME    = '_OutlookSend'

LOCK_STALE_SECONDS = 420   # a send can legitimately take ~5 min (cold Outlook)
LOCK_WAIT_SECONDS  = 600   # wait this long for another job's send to finish

_counter = itertools.count(1)


def _acquire_lock():
    """Cross-process mutex so only one Outlook COM send runs at a time."""
    deadline = time.time() + LOCK_WAIT_SECONDS
    while True:
        try:
            fd = os.open(LOCK_FILE, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, str(os.getpid()).encode())
            os.close(fd)
            return True
        except FileExistsError:
            try:
                age = time.time() - os.path.getmtime(LOCK_FILE)
            except OSError:
                continue          # holder released it between the two calls
            if age > LOCK_STALE_SECONDS:
                # Holder died (task killed / machine rebooted mid-send).
                try:
                    os.remove(LOCK_FILE)
                except OSError:
                    pass
                continue
            if time.time() > deadline:
                return False
            time.sleep(2)


def _release_lock():
    try:
        os.remove(LOCK_FILE)
    except OSError:
        pass


def _ps(cmd):
    subprocess.run(['powershell', '-NoProfile', '-NonInteractive', '-Command', cmd],
                   capture_output=True)


def send(to, subject, body, from_address=None):
    if not _acquire_lock():
        return 'BUSY: another Outlook send held the lock too long'
    try:
        return _send_locked(to, subject, body, from_address)
    finally:
        _release_lock()


def _send_locked(to, subject, body, from_address):
    token = f'{os.getpid()}_{next(_counter)}'
    params_file = os.path.join(BASE, f'_email_params_{token}.json')
    result_file = os.path.join(BASE, f'_send_result_{token}.txt')
    task_name   = f'_OutlookSend_{token}'

    params = {'To': to, 'Subject': subject, 'Body': body}
    if from_address:
        params['From'] = from_address
    with open(params_file, 'w', encoding='utf-8') as f:
        json.dump(params, f)

    # Register an interactive scheduled task that fires in 3 seconds
    ps = f"""
$action = New-ScheduledTaskAction -Execute "powershell.exe" `
    -Argument '-NoProfile -NonInteractive -File "{SEND_SCRIPT}" -ParamsFile "{params_file}" -ResultFile "{result_file}"'
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddSeconds(3)
$trigger.EndBoundary = $null
$principal = New-ScheduledTaskPrincipal -UserId $Env:USERNAME -LogonType Interactive
$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit "00:05:00"
Register-ScheduledTask -TaskName "{task_name}" -Action $action -Trigger $trigger `
    -Principal $principal -Settings $settings -Force | Out-Null
"""
    _ps(ps)

    # Wait for the result. The helper task fires at +3s, launches Outlook COM,
    # sends, then WAITS for the Outbox to actually drain (up to 150s) before
    # writing "Sent." -- so the result can legitimately take a few minutes,
    # especially on the first unattended send of the day when classic OUTLOOK.EXE
    # is cold-started (the user's open app is New Outlook / olk.exe, which has no
    # COM interface, so COM always spins up its own throwaway classic instance).
    # Poll longer than the helper's 5-min ExecutionTimeLimit floor.
    result = 'TIMEOUT'
    for _ in range(210):
        time.sleep(1)
        if os.path.exists(result_file):
            result = open(result_file, encoding='utf-16').read().strip()
            break

    _ps(f'Unregister-ScheduledTask -TaskName "{task_name}" -Confirm:$false '
        f'-ErrorAction SilentlyContinue')
    for f in (params_file, result_file):
        try:
            os.remove(f)
        except OSError:
            pass
    return result


if __name__ == '__main__':
    result = send(
        to='DataOperations@machinify.com',
        subject='RAMP Alert Test',
        body='This is a test email from the RAMP unconfigured files alert script.'
    )
    print(result)
