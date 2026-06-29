import sys, os, re, json, urllib.request

MSG = (
    ":white_check_mark: Kaiser Pareo Prepay - Load & Snap SUCCESS (data date 06/28/2026)\n"
    "Tracked files: MA + SC\n"
    "0100 Stage: SUCCESS - both MA & SC staged (06/28 07:33)\n"
    "0110 Load: SUCCESS (06/28 09:01-09:40)\n"
    "0120 Snap: SUCCESS (06/28 09:41-10:00)\n"
    "Result: both MA & SC staged and loaded - none missing.\n"
    "MA: All_Canonical_Prepay_Claim_Lines_MA_20260628010000.csv\n"
    "SC: All_Canonical_Prepay_Claim_Lines_SC_20260628000000.csv"
)


def sanitize(text):
    text = text.replace('<!here> ', '').replace('<!here>', '')
    text = re.sub(r'(?m)^> ?', '', text)
    text = text.replace('*', '').replace('`', '')
    return text


def post(url, text):
    data = json.dumps({'Text': text}).encode('utf-8')
    req = urllib.request.Request(url, data=data, headers={'Content-Type': 'application/json'})
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.status, r.read().decode('utf-8', 'replace').strip()


def main():
    targets = {
        'support (#team-rdp-operations-support)': r'H:\slack_wf_support.txt',
        'kaiserprepay (#rps_kaiserprepay_discussion)': r'H:\slack_wf_kaiserprepay.txt',
    }
    text = sanitize(MSG)
    for label, f in targets.items():
        url = ''
        if os.path.exists(f):
            url = open(f, encoding='utf-8').read().strip()
        if not url:
            print(f"SKIP {label}: no webhook URL at {f}")
            continue
        try:
            st, body = post(url, text)
            print(f"POSTED {label}: HTTP {st} {body[:80]}")
        except Exception as e:
            print(f"ERROR {label}: {e}")


if __name__ == '__main__':
    main()
