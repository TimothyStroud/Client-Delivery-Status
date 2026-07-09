import re
html = open("ClientDeliveryStatus.html", encoding="utf-8").read()
for pm in re.finditer(r'id="(tab-[^"]+)"(.*?)(?=<section class="month-panel|\Z)', html, re.S):
    panel, body = pm.group(1), pm.group(2)
    for m in re.finditer(r'data-client="JHHC Passfile">[^<]*</td>\s*<td class="([^"]*)">([^<]*)</td>', body):
        mk = m.group(2).replace('✓', '[CHECK]')
        print(f'{panel:22} classes={m.group(1)!r:30} marker={mk!r}')
