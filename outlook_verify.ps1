$out = "C:\Users\tls2\.claude\projects\H--\_verify_result.txt"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $ns = $outlook.GetNamespace("MAPI")
    $lines = @()

    function Search-Folder($folder, $storeName) {
        try {
            $items = $folder.Items
            $n = $items.Count
            $matches = 0
            $hits = @()
            foreach ($it in $items) {
                try {
                    if ($it.Subject -like "*Caresource 0200 Load*") {
                        $so = $null; try { $so = $it.SentOn } catch {}
                        $hits += "    HIT SentOn=$so To=$($it.To) Subj=$($it.Subject)"
                        $matches++
                        if ($matches -ge 8) { break }
                    }
                } catch {}
            }
            if ($matches -gt 0) {
                $script:lines += "[$storeName / $($folder.Name)] total=$n matches=$matches"
                $script:lines += $hits
            }
        } catch {}
    }

    foreach ($store in $ns.Stores) {
        $sname = $store.DisplayName
        try {
            $root = $store.GetRootFolder()
            # try to find a Sent Items folder in this store
            foreach ($sub in $root.Folders) {
                if ($sub.Name -match "Sent") { Search-Folder $sub $sname }
            }
        } catch {
            $lines += "store '$sname' error: $_"
        }
    }
    if ($lines.Count -eq 0) { $lines += "NO CareSource items found in any store's Sent* folder" }
    ($lines -join "`n") | Out-File $out -Encoding utf8
} catch {
    "FAILED: $_" | Out-File $out -Encoding utf8
}
