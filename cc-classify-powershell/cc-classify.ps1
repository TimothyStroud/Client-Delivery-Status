<#
.SYNOPSIS
  cc-classify (PowerShell) - classify Claude Code token spend by initiative.

  Windows port of vlognow/cc-classify's reporting. Scans the local Claude Code
  session transcripts (~\.claude\projects\<initiative>\*.jsonl), prices every
  assistant turn with authoritative Anthropic rates (including the prompt-cache
  5m / 1h write and read tiers), and reports spend.

.PARAMETER Command
  summary       Capitalization report + per-initiative spend (default)
  initiatives   Per-initiative spend only
  sessions      Per-session detail (use -Initiative to scope)

.PARAMETER Since
  1month | Nd | YYYY-MM-DD   Start of the window (default 1month).

.PARAMETER Until
  YYYY-MM-DD                 End of the window, inclusive (default: today).

.PARAMETER Initiative
  For 'sessions': which initiative to break out.

.NOTES
  CAPITALIZATION BUCKETS: real cc-classify assigns Dev / COS / Mixed / Strategy
  from rules in ~/.config/cc-classify/config.toml. Those rules are not available
  here, so every session defaults to the Dev (capitalizable) bucket. Override per
  initiative with the $BucketMap table below. Spend / tokens / sessions / the
  per-initiative view are computed exactly.
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [ValidateSet('summary', 'initiatives', 'sessions')]
    [string]$Command = 'summary',

    [string]$Since = '1month',
    [string]$Until,
    [string]$Initiative
)

$ProjectsDir = Join-Path $env:USERPROFILE '.claude\projects'

# Base per-MILLION-token rates (USD). Cache multipliers: write5m=1.25x input,
# write1h=2x input, read=0.1x input.
$Rates = [ordered]@{
    'claude-fable-5'    = @(10.0, 50.0)
    'claude-opus-4-8'   = @(5.0, 25.0)
    'claude-opus-4-7'   = @(5.0, 25.0)
    'claude-opus-4-6'   = @(5.0, 25.0)
    'claude-opus-4-5'   = @(5.0, 25.0)
    'claude-opus-4'     = @(5.0, 25.0)
    'claude-sonnet-4-6' = @(3.0, 15.0)
    'claude-sonnet-4-5' = @(3.0, 15.0)
    'claude-sonnet-4'   = @(3.0, 15.0)
    'claude-haiku-4-5'  = @(1.0, 5.0)
    'claude-haiku'      = @(1.0, 5.0)
}

# Optional: map an initiative (project-dir name) to a capitalization bucket.
# Anything not listed defaults to 'Dev'.
$BucketMap = @{
    # 'SomeInitiative' = 'Strategy'
}
$DefaultBucket = 'Dev'

# Optional: display-name overrides for initiatives (project-dir name -> label).
# Anything not listed shows its raw project-dir name.
$InitiativeNames = @{
    'H--' = 'RDP Data Operations'
}
function Get-InitiativeName([string]$key) {
    if ($InitiativeNames.ContainsKey($key)) { return $InitiativeNames[$key] }
    return $key
}

function Get-Rate([string]$model) {
    if ([string]::IsNullOrEmpty($model)) { return $null }
    foreach ($k in $Rates.Keys) {
        if ($model.StartsWith($k)) { return $Rates[$k] }
    }
    return $null
}

function Resolve-Since([string]$s) {
    $s = $s.Trim().ToLower()
    $now = (Get-Date).ToUniversalTime()
    if ($s -eq '1month' -or $s -eq '1m' -or $s -eq 'month') { return $now.AddDays(-30) }
    if ($s -match '^(\d+)month$') { return $now.AddDays(-30 * [int]$Matches[1]) }
    if ($s -match '^(\d+)d(ay)?$') { return $now.AddDays(-[int]$Matches[1]) }
    return [datetime]::SpecifyKind([datetime]::ParseExact($s, 'yyyy-MM-dd', $null), 'Utc')
}

function Get-Num($v) { if ($null -eq $v) { return 0 } return [double]$v }

function Get-TurnCost($usage, [double[]]$rate) {
    $inpRate = $rate[0]; $outRate = $rate[1]
    $w5Rate = $inpRate * 1.25; $w1Rate = $inpRate * 2.0; $readRate = $inpRate * 0.10

    $inp = Get-Num $usage.input_tokens
    $out = Get-Num $usage.output_tokens
    $read = Get-Num $usage.cache_read_input_tokens
    $w5 = $null; $w1 = $null
    if ($usage.cache_creation) {
        $w5 = $usage.cache_creation.ephemeral_5m_input_tokens
        $w1 = $usage.cache_creation.ephemeral_1h_input_tokens
    }
    if ($null -eq $w5 -and $null -eq $w1) {
        $w5 = Get-Num $usage.cache_creation_input_tokens; $w1 = 0
    }
    $w5 = Get-Num $w5; $w1 = Get-Num $w1

    $usd = ($inp * $inpRate + $out * $outRate + $w5 * $w5Rate + $w1 * $w1Rate + $read * $readRate) / 1000000.0
    $tok = $inp + $out + $w5 + $w1 + $read
    return [pscustomobject]@{ Usd = $usd; Tokens = $tok }
}

# Local vars must NOT collide with the [string] params $Since/$Until
# (PowerShell vars are case-insensitive, so $since would re-coerce to string).
$sinceDt = Resolve-Since $Since
if ($Until) {
    $untilDt = [datetime]::SpecifyKind([datetime]::ParseExact($Until, 'yyyy-MM-dd', $null), 'Utc').AddDays(1)
} else {
    $untilDt = (Get-Date).ToUniversalTime().AddDays(1)
}

# Aggregations
$initAgg = @{}    # initiative -> @{Cost;Tokens;Sessions(hashset)}
$sessAgg = @{}    # "init|sid" -> @{Init;Sid;Cost;Tokens;First;Last}
$unpriced = @{}

Write-Host ''
Write-Host '  cc-classify (PowerShell)'
Write-Host ("  Window: {0:yyyy-MM-dd} to {1:yyyy-MM-dd}" -f $sinceDt, $untilDt.AddDays(-1))
Write-Host '  Scanning sessions...'

$files = Get-ChildItem -Path $ProjectsDir -Recurse -Filter '*.jsonl' -ErrorAction SilentlyContinue
foreach ($f in $files) {
    $initiative = Split-Path (Split-Path $f.FullName -Parent) -Leaf
    foreach ($line in [System.IO.File]::ReadLines($f.FullName)) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        try { $o = $line | ConvertFrom-Json } catch { continue }
        if ($o.type -ne 'assistant') { continue }
        $usage = $o.message.usage
        if (-not $usage) { continue }
        $ts = $null
        if ($o.timestamp) { try { $ts = ([datetime]$o.timestamp).ToUniversalTime() } catch { $ts = $null } }
        if ($null -eq $ts -or $ts -lt $sinceDt -or $ts -ge $untilDt) { continue }
        $model = [string]$o.message.model
        $rate = Get-Rate $model
        if ($null -eq $rate) {
            if (-not $unpriced.ContainsKey($model)) { $unpriced[$model] = 0 }
            $unpriced[$model]++
            continue
        }
        $c = Get-TurnCost $usage $rate

        if (-not $initAgg.ContainsKey($initiative)) {
            $initAgg[$initiative] = @{ Cost = 0.0; Tokens = 0.0; Sessions = (New-Object System.Collections.Generic.HashSet[string]) }
        }
        $sid = $o.sessionId
        if (-not $sid) { $sid = [System.IO.Path]::GetFileNameWithoutExtension($f.Name) }
        $e = $initAgg[$initiative]
        $e.Cost += $c.Usd; $e.Tokens += $c.Tokens; [void]$e.Sessions.Add($sid)

        $key = "$initiative|$sid"
        if (-not $sessAgg.ContainsKey($key)) {
            $sessAgg[$key] = @{ Init = $initiative; Sid = $sid; Cost = 0.0; Tokens = 0.0; First = $ts; Last = $ts }
        }
        $s = $sessAgg[$key]
        $s.Cost += $c.Usd; $s.Tokens += $c.Tokens
        if ($ts -lt $s.First) { $s.First = $ts }
        if ($ts -gt $s.Last) { $s.Last = $ts }
    }
}

$totalSessions = ($sessAgg.Keys).Count
Write-Host ("  Found {0} sessions." -f $totalSessions)
Write-Host ''

$totCost = 0.0; $totTok = 0.0
foreach ($k in $initAgg.Keys) { $totCost += $initAgg[$k].Cost; $totTok += $initAgg[$k].Tokens }

# ---- CAPITALIZATION REPORT (summary only) --------------------------------
if ($Command -eq 'summary') {
    $buckets = [ordered]@{ Dev = @{C = 0.0; T = 0.0; S = 0 }; COS = @{C = 0.0; T = 0.0; S = 0 };
        Mixed = @{C = 0.0; T = 0.0; S = 0 }; Strategy = @{C = 0.0; T = 0.0; S = 0 } }
    foreach ($k in $initAgg.Keys) {
        $b = $DefaultBucket
        if ($BucketMap.ContainsKey($k)) { $b = $BucketMap[$k] }
        $buckets[$b].C += $initAgg[$k].Cost
        $buckets[$b].T += $initAgg[$k].Tokens
        $buckets[$b].S += $initAgg[$k].Sessions.Count
    }
    Write-Host '  CAPITALIZATION REPORT'
    Write-Host ''
    Write-Host ('  {0,-10}{1,6}{2,14}{3,16}{4,8}' -f 'BUCKET', 'SESS', 'COST', 'TOKENS', '%')
    Write-Host ('  ' + ('-' * 54))
    foreach ($bn in $buckets.Keys) {
        $b = $buckets[$bn]
        $pct = 0.0; if ($totCost -gt 0) { $pct = $b.C / $totCost * 100 }
        Write-Host ('  {0,-10}{1,6}{2,14}{3,16}{4,7:N1}%' -f $bn, $b.S, ('$' + ('{0:N2}' -f $b.C)), ('{0:N0}' -f $b.T), $pct)
    }
    Write-Host ('  ' + ('-' * 54))
    Write-Host ('  {0,-10}{1,6}{2,14}{3,16}{4,7:N1}%' -f 'TOTAL', $totalSessions, ('$' + ('{0:N2}' -f $totCost)), ('{0:N0}' -f $totTok), 100.0)
    Write-Host ''
    $cap = $buckets['Dev'].C + $buckets['Mixed'].C
    $capPct = 0.0; if ($totCost -gt 0) { $capPct = $cap / $totCost * 100 }
    Write-Host ('  Capitalizable (Dev + Mixed): ${0:N2} ({1:N0}%)' -f $cap, $capPct)
    Write-Host ''
    Write-Host ''
}

# ---- INITIATIVES ---------------------------------------------------------
if ($Command -eq 'summary' -or $Command -eq 'initiatives') {
    Write-Host '  INITIATIVES'
    Write-Host '  Per-initiative spend, sorted by cost'
    Write-Host ''
    Write-Host ('  {0,-24}{1,14}{2,16}{3,6}{4,10}{5,8}' -f 'INITIATIVE', 'COST', 'TOKENS', 'SESS', 'BUCKET', 'CAP %')
    Write-Host ('  ' + ('-' * 78))
    $sorted = $initAgg.GetEnumerator() | Sort-Object { $_.Value.Cost } -Descending
    foreach ($kv in $sorted) {
        $b = $DefaultBucket; if ($BucketMap.ContainsKey($kv.Key)) { $b = $BucketMap[$kv.Key] }
        $cap = 100; if ($b -eq 'COS' -or $b -eq 'Strategy') { $cap = 0 }
        Write-Host ('  {0,-24}{1,14}{2,16}{3,6}{4,10}{5,7}%' -f (Get-InitiativeName $kv.Key), ('$' + ('{0:N2}' -f $kv.Value.Cost)), ('{0:N0}' -f $kv.Value.Tokens), $kv.Value.Sessions.Count, $b, $cap)
    }
    Write-Host ('  ' + ('-' * 78))
    Write-Host ('  {0,-24}{1,14}{2,16}{3,6}' -f 'TOTAL', ('$' + ('{0:N2}' -f $totCost)), ('{0:N0}' -f $totTok), $totalSessions)
    Write-Host ''
}

# ---- SESSIONS ------------------------------------------------------------
if ($Command -eq 'sessions') {
    $rows = $sessAgg.Values | Where-Object { -not $Initiative -or $_.Init -eq $Initiative } | Sort-Object { $_.Cost } -Descending
    $sc = 0.0; $st = 0.0; foreach ($r in $rows) { $sc += $r.Cost; $st += $r.Tokens }
    $scope = if ($Initiative) { "'$(Get-InitiativeName $Initiative)'" } else { 'all initiatives' }
    Write-Host ("  SESSIONS in {0}  (sorted by cost)" -f $scope)
    Write-Host ('  ' + ('-' * 80))
    Write-Host ('  {0,-20}{1,-12}{2,12}{3,16}{4,8}' -f 'SESSION START', 'ID', 'COST', 'TOKENS', '%')
    Write-Host ('  ' + ('-' * 80))
    foreach ($r in $rows) {
        $pct = 0.0; if ($sc -gt 0) { $pct = $r.Cost / $sc * 100 }
        $local = $r.First.ToLocalTime()
        Write-Host ('  {0:yyyy-MM-dd HH:mm}   {1,-12}{2,12}{3,16}{4,7:N1}%' -f $local, $r.Sid.Substring(0, [Math]::Min(8, $r.Sid.Length)), ('$' + ('{0:N2}' -f $r.Cost)), ('{0:N0}' -f $r.Tokens), $pct)
    }
    Write-Host ('  ' + ('-' * 80))
    Write-Host ('  {0,-32}{1,12}{2,16}' -f ("TOTAL ($($rows.Count) sessions)"), ('$' + ('{0:N2}' -f $sc)), ('{0:N0}' -f $st))
    Write-Host ''
}

if ($unpriced.Count -gt 0) {
    Write-Host '  [note] unpriced model turns skipped (no billable API rate):'
    foreach ($m in $unpriced.Keys) { Write-Host ("    {0}: {1} turns" -f $m, $unpriced[$m]) }
    Write-Host ''
}
if ($Command -eq 'summary') {
    Write-Host '  * Bucket classification defaults every initiative to Dev; real'
    Write-Host '    cc-classify uses config.toml rules not available here. Edit the'
    Write-Host '    $BucketMap table in this script to reclassify.'
    Write-Host ''
}
