@echo off
setlocal
set "DR_HOME=%~dp0"
set "DR_BAT=%~f0"
set "DR_FILE=%TEMP%\DashboardReports-%RANDOM%.ps1"
more +12 "%~f0" > "%DR_FILE%"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%DR_FILE%" %*
set "DR_ERR=%errorlevel%"
del /f /q "%DR_FILE%" >nul 2>&1
if not "%DR_ERR%"=="0" pause
exit /b %DR_ERR%
:: ===== PowerShell part below (copied out by the lines above) =====
# =====================================================================
#  Dashboard Reports - PowerShell part of the single-file
#  Dashboard-Reports.bat. The batch lines above copy everything from
#  here down into a randomly-named .ps1 in %TEMP% (deleted again when the
#  script exits) and run it with -File - same trick as the toolkit files
#  in ref/.
#
#  Runs the 4 monthly report scripts, UNCHANGED:
#     1 Login_time.py         2 report_pipeline.py
#     3 clean_ready_time.py   4 clean_engage_time.py
#  They are looked for next to this .bat, then in a Utiliz\ sub-folder,
#  then in utils\. The .bat can be started from anywhere (double-click, a
#  shortcut, or a prompt in another folder): it always looks next to
#  ITSELF, never in the current folder.
#  Before anything is written it shows a plan (what is read, what is
#  overwritten) and waits for Y. Nothing is run without that.
#  Every folder is worked out from the chosen month. Every default is in
#  the DEFAULTS block just below, and can be overridden for one session
#  from the menu (Paths) or from the command line.
#  Run with -h for About + Parameters.
# =====================================================================
#  VERSION HISTORY
#  1.0  2026-11-02  FIRST RELEASE. Runs the scripts for real: plan screen,
#                    Y to confirm, per-script PASS/FAIL, summary, log file.
#                    Skips (and says why) a step whose input is missing or
#                    whose output file is open in Excel. Warns when
#                    cleaned\ holds older files (they are left out of the
#                    master). Step 2 runs in a fresh temp folder and copies
#                    the cleaned files back, because report_pipeline.py
#                    cannot re-run a month whose cleaned\ files exist
#                    (Windows rename error on every day). A step that
#                    finishes but skipped some daily files is reported as
#                    WARN, never as PASS. Ready/Engage are skipped when the
#                    raw file names are only for another month (Ready would
#                    otherwise overwrite THAT month's master).
#                    V = preview only. CLI: run ... -Yes runs, no -Yes
#                    previews. Hides the openpyxl warnings.
#  0.2.1 2026-11-02 Lives in the Utiliz folder next to the 4 scripts. Looks
#                    for the scripts next to the .bat first, then in utils\.
#  0.2  2026-11-02  Before a run, lists every MISSING folder (inputs and
#                    outputs of the ticked scripts) and asks Enter/Y to
#                    create them. New key C on the main screen does the
#                    same on demand; -CreateMissing does it from the
#                    command line without asking.
#  0.1  2026-11-02  First experimental build. DRY-RUN only: it works out
#                    the folders and prints the exact python commands, but
#                    runs nothing yet.
# =====================================================================

param(
    [Parameter(Position = 0)] [string]$Action  = '',
    [Parameter(Position = 1)] [string]$Scripts = '',

    [string]$Month = '',
    [string]$Box   = '',

    # Custom folders - each one overrides just that one folder
    [string]$LoginIn     = '',
    [string]$LoginOut    = '',
    [string]$NotReadyIn  = '',
    [string]$NotReadyOut = '',
    [string]$UtilIn      = '',
    [string]$ReadyOut    = '',
    [string]$EngageOut   = '',

    # Command-line mode: create missing folders without asking
    [switch]$CreateMissing,

    # Command-line mode: really run. Without -Yes, "run" only previews.
    [switch]$Yes,

    [Alias('h')]
    [switch]$Help
)

# =====================================================================
#  DEFAULTS - this is the block to edit (menu key E opens this file)
#
#  Folders are relative to the Box folder and may use these tokens:
#     {Mon}    short month name from MonthAbbr      Oct   Sept
#     {Month}  full month name                      October
#     {Year}   4-digit year                         2026
#     {YY}     2-digit year                         26
#     {MM}     2-digit month number                 10
#  Any C:\Users\<name>\ in a full path is re-pointed to whoever is
#  signed in, so the same file works on every PC with the same Box layout.
# =====================================================================
$Defaults = [ordered]@{
    # ----- general -----
    Box       = ''                 # '' = auto: <your profile>\Box
    Month     = 'last'             # last | this | YYYY-MM
    Selected  = '1,2,3,4'          # ticked when the menu opens (1-4, or all)
    MonthAbbr = 'Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sept,Oct,Nov,Dec'

    # ----- folders -----
    LoginIn     = 'Agent Not Ready 2025\{Year}\{Mon} Login-LogOut {Year}'
    LoginOut    = 'Not Ready tracker\Login Logout'
    NotReadyIn  = 'Agent Not Ready 2025\{Year}\{Mon} {Year}'
    NotReadyOut = 'Not Ready tracker\Not Ready\MASTER_{Month}.xlsx'
    UtilIn      = 'Utilization (Tracker)'
    ReadyOut    = 'Utilization (Tracker)\Agent Summary Activity Report (Active)'
    EngageOut   = 'Utilization (Tracker)\Agent Utilization Report (Total Enage Time)'

    # ----- file names the scripts are expected to write -----
    # (the scripts choose these themselves from the data; they are only
    #  used here to tell you "exists - would be overwritten")
    LoginFile   = 'Master_{Month}_Login.xlsx'
    ReadyFile   = 'MASTER_{Month}_{Year}_Ready.xlsx'
    EngageFile  = 'MASTER_{Month}_{Year}_Engage.xlsx'
}

$Script:DR_VERSION = '1.0'
$ErrorActionPreference = 'Stop'
$Script:Interactive = $false
$Script:RunOk = $true
$Script:NrTemp = $null
try { $Host.UI.RawUI.WindowTitle = "Dashboard Reports v$DR_VERSION" } catch {}

$Script:AppDir = $env:DR_HOME
if (-not $Script:AppDir) {
    if ($PSCommandPath) { $Script:AppDir = Split-Path -Parent $PSCommandPath } else { $Script:AppDir = (Get-Location).Path }
}
$Script:AppDir = $Script:AppDir.TrimEnd('\')
# Where the 4 .py scripts are: next to this .bat, or in a Utiliz\ or utils\
# sub-folder. The first one holding Login_time.py wins.
$Script:Utils = Join-Path $Script:AppDir 'Utiliz'
foreach ($cand in @($Script:AppDir, (Join-Path $Script:AppDir 'Utiliz'), (Join-Path $Script:AppDir 'utils'))) {
    if (Test-Path -LiteralPath (Join-Path $cand 'Login_time.py')) { $Script:Utils = $cand; break }
}
$Script:UserDir  = $env:USERPROFILE
$Script:UserName = $env:USERNAME

$Script:MonthNames = @('January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December')
$Script:Abbr3      = @('jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec')
$Script:Abbr = @($Defaults.MonthAbbr -split '\s*,\s*' | Where-Object { $_ })
if ($Script:Abbr.Count -ne 12) { $Script:Abbr = @('Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sept', 'Oct', 'Nov', 'Dec') }
$Script:LevelColor = @{ ok = 'Green'; warn = 'Yellow'; bad = 'Red' }

# Removes the temp .ps1 copy this run came from, plus leftovers from earlier
# runs (random-named copies left behind when a window is closed with the X
# button, which skips the batch "del" line).
function Remove-TempCopies {
    try {
        if ($PSCommandPath -and (Split-Path -Leaf $PSCommandPath) -like 'DashboardReports-*.ps1') {
            Remove-Item -LiteralPath $PSCommandPath -Force -ErrorAction SilentlyContinue
        }
    } catch {}
    try {
        if ($env:TEMP) {
            Get-ChildItem -LiteralPath $env:TEMP -Filter 'DashboardReports-*.ps1' -File -ErrorAction SilentlyContinue |
                Remove-Item -Force -ErrorAction SilentlyContinue
            # temp folders of step 2 left behind by a crashed run (older than a day only)
            Get-ChildItem -LiteralPath $env:TEMP -Directory -Filter 'DashboardReports-cleaned-*' -ErrorAction SilentlyContinue |
                Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-1) } |
                Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        }
    } catch {}
}
Remove-TempCopies

# =====================================================================
#  GENERAL HELPERS (same look and keys as the ref/ toolkit files)
# =====================================================================
function W {
    param([string]$Text = '', [string]$Color)
    if ($Color) { Write-Host $Text -ForegroundColor $Color } else { Write-Host $Text }
}

function Write-Part {
    param([string]$Text = '', [string]$Color)
    if ($Color) { Write-Host -NoNewline $Text -ForegroundColor $Color } else { Write-Host -NoNewline $Text }
}

function Read-Key {
    try { while ([Console]::KeyAvailable) { [void][Console]::ReadKey($true) } } catch {}
    return [Console]::ReadKey($true)
}

# Normalises a key press to one upper-case character. Matches on KeyChar AND
# the ConsoleKey enum name, because KeyChar alone can come back empty/wrong
# on some console window types (see Network-Scan.bat).
function Get-KeyChar($k) {
    $c = [string]$k.KeyChar
    if ($c -and -not [char]::IsControl($k.KeyChar)) { return $c.ToUpper() }
    $n = $k.Key.ToString()
    if ($n -match '^D([0-9])$') { return $Matches[1] }
    if ($n -match '^NumPad([0-9])$') { return $Matches[1] }
    if ($n -eq 'OemMinus' -or $n -eq 'Subtract') { return '-' }
    if ($n -eq 'OemPlus' -or $n -eq 'Add') { return '+' }
    if ($n.Length -eq 1) { return $n.ToUpper() }
    return ''
}

function Read-Input {
    param([string]$Prompt, [string]$Default = '')
    Write-Host -NoNewline "  $Prompt"
    if ($Default -ne '') { Write-Host -NoNewline " [$Default]" -ForegroundColor DarkGray }
    Write-Host -NoNewline ': '
    $buf = New-Object System.Text.StringBuilder
    while ($true) {
        $k = [Console]::ReadKey($true)
        if ($k.Key -eq 'Escape') { W ''; return $null }
        if ($k.Key -eq 'Enter') {
            W ''
            if ($buf.Length -eq 0) { return $Default }
            return $buf.ToString().Trim()
        }
        if ($k.Key -eq 'Backspace') {
            if ($buf.Length -gt 0) { [void]$buf.Remove($buf.Length - 1, 1); Write-Host -NoNewline "`b `b" }
            continue
        }
        if (-not [char]::IsControl($k.KeyChar)) { [void]$buf.Append($k.KeyChar); Write-Host -NoNewline $k.KeyChar }
    }
}

function Show-Menu {
    param([string]$Title, [string[]]$Items, [string]$EscLabel = 'Back', [string]$Default = '', [scriptblock]$Info)
    while ($true) {
        Clear-Host
        W ''
        W "  $Title" Cyan
        W ('  ' + ('=' * $Title.Length)) Cyan
        if ($Info) { try { & $Info } catch {} }
        W ''
        foreach ($it in $Items) {
            $p = $it.Split('|', 2)
            $tag = if ($Default -and $p[0] -eq $Default) { '   (Enter)' } else { '' }
            W ("  [{0}]  {1}{2}" -f $p[0], $p[1], $tag)
        }
        W "  [Esc] $EscLabel" DarkGray
        W ''
        Write-Host -NoNewline '  Choose: '
        $k = Read-Key
        if ($k.Key -eq 'Escape') { W ''; return $null }
        if ($k.Key -eq 'Enter' -and $Default) { W $Default; return $Default }
        $ch = Get-KeyChar $k
        foreach ($it in $Items) {
            $tag = $it.Split('|', 2)[0]
            if ($ch -and $tag -eq $ch) { W $tag; return $tag }
        }
    }
}

function Wait-Back {
    param([string]$Text = 'Press any key to go back')
    W ''
    Write-Host -NoNewline "  $Text..." -ForegroundColor DarkGray
    [void](Read-Key)
    W ''
}

function Show-ExtraSteps {
    param([string[]]$Steps)
    W ''
    if (-not $Steps -or $Steps.Count -eq 0) { W 'Extra steps needed: none.' Cyan; return }
    W 'Extra steps needed:' Yellow
    $i = 1
    foreach ($s in $Steps) { W "  $i. $s"; $i++ }
}

function Invoke-Safe {
    param([scriptblock]$Block)
    try { & $Block } catch { W ''; W "ERROR: $($_.Exception.Message)" Red; $Script:RunOk = $false }
}

function Format-Arg([string]$s) {
    if ($s -match '[\s"]' -or $s -match '\\') { return '"' + ($s -replace '"', '\"') + '"' }
    return $s
}

function Format-Command([string]$ScriptPath, [string[]]$ArgList) {
    $parts = @('python', (Format-Arg $ScriptPath))
    foreach ($a in $ArgList) { $parts += (Format-Arg $a) }
    return ($parts -join ' ')
}

# =====================================================================
#  USER / BOX PATHS
# =====================================================================
# Re-points a leading C:\Users\<anyone>\ to the signed-in user, so a path
# typed (or saved in the DEFAULTS) for another PC still works on this one.
function Convert-UserPath([string]$p) {
    if ([string]::IsNullOrWhiteSpace($p)) { return $p }
    $p = [Environment]::ExpandEnvironmentVariables($p.Trim().Trim('"'))
    if ($p -match '^[A-Za-z]:\\Users\\[^\\]+(\\.*)?$') {
        $rest = ''
        if ($Matches[1]) { $rest = $Matches[1] }
        return ($Script:UserDir + $rest)
    }
    return $p
}

function Resolve-BoxRoot {
    $b = $Box
    if (-not $b) { $b = $Defaults.Box }
    if ($b) { return (Convert-UserPath $b) }
    return (Join-Path $Script:UserDir 'Box')
}

# =====================================================================
#  MONTH HANDLING
# =====================================================================
function Resolve-Month([string]$s) {
    $now = Get-Date
    $t = ([string]$s).Trim().ToLower()
    if ($t -eq '' -or $t -eq 'last') {
        $d = (New-Object DateTime $now.Year, $now.Month, 1).AddMonths(-1)
        return [pscustomobject]@{ Year = $d.Year; Num = $d.Month }
    }
    if ($t -eq 'this') { return [pscustomobject]@{ Year = $now.Year; Num = $now.Month } }
    if ($t -match '^(\d{4})[-/](\d{1,2})$') {
        $y = [int]$Matches[1]; $n = [int]$Matches[2]
        if ($n -ge 1 -and $n -le 12) { return [pscustomobject]@{ Year = $y; Num = $n } }
        return $null
    }
    $n = 0; $y = 0
    if ($t -match '^(\d{1,2})$') {
        $n = [int]$Matches[1]
    } elseif ($t -match '^([a-z]{3,9})\.?\s*(\d{4})?$') {
        $name = $Matches[1]
        if ($Matches[2]) { $y = [int]$Matches[2] }
        for ($i = 1; $i -le 12; $i++) {
            if ($Script:MonthNames[$i - 1].ToLower().StartsWith($name)) { $n = $i; break }
        }
    }
    if ($n -lt 1 -or $n -gt 12) { return $null }
    # No year given: the most recent past-or-current occurrence of that month
    if ($y -eq 0) { $y = $now.Year; if ($n -gt $now.Month) { $y = $y - 1 } }
    return [pscustomobject]@{ Year = $y; Num = $n }
}

function Get-YM { return ('{0}-{1:00}' -f $Script:M.Year, $Script:M.Num) }

function Get-MonthLabel {
    return ('{0} {1}  ({2})' -f $Script:MonthNames[$Script:M.Num - 1], $Script:M.Year, (Get-YM))
}

function Expand-Pattern([string]$pat, $m, [string]$MonOverride = '') {
    $mon = $MonOverride
    if (-not $mon) { $mon = $Script:Abbr[$m.Num - 1] }
    return $pat.Replace('{Mon}', $mon).Replace('{Month}', $Script:MonthNames[$m.Num - 1]).Replace('{Year}', [string]$m.Year).Replace('{YY}', ('{0:00}' -f ($m.Year % 100))).Replace('{MM}', ('{0:00}' -f $m.Num))
}

# Spellings of the month to try when the default folder name is not there:
# the configured short name first, then the full name, 3 letters, and Sept.
function Get-MonCandidates($m) {
    $full = $Script:MonthNames[$m.Num - 1]
    $list = @($Script:Abbr[$m.Num - 1], $full, $full.Substring(0, 3))
    if ($m.Num -eq 9) { $list += 'Sept' }
    return @($list | Select-Object -Unique)
}

# =====================================================================
#  FOLDER SLOTS - the 7 folders the 4 scripts use
# =====================================================================
$Script:Slots = [ordered]@{}
$slotSpec = @(
    @('LoginIn',     'Login files (input)',        'dir'),
    @('LoginOut',    'Login master (output dir)',  'dir'),
    @('NotReadyIn',  'Not Ready files (input)',    'dir'),
    @('NotReadyOut', 'Not Ready master (file)',    'file'),
    @('UtilIn',      'Utilization raw (input)',    'dir'),
    @('ReadyOut',    'Ready master (output dir)',  'dir'),
    @('EngageOut',   'Engage master (output dir)', 'dir')
)
$cli = @{
    LoginIn = $LoginIn; LoginOut = $LoginOut; NotReadyIn = $NotReadyIn; NotReadyOut = $NotReadyOut
    UtilIn = $UtilIn; ReadyOut = $ReadyOut; EngageOut = $EngageOut
}
foreach ($s in $slotSpec) {
    $custom = $null
    if ($cli[$s[0]]) { $custom = $cli[$s[0]] }
    $Script:Slots[$s[0]] = [pscustomobject]@{
        Key = $s[0]; Label = $s[1]; Kind = $s[2]; Pattern = $Defaults[$s[0]]; Custom = $custom
    }
}

function Test-SlotPath($slot, [string]$p) {
    if ($slot.Kind -eq 'file') { return (Test-Path -LiteralPath $p -PathType Leaf) }
    return (Test-Path -LiteralPath $p -PathType Container)
}

# Resolves one slot to a path. Order: custom path > default name > other
# spellings of the month that actually exist on disk.
function Get-SlotInfo($slot) {
    if ($slot.Custom) {
        $p = Convert-UserPath $slot.Custom
        return [pscustomobject]@{ Path = $p; Source = 'custom'; Exists = (Test-SlotPath $slot $p) }
    }
    if ($slot.Pattern -match '\{Mon\}') {
        $first = $true
        foreach ($c in (Get-MonCandidates $Script:M)) {
            $p = Join-Path $Script:BoxRoot (Expand-Pattern $slot.Pattern $Script:M $c)
            if (Test-SlotPath $slot $p) {
                $src = 'default'
                if (-not $first) { $src = 'alt name' }
                return [pscustomobject]@{ Path = $p; Source = $src; Exists = $true }
            }
            $first = $false
        }
    }
    $p = Join-Path $Script:BoxRoot (Expand-Pattern $slot.Pattern $Script:M)
    return [pscustomobject]@{ Path = $p; Source = 'default'; Exists = (Test-SlotPath $slot $p) }
}

function Get-SlotPath([string]$key) { return (Get-SlotInfo $Script:Slots[$key]).Path }

function Get-ShortPath([string]$p) {
    $b = $Script:BoxRoot.TrimEnd('\')
    if ($p.StartsWith($b + '\', [StringComparison]::OrdinalIgnoreCase)) { return '[Box]\' + $p.Substring($b.Length + 1) }
    return $p
}

# =====================================================================
#  THE 4 STEPS
# =====================================================================
$Script:Steps = @(
    [pscustomobject]@{ Id = 1; Name = 'Login time';         Py = 'Login_time.py';         In = 'LoginIn';    Out = 'LoginOut';  Kind = 'daily' },
    [pscustomobject]@{ Id = 2; Name = 'Not Ready reasons';  Py = 'report_pipeline.py';    In = 'NotReadyIn'; Out = 'NotReadyOut'; Kind = 'daily' },
    [pscustomobject]@{ Id = 3; Name = 'Ready time';         Py = 'clean_ready_time.py';   In = 'UtilIn';     Out = 'ReadyOut';  Kind = 'raw' },
    [pscustomobject]@{ Id = 4; Name = 'Engage, Hold, Wrap'; Py = 'clean_engage_time.py';  In = 'UtilIn';     Out = 'EngageOut'; Kind = 'raw' }
)

# Step 2 (report_pipeline.py) is run with --outdir pointing at a FRESH temporary
# folder. Reason: that script renames each cleaned file to
# "<DD> <Month> <Year>_clean.xlsx", and on Windows the rename FAILS when a file
# with that name already exists - so a second run of the same month skipped
# every single day. A fresh folder also means the master is built from this
# run's files only. The cleaned files are copied back into <month folder>\cleaned
# afterwards (overwriting), exactly where a first run would have put them.
function Get-NrTemp {
    if (-not $Script:NrTemp) {
        $Script:NrTemp = Join-Path $env:TEMP ('DashboardReports-cleaned-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    }
    return $Script:NrTemp
}

function Get-StepArgs([int]$id) {
    switch ($id) {
        1 { return @((Get-SlotPath 'LoginIn'), (Get-SlotPath 'LoginOut')) }
        2 { return @('--folder', (Get-SlotPath 'NotReadyIn'), '--output', (Get-SlotPath 'NotReadyOut'), '--year', [string]$Script:M.Year, '--outdir', (Get-NrTemp)) }
        3 { return @('--folder', (Get-SlotPath 'UtilIn'), '--outdir', (Get-SlotPath 'ReadyOut')) }
        4 { return @('--folder', (Get-SlotPath 'UtilIn'), '--outdir', (Get-SlotPath 'EngageOut'), '--month', (Get-YM)) }
    }
}

function Get-ExpectedOutput([int]$id) {
    $dir = ''; $name = ''
    switch ($id) {
        1 { $dir = Get-SlotPath 'LoginOut';  $name = $Defaults.LoginFile }
        2 { return (Get-SlotPath 'NotReadyOut') }
        3 { $dir = Get-SlotPath 'ReadyOut';  $name = $Defaults.ReadyFile }
        4 { $dir = Get-SlotPath 'EngageOut'; $name = $Defaults.EngageFile }
    }
    return (Join-Path $dir (Expand-Pattern $name $Script:M))
}

function Get-XlsxFiles([string]$dir) {
    if (-not (Test-Path -LiteralPath $dir -PathType Container)) { return @() }
    return @(Get-ChildItem -LiteralPath $dir -File -Filter '*.xlsx' -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike '~$*' })
}

# Looks at a step's input folder and says whether it looks ready to run.
function Get-StepStatus($st) {
    $slot = Get-SlotInfo $Script:Slots[$st.In]
    $leaf = Split-Path -Leaf $slot.Path
    if (-not $slot.Exists) {
        return [pscustomobject]@{ Level = 'bad'; Text = 'folder not found'; Leaf = $leaf; Info = $slot }
    }
    $files = Get-XlsxFiles $slot.Path
    $n = $files.Count
    if ($n -eq 0) {
        return [pscustomobject]@{ Level = 'bad'; Text = 'no .xlsx files in it'; Leaf = $leaf; Info = $slot }
    }
    $note = ''
    if ($slot.Source -ne 'default') { $note = "  ($($slot.Source))" }

    if ($st.Kind -eq 'daily') {
        $dim = [DateTime]::DaysInMonth($Script:M.Year, $Script:M.Num)
        $lvl = 'ok'
        if ($n -lt $dim) { $lvl = 'warn' }
        return [pscustomobject]@{ Level = $lvl; Text = ("{0} of {1} days{2}" -f $n, $dim, $note); Leaf = $leaf; Info = $slot }
    }

    # Raw monthly exports: the scripts do not look at the file name, so warn
    # if the names mention a different month than the one selected.
    #  - names are ONLY for other months -> 'bad' (the step is skipped): the
    #    Ready script has no month filter, so it would write that other month's
    #    master (overwriting it), and Engage would crash with nothing left.
    #  - names for this month AND others -> 'warn': Ready would mix the months.
    $others = @()
    $match = 0
    foreach ($f in $files) {
        if ($f.Name -match '^\s*(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)') {
            $idx = [array]::IndexOf($Script:Abbr3, $Matches[1].ToLower()) + 1
            if ($idx -ne $Script:M.Num) { $others += $idx } else { $match++ }
        }
    }
    $others = @($others | Select-Object -Unique)
    if ($others.Count -gt 0) {
        $said = ($others | ForEach-Object { $Script:MonthNames[$_ - 1].Substring(0, 3) }) -join '/'
        $want = $Script:MonthNames[$Script:M.Num - 1].Substring(0, 3)
        if ($match -eq 0) {
            return [pscustomobject]@{ Level = 'bad'; Text = ("{0} xlsx - files are for {1}, not {2}: put the {2} files in, or pick {1}{3}" -f $n, $said, $want, $note); Leaf = $leaf; Info = $slot }
        }
        return [pscustomobject]@{ Level = 'warn'; Text = ("{0} xlsx - {1} AND {2} files mixed: move the {2} ones out (Ready has no month filter){3}" -f $n, $want, $said, $note); Leaf = $leaf; Info = $slot }
    }
    return [pscustomobject]@{ Level = 'ok'; Text = ("{0} xlsx file(s){1}" -f $n, $note); Leaf = $leaf; Info = $slot }
}

# =====================================================================
#  SELECTION (ticks)
# =====================================================================
function Get-DefaultIds {
    $t = ([string]$Defaults.Selected).Trim().ToLower()
    if ($t -ne '' -and $t -ne 'all') {
        $ids = @([regex]::Matches($t, '[1-4]') | ForEach-Object { [int]$_.Value } | Select-Object -Unique | Sort-Object)
        if ($ids.Count -gt 0) { return $ids }
    }
    return @(1, 2, 3, 4)
}

$Script:Sel = @{}
function Set-DefaultTicks {
    $d = Get-DefaultIds
    foreach ($id in 1..4) { $Script:Sel[$id] = ($d -contains $id) }
}
function Get-TickedIds { return @(1..4 | Where-Object { $Script:Sel[$_] }) }

function Resolve-Ids([string]$s) {
    $t = ([string]$s).Trim().ToLower()
    if ($t -eq '') { return @(Get-DefaultIds) }
    if ($t -eq 'all') { return @(1, 2, 3, 4) }
    if ($t -match '^[1-4,;+\s]+$') {
        return @([regex]::Matches($t, '[1-4]') | ForEach-Object { [int]$_.Value } | Select-Object -Unique | Sort-Object)
    }
    return $null
}

# =====================================================================
#  SCREENS
# =====================================================================
# =====================================================================
#  MISSING FOLDERS - list them all, then create only after a yes
# =====================================================================
# Every folder the given steps need (their input and their output) that does
# not exist yet. A "file" slot (the Not Ready master) needs its parent folder.
function Get-MissingFolders([int[]]$Ids) {
    $seen = @{}
    $list = @()
    foreach ($id in $Ids) {
        $st = $Script:Steps[$id - 1]
        foreach ($pair in @(@($st.In, 'input'), @($st.Out, 'output'))) {
            $slot = $Script:Slots[$pair[0]]
            $p = (Get-SlotInfo $slot).Path
            if ($slot.Kind -eq 'file') { $p = Split-Path -Parent $p }
            $k = $p.ToLower()
            if ($seen.ContainsKey($k)) { continue }
            $seen[$k] = $true
            if (-not (Test-Path -LiteralPath $p -PathType Container)) {
                $list += [pscustomobject]@{ Path = $p; Role = $pair[1]; Step = $st.Name }
            }
        }
    }
    return $list
}

# Lists the missing folders and creates them if the user says yes (or
# $AutoCreate is set). Returns how many were created.
function Confirm-MissingFolders([int[]]$Ids, [bool]$AutoCreate = $false) {
    $missing = @(Get-MissingFolders $Ids)
    if ($missing.Count -eq 0) { return 0 }
    W ''
    W ('Missing folders ({0}):' -f $missing.Count) Yellow
    foreach ($m in $missing) {
        W ('  [{0,-6}] {1}' -f $m.Role, $m.Path)
        W ('           for: {0}' -f $m.Step) DarkGray
    }
    if (@($missing | Where-Object Role -eq 'input').Count -gt 0) {
        W '  (a new input folder is created empty - you still have to put the files in it)' DarkGray
    }

    $go = $AutoCreate
    if (-not $go) {
        if ($Script:Interactive) {
            W ''
            Write-Host -NoNewline '  Create these folders?  Enter/Y = create, Esc / any other key = skip: '
            $k = Read-Key
            W ''
            $go = ($k.Key -eq 'Enter' -or (Get-KeyChar $k) -eq 'Y')
        } else {
            W '  Not created. Add -CreateMissing to create them from the command line.' DarkGray
        }
    }
    if (-not $go) { if ($Script:Interactive) { W '  Skipped - nothing was created.' DarkGray }; return 0 }

    $made = 0
    foreach ($m in $missing) {
        try {
            [void](New-Item -ItemType Directory -Force -Path $m.Path)
            W ('  created: {0}' -f $m.Path) Green
            $made++
        } catch {
            W ('  FAILED : {0}  ({1})' -f $m.Path, $_.Exception.Message) Red
        }
    }
    return $made
}

function Write-StepRows {
    foreach ($st in $Script:Steps) {
        $on = [bool]$Script:Sel[$st.Id]
        $sx = Get-StepStatus $st
        $dim = ''
        if (-not $on) { $dim = 'DarkGray' }
        $mark = '[ ]'
        if ($on) { $mark = '[x]' }
        $col = 'DarkGray'
        if ($on) { $col = $Script:LevelColor[$sx.Level] }
        Write-Part ('  [{0}] {1} {2,-20}' -f $st.Id, $mark, $st.Name) $dim
        Write-Part ('{0,-24}' -f $sx.Leaf) $dim
        Write-Part $sx.Text $col
        W ''
    }
}

function Write-MainScreen {
    $title = "Dashboard Reports v$DR_VERSION"
    W ''
    W "  $title" Cyan
    W ('  ' + ('=' * $title.Length)) Cyan
    W '  Shows a plan first and asks Y before it runs or overwrites anything.' DarkGray
    W ('  User : {0}        Box: {1}' -f $Script:UserName, $Script:BoxRoot) DarkGray
    if (-not (Test-Path -LiteralPath $Script:BoxRoot -PathType Container)) {
        W '  Box folder not found - press P, then B, to point at it.' Red
    }
    Write-Part '  Month: '
    W (Get-MonthLabel) Cyan
    W ''
    W '  Keys: 1-4 tick/untick   T all/none   D reset ticks to default' DarkGray
    W '        Enter run ticked   A run ALL   V preview only (runs nothing)' DarkGray
    W '        M month   + / - next/prev month   P paths   C create missing folders' DarkGray
    W '        E edit defaults   H help   Esc exit' DarkGray
    W ''
    Write-StepRows
    W ''
    $ids = Get-TickedIds
    W ('  Ticked: {0} of {1}      Default ticks: {2}' -f $ids.Count, $Script:Steps.Count, $Defaults.Selected)
    W ''
}

function Show-Summary {
    W ''
    W ('User : {0}' -f $Script:UserName) DarkGray
    W ('Box  : {0}' -f $Script:BoxRoot) DarkGray
    Write-Part 'Month: '
    W (Get-MonthLabel) Cyan
    W ''
    Write-StepRows
    W ''
    W 'Folders' Yellow
    foreach ($key in $Script:Slots.Keys) {
        $s = $Script:Slots[$key]
        $i = Get-SlotInfo $s
        $tag = ''
        if ($i.Source -ne 'default') { $tag = "  ($($i.Source))" }
        $col = 'Green'
        if (-not $i.Exists) { $col = 'DarkGray' }
        W ('  {0,-12} {1}{2}' -f $key, $i.Path, $tag) $col
    }
    W '  (grey = does not exist yet; output files/folders are created by the scripts)' DarkGray
}

# =====================================================================
#  RUNNING - plan, confirm, run, summary
# =====================================================================
# True if another program (Excel, a sync client) holds the file open.
function Test-FileLocked([string]$p) {
    if (-not (Test-Path -LiteralPath $p -PathType Leaf)) { return $false }
    try {
        $fs = [IO.File]::Open($p, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        $fs.Close()
        return $false
    } catch { return $true }
}

# .xlsx files in an output folder written at or after $since.
function Get-NewOutputs([string]$dir, [datetime]$since) {
    if (-not (Test-Path -LiteralPath $dir -PathType Container)) { return @() }
    return @(Get-ChildItem -LiteralPath $dir -File -Filter '*.xlsx' -ErrorAction SilentlyContinue |
             Where-Object { $_.LastWriteTime -ge $since -and $_.Name -notlike '~$*' })
}

function Get-StepOutDir([int]$id) {
    $key = $Script:Steps[$id - 1].Out
    $p = Get-SlotPath $key
    if ($Script:Slots[$key].Kind -eq 'file') { return (Split-Path -Parent $p) }
    return $p
}

# The Not Ready script writes <input>\cleaned\<DD> <Month> <Year>_clean.xlsx and
# then combines EVERY *_clean.xlsx found in that folder into the master.
function Get-CleanedInfo {
    $dir = Join-Path (Get-SlotPath 'NotReadyIn') 'cleaned'
    $info = [pscustomobject]@{ Dir = $dir; Exists = (Test-Path -LiteralPath $dir -PathType Container); Total = 0; Other = 0 }
    if (-not $info.Exists) { return $info }
    $pat = '^\d{2} ' + $Script:MonthNames[$Script:M.Num - 1] + ' ' + $Script:M.Year + '_clean\.xlsx$'
    $files = @(Get-ChildItem -LiteralPath $dir -File -Filter '*_clean.xlsx' -ErrorAction SilentlyContinue)
    $info.Total = $files.Count
    $info.Other = @($files | Where-Object { $_.Name -notmatch $pat }).Count
    return $info
}

# Preview ($Execute = $false): prints the plan and the exact commands, runs nothing.
# Run ($Execute = $true): prints the plan, waits for Enter/Y (unless $AutoYes),
# then runs the scripts one after another and prints a summary.
function Invoke-Steps([int[]]$Ids, [bool]$Execute = $false, [bool]$AutoYes = $false) {
    $problems = @()
    $warnings = @()
    W ''
    if ($Execute) { W ('PLAN  -  {0}' -f (Get-MonthLabel)) Cyan } else { W ('PREVIEW (nothing will run)  -  {0}' -f (Get-MonthLabel)) Cyan }
    W ('Box  : {0}' -f $Script:BoxRoot) DarkGray

    $py = Get-Command python -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $py) {
        W 'Python: not found on PATH' Red
        $problems += 'Install Python or add it to PATH (python --version must work in a new Command Prompt).'
    } elseif ($py.Source -like '*\WindowsApps\*') {
        W ("Python: {0}  (Microsoft Store shortcut - may not be a real Python)" -f $py.Source) Yellow
        $warnings += 'python resolves to the Microsoft Store shortcut. If the run fails, install Python or put it before WindowsApps in PATH.'
    } else {
        W ("Python: {0}" -f $py.Source) DarkGray
    }

    # ----- build the plan -----
    $plan = @()
    $i = 0
    foreach ($id in $Ids) {
        $i++
        $st = $Script:Steps[$id - 1]
        $sx = Get-StepStatus $st
        $scriptPath = Join-Path $Script:Utils $st.Py
        $out = Get-ExpectedOutput $id
        $skip = ''
        if (-not (Test-Path -LiteralPath $scriptPath)) { $skip = "script not found: $scriptPath" }
        elseif ($sx.Level -eq 'bad') { $skip = "input problem - $($sx.Text)" }
        elseif (Test-FileLocked $out) { $skip = "output file is open in another program (Excel?) - close it and run again: $out" }

        W ''
        W ('[{0}/{1}]  {2}   ({3})' -f $i, $Ids.Count, $st.Name, $st.Py) Cyan
        W ("  Input : {0}   [{1}]" -f $sx.Info.Path, $sx.Text) $Script:LevelColor[$sx.Level]
        $state = 'new file'
        if (Test-Path -LiteralPath $out) { $state = 'EXISTS - will be overwritten' }
        W ("  Output: {0}   [{1}]" -f $out, $state)

        if ($id -eq 2 -and -not $skip) {
            $ci = Get-CleanedInfo
            if (-not $ci.Exists) {
                W ("  Also  : the cleaned daily files go to {0}  (new folder)" -f $ci.Dir) DarkGray
            } else {
                W ("  Also  : the {0} cleaned daily file(s) in {1} are refreshed (overwritten)" -f $ci.Total, $ci.Dir) DarkGray
                if ($ci.Other -gt 0) {
                    W ("          {0} other file(s) there (older names / other months) stay as they are and are NOT part of the master" -f $ci.Other) DarkGray
                }
            }
            W '          (built in a fresh temp folder first, so running the same month again works)' DarkGray
        }
        if ($skip) {
            W "  SKIP  : $skip" Red
            $problems += "Step $id ($($st.Name)) will be skipped: $skip"
        } else {
            if ($sx.Level -eq 'warn') { $warnings += "Step $id ($($st.Name)): $($sx.Text)" }
            if (-not $Execute) { W ('  ' + (Format-Command $scriptPath @(Get-StepArgs $id))) Cyan }
        }
        $plan += [pscustomobject]@{ Id = $id; Step = $st; Script = $scriptPath; Skip = $skip; Out = $out; Level = $sx.Level }
    }

    $runIds = @($plan | Where-Object { -not $_.Skip } | ForEach-Object { $_.Id })
    $missing = @()
    if ($runIds.Count -gt 0) { $missing = @(Get-MissingFolders $runIds) }

    if ($warnings.Count -gt 0) {
        W ''
        W 'Check before you continue:' Yellow
        foreach ($w in $warnings) { W "  ! $w" Yellow }
    }
    if ($missing.Count -gt 0) {
        W ''
        W ('Missing output folders ({0}) - created when you confirm:' -f $missing.Count) Yellow
        foreach ($m in $missing) { W ('  {0}' -f $m.Path) }
    }

    if ($runIds.Count -eq 0) {
        W ''
        W 'Nothing can run - see the problems above.' Red
        $Script:RunOk = $false
        Show-ExtraSteps $problems
        return
    }

    # ----- preview stops here -----
    if (-not $Execute) {
        if ($CreateMissing.IsPresent -and $missing.Count -gt 0) { [void](Confirm-MissingFolders $runIds $true) }
        W ''
        W 'PREVIEW: nothing was executed.' Yellow
        if ($plan | Where-Object { $_.Skip }) { $Script:RunOk = $false }
        Show-ExtraSteps ($problems + $warnings)
        return
    }
    if (-not $py) {
        W ''
        W 'Cannot run: Python was not found.' Red
        $Script:RunOk = $false
        Show-ExtraSteps $problems
        return
    }

    # ----- confirm -----
    if (-not $AutoYes) {
        W ''
        $q = ('Run {0} of {1} script(s)' -f $runIds.Count, $Ids.Count)
        if ($warnings.Count -gt 0) { $q += ' despite the warnings above' }
        W "  $q?" Yellow
        Write-Host -NoNewline '  Enter/Y = run, Esc / any other key = cancel: '
        $k = Read-Key
        W ''
        if ($k.Key -ne 'Enter' -and (Get-KeyChar $k) -ne 'Y') {
            W '  Cancelled - nothing was run or changed.' Yellow
            return
        }
    }

    # ----- run -----
    $log = Join-Path $env:TEMP 'DashboardReports-last-run.log'
    $prevEap = $ErrorActionPreference
    $prevWarn = $env:PYTHONWARNINGS
    $prevEnc = $env:PYTHONIOENCODING
    $prevOut = $null
    try { $prevOut = [Console]::OutputEncoding; [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch {}
    $ErrorActionPreference = 'Continue'
    $env:PYTHONWARNINGS = 'ignore'
    $env:PYTHONIOENCODING = 'utf-8'
    $results = @()
    try {
        Set-Content -LiteralPath $log -Encoding UTF8 -Value ("Dashboard Reports v{0}  {1}  {2}  user {3}" -f $DR_VERSION, (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), (Get-MonthLabel), $Script:UserName)

        foreach ($m in $missing) {
            try { [void](New-Item -ItemType Directory -Force -Path $m.Path); W ('created: {0}' -f $m.Path) Green }
            catch { W ('FAILED to create {0}: {1}' -f $m.Path, $_.Exception.Message) Red }
        }

        & $py.Source -c 'import pandas, openpyxl' 2>&1 | Out-Null
        if ($LASTEXITCODE -ne 0) {
            W ''
            W 'Python is missing pandas or openpyxl, so nothing was run.' Red
            Show-ExtraSteps @('Run:  python -m pip install pandas openpyxl', 'Then start this again.')
            $Script:RunOk = $false
            return
        }

        $n = 0
        foreach ($p in $plan) {
            $n++
            $st = $p.Step
            W ''
            W ('=== [{0}/{1}] {2}   ({3}) ===' -f $n, $plan.Count, $st.Name, $st.Py) Cyan
            if ($p.Skip) {
                W "  SKIPPED: $($p.Skip)" Red
                $results += [pscustomobject]@{ Id = $p.Id; Name = $st.Name; Status = 'SKIP'; Secs = 0; Files = $p.Skip }
                $Script:RunOk = $false
                continue
            }
            $pyArgs = @($p.Script) + @(Get-StepArgs $p.Id)
            W ('  ' + (Format-Command $p.Script @(Get-StepArgs $p.Id))) DarkGray
            $t0 = Get-Date
            $sw = [Diagnostics.Stopwatch]::StartNew()
            $cnt = @{ hidden = 0; lines = (New-Object System.Collections.ArrayList) }
            & $py.Source @pyArgs 2>&1 | ForEach-Object {
                $txt = "$_"
                [void]$cnt.lines.Add($txt)
                if ($txt -match '^\s*(Cleaned|Loaded|Read)\s*:' -or $txt -match '^\s*\S+\.xlsx:\s+\d+\s+rows') { $cnt.hidden++ }
                elseif ($_ -is [System.Management.Automation.ErrorRecord]) { W "    $txt" Red }
                else { W "    $txt" DarkGray }
            }
            $code = $LASTEXITCODE
            $sw.Stop()
            if ($cnt.hidden -gt 0) { W ("    ({0} per-file progress line(s) hidden - all of it is in the log)" -f $cnt.hidden) DarkGray }
            Add-Content -LiteralPath $log -Encoding UTF8 -Value (@('', ('--- step {0} {1} (exit {2}) ---' -f $p.Id, $st.Name, $code)) + @($cnt.lines))

            $outs = @(Get-NewOutputs (Get-StepOutDir $p.Id) $t0.AddSeconds(-2))
            $names = ($outs | ForEach-Object { $_.Name }) -join ', '

            # The two daily-file steps print one "Skipped" / "[!]" / "WARNING" line
            # per file they could not use. The master is still written, but without
            # that day - so this must not be reported as a clean PASS.
            $bad = @()
            if ($p.Id -le 2) { $bad = @($cnt.lines | Where-Object { $_ -match '^\s*(Skipped\s*:|\[!\]|WARNING:)' }) }

            # Step 2 worked in a temporary folder: put the fresh cleaned files back
            # next to the raw ones, like a first run would have.
            if ($p.Id -eq 2 -and $code -eq 0 -and $outs.Count -gt 0) {
                $tmp = Get-NrTemp
                $dest = Join-Path (Get-SlotPath 'NotReadyIn') 'cleaned'
                if (Test-Path -LiteralPath $tmp -PathType Container) {
                    try {
                        if (-not (Test-Path -LiteralPath $dest -PathType Container)) { [void](New-Item -ItemType Directory -Force -Path $dest) }
                        $copied = 0; $notCopied = @()
                        foreach ($cf in @(Get-ChildItem -LiteralPath $tmp -File -Filter '*_clean.xlsx')) {
                            try { Copy-Item -LiteralPath $cf.FullName -Destination $dest -Force -ErrorAction Stop; $copied++ }
                            catch { $notCopied += $cf.Name }
                        }
                        W ("  cleaned: {0} file(s) refreshed in {1}" -f $copied, $dest) DarkGray
                        if ($notCopied.Count -gt 0) {
                            $bad += ("{0} cleaned file(s) could not be refreshed in cleaned\ (open in Excel?): {1}" -f $notCopied.Count, ($notCopied -join ', '))
                        }
                    } catch { $bad += "could not refresh cleaned\: $($_.Exception.Message)" }
                }
            }

            if ($code -ne 0) {
                W ("  FAIL  (exit code {0}) - see the red lines above, or the log." -f $code) Red
                $results += [pscustomobject]@{ Id = $p.Id; Name = $st.Name; Status = 'FAIL'; Secs = $sw.Elapsed.TotalSeconds; Files = "exit code $code" }
                $Script:RunOk = $false
            } elseif ($outs.Count -eq 0) {
                W '  WARN  finished without an error, but no new output file was found.' Yellow
                $results += [pscustomobject]@{ Id = $p.Id; Name = $st.Name; Status = 'WARN'; Secs = $sw.Elapsed.TotalSeconds; Files = 'no new output file found' }
                $Script:RunOk = $false
            } elseif ($bad.Count -gt 0) {
                foreach ($o in $outs) { W ('  wrote : {0}  ({1:N0} KB)' -f $o.FullName, ($o.Length / 1KB)) Green }
                W ("  WARN  finished, but {0} file(s) were skipped or had problems - the master may be missing days:" -f $bad.Count) Yellow
                foreach ($b in @($bad | Select-Object -First 5)) {
                    $s = ([string]$b).Trim()
                    if ($s.Length -gt 170) { $s = $s.Substring(0, 170) + ' ...' }
                    W "    $s" Yellow
                }
                if ($bad.Count -gt 5) { W ("    ... and {0} more (all in the log)" -f ($bad.Count - 5)) Yellow }
                $results += [pscustomobject]@{ Id = $p.Id; Name = $st.Name; Status = 'WARN'; Secs = $sw.Elapsed.TotalSeconds; Files = ("{0} skipped/problem line(s) - see the log" -f $bad.Count) }
                $Script:RunOk = $false
            } else {
                foreach ($o in $outs) { W ('  wrote : {0}  ({1:N0} KB)' -f $o.FullName, ($o.Length / 1KB)) Green }
                W ('  PASS  ({0:N1} s)' -f $sw.Elapsed.TotalSeconds) Green
                $results += [pscustomobject]@{ Id = $p.Id; Name = $st.Name; Status = 'PASS'; Secs = $sw.Elapsed.TotalSeconds; Files = $names }
            }
        }
    } finally {
        $ErrorActionPreference = $prevEap
        $env:PYTHONWARNINGS = $prevWarn
        $env:PYTHONIOENCODING = $prevEnc
        try { if ($prevOut) { [Console]::OutputEncoding = $prevOut } } catch {}
        # remove the temporary folder step 2 worked in (only ever our own, inside %TEMP%)
        try {
            $nt = $Script:NrTemp
            if ($nt -and $nt.StartsWith((Join-Path $env:TEMP 'DashboardReports-cleaned-'), [StringComparison]::OrdinalIgnoreCase) -and (Test-Path -LiteralPath $nt)) {
                Remove-Item -LiteralPath $nt -Recurse -Force -ErrorAction SilentlyContinue
            }
        } catch {}
    }

    # ----- summary -----
    if ($results.Count -eq 0) { return }
    W ''
    W ('SUMMARY  -  {0}' -f (Get-MonthLabel)) Cyan
    foreach ($r in $results) {
        $col = 'Green'
        if ($r.Status -eq 'FAIL' -or $r.Status -eq 'SKIP') { $col = 'Red' } elseif ($r.Status -eq 'WARN') { $col = 'Yellow' }
        W ('  [{0}] {1,-20} {2,-5} {3,6:N1} s   {4}' -f $r.Id, $r.Name, $r.Status, $r.Secs, $r.Files) $col
    }
    $passed = @($results | Where-Object { $_.Status -eq 'PASS' }).Count
    W ''
    W ('{0} of {1} passed.   Log: {2}' -f $passed, $results.Count, $log) DarkGray
    if (@($results | Where-Object { $_.Id -ge 3 -and $_.Status -eq 'PASS' }).Count -gt 0) {
        W 'Ready and Engage read the same folder - a "Skipped : <other report>" line from them is normal.' DarkGray
    }

    $after = @()
    if ($passed -gt 0) { $after += 'Refresh the Power BI report (Home > Refresh). If it reports a missing file, check Transform data > Data source settings.' }
    foreach ($r in $results) { if ($r.Status -ne 'PASS') { $after += ('Step {0} ({1}): {2} - {3}' -f $r.Id, $r.Name, $r.Status, $r.Files) } }
    Show-ExtraSteps $after
}

function Start-Run([int[]]$Ids, [bool]$Execute = $true) {
    if ($Ids.Count -eq 0) {
        W ''
        W 'Nothing is ticked - press 1-4 to tick a script, or A to run all.' Yellow
        Wait-Back
        return
    }
    Clear-Host
    Invoke-Safe { Invoke-Steps $Ids $Execute $false }
    Wait-Back
}

function Invoke-EditDefaults {
    $bat = $env:DR_BAT
    if (-not $bat -or -not (Test-Path -LiteralPath $bat)) { W 'Cannot find the .bat file to open.' Red; return }
    Start-Process -FilePath 'notepad.exe' -ArgumentList ('"' + $bat + '"')
    W 'Opened in Notepad. Edit the DEFAULTS block near the top and save,' Cyan
    W 'then close and re-open this window for the change to apply.' Cyan
}

# =====================================================================
#  HELP
# =====================================================================
function Show-About {
    W ''
    W "ABOUT (v$DR_VERSION)" Yellow
    W '  Runs the 4 monthly report scripts for the dashboards, with every'
    W '  folder worked out from the month you pick:'
    W '    1  Login time          Login_time.py'
    W '    2  Not Ready reasons   report_pipeline.py'
    W '    3  Ready time          clean_ready_time.py'
    W '    4  Engage, Hold, Wrap  clean_engage_time.py'
    W '  The .py files are used unchanged. They are looked for next to this .bat,'
    W '  then in a Utiliz folder, then in utils. Start it from anywhere.'
    W '  Needs Python on PATH (pandas + openpyxl installed).'
    W ''
    W '  Defaults: month = last month, folders = the Box layout, all 4 ticked.'
    W '  They are in the DEFAULTS block at the top of this .bat (key E opens it).'
    W '  Any other user on the same Box layout can use the same file: the'
    W '  C:\Users\<name> part of every path becomes the signed-in user.'
    W ''
    W '  Input folders are found by the short month name first (Oct 2026,'
    W '  Sept Login-LogOut 2026) and then by the other spellings (September 2026).'
    W '  Anything can be pointed somewhere else with P (Paths) or on the command line.'
    W ''
    W '  Keys: press the number/letter, Enter = default, Esc = back.'
    W '  Running: Enter (ticked) or A (all) shows a PLAN - inputs, outputs, which'
    W '  existing files get overwritten, problems - and waits for Y. Any other'
    W '  key cancels. V shows the plan and the exact commands, and runs nothing.'
    W '  Not Ready is built in a fresh temp folder and the cleaned files are then'
    W '  copied back into the month folder\cleaned, so the same month can be run again.'
    W '  A script is skipped (with the reason) if its input is missing or its'
    W '  output file is open in Excel. Missing folders are listed and created'
    W '  only after you say yes. A log of the last run: %TEMP%\DashboardReports-last-run.log'
}

function Show-Parameters {
    W ''
    W 'PARAMETERS  (run without the menu)' Yellow
    W '  Dashboard-Reports.bat run [scripts] [options]         PREVIEW: plan + commands, runs nothing'
    W '  Dashboard-Reports.bat run [scripts] [options] -Yes    RUN for real (no questions asked)'
    W '  Dashboard-Reports.bat show [options]                  month, folders and status'
    W ''
    W '  scripts    all | 1,2,3,4 or any subset, e.g. 1,3     default: the Default ticks'
    W '  -Month     2026-09 | sep | September | last | this   default: last'
    W '  -Box       Box folder                                default: <your profile>\Box'
    W '  -LoginIn -LoginOut -NotReadyIn -NotReadyOut -UtilIn -ReadyOut -EngageOut'
    W '             a full path that replaces just that one folder'
    W '  -CreateMissing   create missing folders without asking (menu mode always asks)'
    W '  -h         show this help'
    W ''
    W 'DEFAULTS' Yellow
    W ('  Month: {0}     Box: {1}     Default ticks: {2}' -f $Defaults.Month, $(if ($Defaults.Box) { $Defaults.Box } else { 'auto' }), $Defaults.Selected)
    W ('  Short month names: {0}' -f ($Script:Abbr -join ' '))
    W '  Folders (relative to Box; {Mon} {Month} {Year} {YY} {MM} are filled in):'
    foreach ($key in $Script:Slots.Keys) { W ('    {0,-12} {1}' -f $key, $Script:Slots[$key].Pattern) }
    W ''
    W '  Example: Dashboard-Reports.bat run 1,3 -Month 2026-09          (preview)'
    W '           Dashboard-Reports.bat run all -Month 2026-09 -Yes      (run all four)'
}

# =====================================================================
#  MENUS
# =====================================================================
function Menu-Month {
    $cur = ('{0}-{1:00}' -f $Script:M.Year, $Script:M.Num)
    $v = Read-Input 'Month: 2026-09, sep, September, last, this' $cur
    if ($null -eq $v) { return }
    $m = Resolve-Month $v
    if ($null -eq $m) { W "Cannot read '$v' as a month." Red; Wait-Back; return }
    $Script:M = $m
}

function Step-Month([int]$delta) {
    $d = (New-Object DateTime $Script:M.Year, $Script:M.Num, 1).AddMonths($delta)
    $Script:M = [pscustomobject]@{ Year = $d.Year; Num = $d.Month }
}

function Menu-Paths {
    while ($true) {
        $items = @()
        $i = 0
        foreach ($key in $Script:Slots.Keys) {
            $i++
            $s = $Script:Slots[$key]
            $inf = Get-SlotInfo $s
            $flag = ''
            if ($inf.Source -ne 'default') { $flag = "  ($($inf.Source))" }
            if (-not $inf.Exists) { $flag += '  [missing]' }
            $items += ('{0}|{1,-27} {2}{3}' -f $i, $s.Label, (Get-ShortPath $inf.Path), $flag)
        }
        $items += 'B|Box folder'
        $items += 'R|Reset every path to its default'
        $c = Show-Menu -Title 'Paths' -Items $items -Info {
            W ("  Month {0}   Box {1}" -f (Get-MonthLabel), $Script:BoxRoot) DarkGray
            W '  Pick a number to type a custom path (- resets that one). Changes last until you close this window;' DarkGray
            W '  to change them for good, edit the DEFAULTS block (E on the main screen).' DarkGray
        }
        if ($null -eq $c) { return }
        W ''
        if ($c -eq 'B') {
            $v = Read-Input 'Box folder (Esc = cancel)' $Script:BoxRoot
            if ($null -ne $v -and $v -ne '') { $Script:BoxRoot = Convert-UserPath $v }
        } elseif ($c -eq 'R') {
            foreach ($key in $Script:Slots.Keys) { $Script:Slots[$key].Custom = $null }
            $Script:BoxRoot = Resolve-BoxRoot
        } else {
            $key = @($Script:Slots.Keys)[[int]$c - 1]
            $s = $Script:Slots[$key]
            $cur = (Get-SlotInfo $s).Path
            $v = Read-Input ("New path for '{0}' (Esc = cancel, - = reset)" -f $s.Label) $cur
            if ($null -eq $v -or $v -eq $cur) { continue }
            if ($v -eq '-') { $s.Custom = $null } else { $s.Custom = $v }
        }
    }
}

function Menu-Help {
    while ($true) {
        $c = Show-Menu -Title 'Help' -Items @('1|About', '2|Parameters (command line)')
        if ($null -eq $c) { return }
        Clear-Host
        if ($c -eq '1') { Show-About } else { Show-Parameters }
        Wait-Back
    }
}

function Menu-Main {
    while ($true) {
        Clear-Host
        Write-MainScreen
        Write-Host -NoNewline '  Choose: '
        $k = Read-Key
        if ($k.Key -eq 'Escape') { W ''; return }
        if ($k.Key -eq 'Enter') { W ''; Start-Run (Get-TickedIds) $true; continue }
        $ch = Get-KeyChar $k
        switch -Regex ($ch) {
            '^[1-4]$' { $id = [int]$ch; $Script:Sel[$id] = -not $Script:Sel[$id] }
            '^T$'     { $any = @(Get-TickedIds).Count -lt 4; foreach ($id in 1..4) { $Script:Sel[$id] = $any } }
            '^D$'     { Set-DefaultTicks }
            '^A$'     { W ''; Start-Run @(1, 2, 3, 4) $true }
            '^V$'     { W ''; Start-Run (Get-TickedIds) $false }
            '^M$'     { W ''; Menu-Month }
            '^\+$'    { Step-Month 1 }
            '^-$'     { Step-Month -1 }
            '^P$'     { Menu-Paths }
            '^C$'     {
                Clear-Host
                $ids = Get-TickedIds
                if ($ids.Count -eq 0) { W 'Nothing is ticked.' Yellow }
                elseif (@(Get-MissingFolders $ids).Count -eq 0) { W ''; W 'No missing folders for the ticked scripts.' Green }
                else { Invoke-Safe { [void](Confirm-MissingFolders $ids $false) } }
                Wait-Back
            }
            '^E$'     { W ''; Invoke-EditDefaults; Wait-Back }
            '^H$'     { Menu-Help }
        }
    }
}

# =====================================================================
#  ENTRY POINT
# =====================================================================
if ($Help -or $Action -match '^(help|/\?|/h|--help)$') {
    Show-About
    Show-Parameters
    W ''
    Remove-TempCopies
    exit 0
}

$Script:BoxRoot = Resolve-BoxRoot
$monthText = $Month
if (-not $monthText) { $monthText = $Defaults.Month }
$Script:M = Resolve-Month $monthText
if ($null -eq $Script:M) {
    W "Cannot read '$monthText' as a month. Use e.g. 2026-09, sep, last, this." Red
    Remove-TempCopies
    exit 1
}
Set-DefaultTicks

if ($Action) {
    # Command-line mode: do one thing and exit
    $ok = $true
    switch ($Action.ToLower()) {
        'run' {
            $ids = Resolve-Ids $Scripts
            if ($null -eq $ids) { W 'Usage: run [all | 1,2,3,4 or a subset] [-Yes]   (-h for help)' Red; $ok = $false }
            else {
                Invoke-Safe { Invoke-Steps $ids $Yes.IsPresent $Yes.IsPresent }
                if (-not $Script:RunOk) { $ok = $false }
            }
        }
        'show' { Invoke-Safe { Show-Summary } }
        default { W "Unknown action '$Action'. Use run or show (-h for help)." Red; $ok = $false }
    }
    Remove-TempCopies
    if ($ok) { exit 0 } else { exit 1 }
}

# No arguments: interactive menu (double-click). Esc on the menu exits here.
$Script:Interactive = $true
Menu-Main
Remove-TempCopies
Clear-Host
exit 0
