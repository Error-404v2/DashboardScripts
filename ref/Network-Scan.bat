@echo off
setlocal
set "NS_FILE=%TEMP%\FreshSetupToolkit-%RANDOM%.ps1"
more +10 "%~f0" > "%NS_FILE%"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%NS_FILE%" %*
set "NS_ERR=%errorlevel%"
del /f /q "%NS_FILE%" >nul 2>&1
if not "%NS_ERR%"=="0" pause
exit /b %NS_ERR%
:: ===== PowerShell part below (copied out by the lines above) =====
# =====================================================================
#  Network Scan - PowerShell part of the single-file Network-Scan.bat.
#  The batch lines above copy everything from here down into a freshly
#  randomly-named .ps1 in %TEMP% (deleted again when the script exits)
#  and run it with -File (Windows PowerShell 5.1 only accepts -File for
#  names ending in .ps1).
#  Finds the network this PC is on, pings every address in it, and
#  lists IP, MAC and ping time of each live device as it's found.
#  Run with -h for About + Parameters.
# =====================================================================
#  VERSION HISTORY
#  1.1  2026-10-06  Fix: the "scan is large, continue?" confirmation only
#                    accepted a literal Y keypress - Enter's KeyChar isn't
#                    'Y', so pressing Enter fell through to "Cancelled."
#                    instead of continuing like every other menu/prompt in
#                    this script. Enter now continues (shown as Enter/Y).
#  1.0  (initial)    First tracked version.
# =====================================================================

param(
    [Parameter(Position = 0)] [Alias('Subnet')] [string]$Network = '',
    [int]$Count   = 2,
    [int]$Timeout = 100,
    [switch]$Force,
    [Alias('h')]
    [switch]$Help
)

$Script:NS_VERSION = '1.1'
$ErrorActionPreference = 'Stop'
try { $Host.UI.RawUI.WindowTitle = "Network Scan v$NS_VERSION" } catch {}

# Removes the temp .ps1 copy this run came from, plus leftovers from earlier
# runs: random-named copies in %TEMP% (left behind when a window is closed
# with the X button, which skips the batch "del" line), and the old fixed
# %LOCALAPPDATA%\FreshSetupTools folder used by earlier versions. Safe while
# running: PowerShell has already read the whole script into memory.
function Remove-TempCopies {
    try {
        if ($PSCommandPath -and (Split-Path -Leaf $PSCommandPath) -like 'FreshSetupToolkit-*.ps1') {
            Remove-Item -LiteralPath $PSCommandPath -Force -ErrorAction SilentlyContinue
        }
    } catch {}
    try {
        if ($env:TEMP) {
            Get-ChildItem -LiteralPath $env:TEMP -Filter 'FreshSetupToolkit-*.ps1' -File -ErrorAction SilentlyContinue |
                Remove-Item -Force -ErrorAction SilentlyContinue
        }
    } catch {}
    try {
        $old = Join-Path $env:LOCALAPPDATA 'FreshSetupTools'
        if ($env:LOCALAPPDATA -and (Test-Path -LiteralPath $old)) {
            Get-ChildItem -LiteralPath $old -Filter '*.ps1' -File -ErrorAction SilentlyContinue |
                Remove-Item -Force -ErrorAction SilentlyContinue
            if (-not (Get-ChildItem -LiteralPath $old -Force -ErrorAction SilentlyContinue)) {
                Remove-Item -LiteralPath $old -Force -ErrorAction SilentlyContinue
            }
        }
    } catch {}
}
Remove-TempCopies

function W {
    param([string]$Text = '', [string]$Color)
    if ($Color) { Write-Host $Text -ForegroundColor $Color } else { Write-Host $Text }
}

function Read-Key {
    try { while ([Console]::KeyAvailable) { [void][Console]::ReadKey($true) } } catch {}
    return [Console]::ReadKey($true)
}

function Read-Input {
    param([string]$Prompt, [string]$Default = '')
    Write-Host -NoNewline "  $Prompt"
    if ($Default -ne '') { Write-Host -NoNewline " [$Default]" -ForegroundColor DarkGray }
    Write-Host -NoNewline ': '
    $buf = New-Object System.Text.StringBuilder
    # No stale-key flush here on purpose: flushing right before reading can
    # eat real characters if you type ahead quickly across prompts.
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
        # Match on KeyChar AND the ConsoleKey enum name - KeyChar alone can
        # come back empty/wrong on some console window types (seen: options
        # not registering on a freshly double-clicked window even though
        # Escape, which only checks .Key, worked fine there). D1/D2/.. covers
        # top-row digits, NumPad1/.. covers the numeric keypad.
        $chFromChar = ([string]$k.KeyChar).ToUpper()
        $keyName = $k.Key.ToString()
        $chFromKey = $null
        if ($keyName -match '^D([0-9])$') { $chFromKey = $Matches[1] }
        elseif ($keyName -match '^NumPad([0-9])$') { $chFromKey = $Matches[1] }
        elseif ($keyName.Length -eq 1) { $chFromKey = $keyName.ToUpper() }
        foreach ($it in $Items) {
            $tag = $it.Split('|', 2)[0]
            if ($tag -eq $chFromChar -or ($chFromKey -and $tag -eq $chFromKey)) { W $tag; return $tag }
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
    try { & $Block } catch { W ''; W "ERROR: $($_.Exception.Message)" Red }
}

# =====================================================================
#  HELP
# =====================================================================
function Show-About {
    W ''
    W "ABOUT (v$NS_VERSION)" Yellow
    W '  Single file: copy Network-Scan.bat to any Windows 10/11 PC and'
    W '  double-click it. Nothing else to copy, nothing to install.'
    W '  Finds the network this PC is on (the one with the default gateway),'
    W '  pings every address in it, and lists live devices right away with'
    W '  their IP, MAC address and ping time.'
    W ''
    W '  Keys: press the number, Enter = default, Esc = back/stop.'
}

function Show-Parameters {
    W ''
    W 'PARAMETERS' Yellow
    W '  Network-Scan.bat [range] [options]'
    W ''
    W '  range      CIDR (192.168.1.0/24) or 3 parts (192.168.1)'
    W '             default: auto-detect this PC''s network'
    W '  -Count     pings per address        default 2'
    W '  -Timeout   ms per ping              default 100'
    W '  -Force     no prompt for ranges over 1024 addresses'
    W '  -h         show this help'
    W ''
    W '  Example: Network-Scan.bat 192.168.1 -Count 1 -Timeout 20'
}

# =====================================================================
#  IP / SCAN HELPERS
# =====================================================================
function Get-PrefixFromMask([string]$mask) {
    $bits = 0
    foreach ($byte in ([System.Net.IPAddress]::Parse($mask)).GetAddressBytes()) {
        $bits += ([Convert]::ToString($byte, 2).ToCharArray() | Where-Object { $_ -eq '1' }).Count
    }
    return $bits
}
function ConvertTo-Num([string]$ip) {
    $b = ([System.Net.IPAddress]::Parse($ip)).GetAddressBytes()
    return ([int64]$b[0] * 16777216) + ([int64]$b[1] * 65536) + ([int64]$b[2] * 256) + [int64]$b[3]
}
function ConvertTo-IP([int64]$n) {
    return '{0}.{1}.{2}.{3}' -f (($n -shr 24) -band 255), (($n -shr 16) -band 255), (($n -shr 8) -band 255), ($n -band 255)
}
function Format-Mac($bytes) {
    if (-not $bytes -or $bytes.Length -eq 0) { return $null }
    return (($bytes | ForEach-Object { $_.ToString('x2') }) -join '-')
}
function Format-Ms($ms) { if ($ms -lt 1) { return '<1 ms' }; return ('{0} ms' -f [Math]::Round($ms)) }
function Get-LineWidth { try { return [Math]::Max(40, [Console]::WindowWidth - 1) } catch { return 79 } }
function Write-Status([string]$text) {
    $w = Get-LineWidth
    if ($text.Length -gt $w) { $text = $text.Substring(0, $w) }
    Write-Host ("`r" + $text.PadRight($w)) -NoNewline -ForegroundColor DarkGray
}
function Clear-Status { Write-Host ("`r" + (' ' * (Get-LineWidth)) + "`r") -NoNewline }
function Get-Mac([string]$ip) {
    # Two separate problems here, not one:
    #  1. "arp -a <ip>" only searches the ARP table of whichever interface
    #     Windows picks as default when you don't say -N - on a PC with more
    #     than one active NIC (Wi-Fi + Ethernet, a VPN adapter, etc.) that can
    #     silently be the WRONG interface even though the ping itself went
    #     out and came back fine. Fixed by dumping the FULL table ("arp -a",
    #     no IP filter - it lists every interface) and grep'ing it ourselves.
    #  2. The OS can take a few ms to commit a fresh ARP entry after a ping
    #     reply, which a fast scan (low -Count/-Timeout) can outrun. Fixed
    #     by retrying with a short wait instead of a single immediate read.
    for ($try = 1; $try -le 3; $try++) {
        $line = $null
        try { $line = arp -a 2>$null | Select-String -Pattern ('^\s*' + [regex]::Escape($ip) + '\s') } catch {}
        if ($line) {
            $m = [regex]::Match($line.Line, '([0-9a-fA-F]{2}[-:]){5}[0-9a-fA-F]{2}')
            if ($m.Success) { return $m.Value.ToLower() }
        }
        if ($try -lt 3) { Start-Sleep -Milliseconds 40 }
    }
    return $null
}

function Get-DefaultNetwork {
    $nics = [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces() | Where-Object {
        $_.OperationalStatus -eq 'Up' -and $_.NetworkInterfaceType -ne 'Loopback' -and $_.NetworkInterfaceType -ne 'Tunnel'
    }
    $candidates = @()
    foreach ($nic in $nics) {
        $props = $nic.GetIPProperties()
        $gw = $props.GatewayAddresses | Where-Object {
            $_.Address.AddressFamily -eq 'InterNetwork' -and $_.Address.ToString() -ne '0.0.0.0'
        } | Select-Object -First 1
        if (-not $gw) { continue }
        $ua = $props.UnicastAddresses | Where-Object {
            $_.Address.AddressFamily -eq 'InterNetwork' -and $_.Address.ToString() -notlike '169.254.*'
        } | Select-Object -First 1
        if (-not $ua) { continue }

        $prefix = 0
        try { $prefix = [int]$ua.PrefixLength } catch {}
        if ($prefix -le 0 -or $prefix -gt 32) {
            try { $prefix = Get-PrefixFromMask $ua.IPv4Mask.ToString() } catch { $prefix = 24 }
        }
        $index = $null
        try { $index = $props.GetIPv4Properties().Index } catch {}

        $candidates += [PSCustomObject]@{
            Name = $nic.Name; IP = $ua.Address.ToString(); Prefix = $prefix; Gateway = $gw.Address.ToString()
            Mac  = Format-Mac $nic.GetPhysicalAddress().GetAddressBytes(); Index = $index; Metric = [int]::MaxValue
        }
    }
    if ($candidates.Count -eq 0) { return $null }

    if ($candidates.Count -gt 1) {
        try {
            $routes = Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction Stop
            foreach ($c in $candidates) {
                $r = $routes | Where-Object InterfaceIndex -eq $c.Index | Select-Object -First 1
                if ($r) {
                    $ifMetric = 0
                    try { $ifMetric = (Get-NetIPInterface -InterfaceIndex $c.Index -AddressFamily IPv4 -ErrorAction Stop).InterfaceMetric } catch {}
                    $c.Metric = [int]$r.RouteMetric + [int]$ifMetric
                }
            }
        } catch {}
    }
    return $candidates | Sort-Object Metric | Select-Object -First 1
}

function Test-StopKey {
    try {
        while ([Console]::KeyAvailable) {
            $k = [Console]::ReadKey($true)
            if ($k.Key -eq 'Escape') { return $true }
            if ($k.Key -eq 'C' -and ($k.Modifiers -band [ConsoleModifiers]::Control)) { return $true }
        }
    } catch {}
    return $false
}

function Start-NetworkScan {
    param([string]$Net = '', [int]$Cnt = 2, [int]$To = 100, [bool]$NoPrompt = $false)

    $thisPC = $null; $gw = $null
    if ([string]::IsNullOrWhiteSpace($Net)) {
        $thisPC = Get-DefaultNetwork
        if (-not $thisPC) {
            W "Couldn't find a connected network with a default gateway." Red
            Show-ExtraSteps @('Connect to a network, or scan a custom range.')
            return
        }
        $baseIP = $thisPC.IP; $prefix = $thisPC.Prefix; $gw = $thisPC.Gateway
    } elseif ($Net -match '^(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})/(\d{1,2})$') {
        $baseIP = $Matches[1]; $prefix = [int]$Matches[2]
    } elseif ($Net -match '^\d{1,3}\.\d{1,3}\.\d{1,3}$') {
        $baseIP = "$Net.0"; $prefix = 24
    } else {
        W "Invalid range '$Net'. Use e.g. 192.168.1 or 192.168.1.0/24" Red
        return
    }
    try { $null = [System.Net.IPAddress]::Parse($baseIP) } catch { W "Invalid address '$baseIP'." Red; return }
    if ($prefix -lt 8 -or $prefix -gt 30) { W "Prefix /$prefix isn't supported. Use /8 to /30." Red; return }
    if ($Cnt -lt 1) { $Cnt = 1 }
    if ($To -lt 1)  { $To = 1 }

    $block   = [int64][Math]::Pow(2, 32 - $prefix)
    $netNum  = (ConvertTo-Num $baseIP) - ((ConvertTo-Num $baseIP) % $block)
    $firstN  = $netNum + 1
    $lastN   = $netNum + $block - 2
    $total   = [int64]($lastN - $firstN + 1)
    $netText = '{0}/{1}' -f (ConvertTo-IP $netNum), $prefix

    if ($total -gt 1024 -and -not $NoPrompt) {
        $estMin = [Math]::Ceiling($total * $Cnt * $To / 60000)
        W "$netText has $total addresses - worst case about $estMin minutes." Yellow
        Write-Host -NoNewline '  Enter/Y = continue, Esc / any other key = cancel: '
        $k = Read-Key
        W ''
        if ($k.Key -ne 'Enter' -and ([string]$k.KeyChar).ToUpper() -ne 'Y') { W 'Cancelled.'; return }
    }

    W ''
    if ($thisPC) {
        W ('Network: {0}  via {1}  (this PC {2}, gateway {3})' -f $netText, $thisPC.Name, $thisPC.IP, $thisPC.Gateway) Cyan
    } else {
        W ('Network: {0}' -f $netText) Cyan
    }
    W ('Scanning {0} - {1}  ({2} addresses, {3} ping(s) each, {4} ms timeout)' -f (ConvertTo-IP $firstN), (ConvertTo-IP $lastN), $total, $Cnt, $To) Cyan
    W 'Press Esc (or Ctrl+C) to stop.'
    W ''
    $fmt = '{0,-4} {1,-16} {2,-19} {3,-8} {4,-13} {5}'
    W ($fmt -f '#', 'IP', 'MAC', 'Ping', 'Pos', 'Note') Cyan
    W ($fmt -f '--', '--', '---', '----', '---', '----') Cyan

    $pinger  = New-Object System.Net.NetworkInformation.Ping
    $found   = New-Object System.Collections.ArrayList
    $timer   = [Diagnostics.Stopwatch]::StartNew()
    $lastMsg = '-'; $pos = 0; $stopped = $false

    $oldCtrlC = $false
    try { $oldCtrlC = [Console]::TreatControlCAsInput; [Console]::TreatControlCAsInput = $true } catch {}
    try {
        for ($n = $firstN; $n -le $lastN; $n++) {
            $ip = ConvertTo-IP $n
            $pos++
            $pct = [Math]::Floor($pos * 100 / $total)
            $times = @()
            for ($c = 1; $c -le $Cnt; $c++) {
                if (Test-StopKey) { $stopped = $true; break }
                Write-Status ('> Pinging {0,-15} [{1}/{2}] {3,3}%  try {4}/{5}  last: {6}  found: {7}' -f $ip, $pos, $total, $pct, $c, $Cnt, $lastMsg, $found.Count)
                try {
                    $reply = $pinger.Send($ip, $To)
                    if ($reply.Status -eq 'Success') { $times += $reply.RoundtripTime }
                } catch {}
            }
            if ($stopped) { break }

            if ($times.Count -gt 0) {
                $avg = ($times | Measure-Object -Average).Average
                $lastMsg = "$ip " + (Format-Ms $avg)
                $note = ''; $mac = $null
                if ($thisPC -and $ip -eq $thisPC.IP) { $note = 'this PC'; $mac = $thisPC.Mac }
                else { $mac = Get-Mac $ip; if ($gw -and $ip -eq $gw) { $note = 'gateway' } }
                if (-not $mac) { $mac = '(unknown)' }
                [void]$found.Add([PSCustomObject]@{ IP = $ip; MAC = $mac; Ms = $avg })
                Clear-Status
                W ($fmt -f $found.Count, $ip, $mac, (Format-Ms $avg), "[$pos/$total]", $note) Green
            } else {
                $lastMsg = "$ip timeout"
            }
        }
    } finally {
        try { [Console]::TreatControlCAsInput = $oldCtrlC } catch {}
    }

    $timer.Stop()
    Clear-Status
    W ''
    if ($stopped) { W "Stopped at $ip ($pos of $total)." Yellow }
    W ('Done: {0} device(s) found, {1} of {2} addresses checked in {3:N1} s' -f $found.Count, $pos, $total, $timer.Elapsed.TotalSeconds) Cyan
    if ($found.Count -gt 0) {
        $fastest = $found | Sort-Object Ms | Select-Object -First 1
        $slowest = $found | Sort-Object Ms -Descending | Select-Object -First 1
        W ('Fastest: {0} ({1})   Slowest: {2} ({3})' -f $fastest.IP, (Format-Ms $fastest.Ms), $slowest.IP, (Format-Ms $slowest.Ms))
    }
    if (@($found | Where-Object MAC -eq '(unknown)').Count -gt 0) {
        W '(unknown) MAC = the device replied but is not in the ARP table, e.g. it is on another subnet.' DarkGray
    }
    Show-ExtraSteps @()
}

function Invoke-CustomScan {
    $r = Read-Input 'Range, e.g. 192.168.1 or 10.0.0.0/22 (Enter = auto)' ''
    if ($null -eq $r) { return $false }
    $p = Read-Input 'Pings per address' '2'
    if ($null -eq $p) { return $false }
    $t = Read-Input 'Timeout per ping in ms' '100'
    if ($null -eq $t) { return $false }
    $pi = 0; $ti = 0
    if (-not [int]::TryParse($p, [ref]$pi) -or -not [int]::TryParse($t, [ref]$ti)) {
        W 'Pings and timeout must be whole numbers.' Red
    } else {
        Invoke-Safe { Start-NetworkScan $r $pi $ti }
    }
    return $true
}

function Menu-Scan {
    while ($true) {
        $c = Show-Menu -Title "Network scan (v$NS_VERSION)" -EscLabel 'Exit' -Default '1' -Items @(
            '1|Scan my network  (1 ping, 60 ms - may miss slow devices)',
            '2|Thorough scan    (3 pings, 500 ms)',
            '3|Custom range / pings / timeout',
            'H|Help') -Info { W '  Esc or Ctrl+C stops a running scan.' DarkGray }
        if ($null -eq $c) { return }
        W ''
        $wait = $true
        switch ($c) {
            '1' { Invoke-Safe { Start-NetworkScan '' 1 60 } }
            '2' { Invoke-Safe { Start-NetworkScan '' 3 500 } }
            '3' { $wait = Invoke-CustomScan }
            'H' { Clear-Host; Show-About; Show-Parameters }
        }
        if ($wait) { Wait-Back }
    }
}

# =====================================================================
#  ENTRY POINT
# =====================================================================
if ($Help -or $Network -match '^(--?h(elp)?|/\?|/h)$') {
    Show-About
    Show-Parameters
    W ''
    Remove-TempCopies
    exit 0
}

if ($Network) {
    # Command-line mode: scan once and exit
    Invoke-Safe { Start-NetworkScan $Network $Count $Timeout $Force.IsPresent }
    Remove-TempCopies
    exit 0
}

# No arguments: interactive menu (double-click). Esc on the menu exits here.
Menu-Scan
Remove-TempCopies
Clear-Host
exit 0
