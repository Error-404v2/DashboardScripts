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
#  Fresh Setup Toolkit - PowerShell part of the single-file
#  Fresh-Setup-Toolkit.bat. The batch lines above copy everything from
#  here down into a fresh randomly-named .ps1 in %TEMP% and run it with
#  -File (Windows PowerShell 5.1 only accepts -File for names ending in
#  .ps1). A random name each run avoids a real bug: a FIXED filename
#  that was once written by an elevated (admin) run can end up with
#  permissions a later non-elevated run can't overwrite, which silently
#  breaks the copy step and makes the window flash and close instantly.
#  Run with -h for About + Parameters.
#
#  Network scan lives in its own file on purpose: Network-Scan.bat.
# =====================================================================

param(
    [Parameter(Position = 0)] [string]$Tool   = '',
    [Parameter(Position = 1)] [string]$Action = '',

    # Omada route
    [string]$Network = '',
    [string]$Mask    = '255.255.255.0',
    [string]$Gateway = '192.168.1.57',
    [string]$TestIP  = '192.168.0.101',

    # SMB user - print-only, nothing here is ever executed by this script.
    # Always the exact name/password typed at runtime, never defaulted to
    # the current logged-in user.
    [string]$User = '',
    [string]$Password = '',

    # Claude models
    [string[]]$Models = @('claude-sonnet-5', 'claude-opus-5-5', 'claude-haiku-4-5'),
    [ValidateSet('low', 'medium', 'high', 'xhigh', 'max')]
    [string]$DefaultEffort = 'medium',
    # [bool] here looks natural but is actually broken for a .bat/CLI tool:
    # when PowerShell is launched with -File, every argument arrives as a
    # plain string, and PowerShell's own [bool] converter rejects "0", "1",
    # "true", and "True" identically ("Cannot convert value "System.String"
    # to type "System.Boolean""). Confirmed by direct reproduction. Using a
    # string here instead and parsing it ourselves (see Resolve-Bool) is
    # the fix - don't change this back to [bool].
    [string]$AlwaysUseDefault = 'true',

    # This PC folders shortcut icon. 104 is the confirmed-working value in
    # imageres.dll on Ahmed's PC (verified by reading the IconIndex field
    # back out of a shortcut he fixed by hand via Properties > Change Icon).
    # The registry's own DefaultIcon value for this CLSID uses a NEGATIVE
    # resource-ID style reference (imageres.dll,-109), but a .lnk file's
    # binary IconIndex field does not resolve that the same way - it wants a
    # plain positive array index, and -109 rendered blank. Don't change this
    # back to a negative number.
    [int]$IconIndex = 104,

    # Internal: set when the toolkit relaunched itself as administrator
    [switch]$Elevated,

    [Alias('h')]
    [switch]$Help
)

$ErrorActionPreference = 'Stop'
$script:Interactive = $false
try { $Host.UI.RawUI.WindowTitle = 'Fresh Setup Toolkit' } catch {}

# Fixed values
$SpotClsid     = '{2cc5ca98-6485-489a-920e-b3e88a6ccce3}'
$SpotKey       = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel'
$SpotTask      = 'HideSpotlightDesktopIcon'
$LocalData     = if ($env:LOCALAPPDATA) { $env:LOCALAPPDATA } else { [Environment]::GetFolderPath('LocalApplicationData') }
$ClaudeLib     = Join-Path $LocalData 'Claude-3p\configLibrary'
$WebSearchKey  = 'HKCU:\Software\Policies\Microsoft\Windows\Explorer'
$NameSpaceRoot = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace'
# Quick access (the Explorer sidebar) is a completely different mechanism
# from the NameSpace registry keys above - those only control what shows up
# INSIDE the "This PC" window's own main canvas, not the sidebar. Pinning to
# Quick access is done through the Shell.Application COM "pintohome" /
# "unpinfromhome" verbs instead, which is why it is a separate feature below.
$QuickAccessNamespace = 'shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}'
# Desktop is deliberately NOT in this list: Explorer already shows Desktop in
# the sidebar on its own, so unhiding its CLSID here only adds an unwanted
# duplicate tile in the This PC / Home "Folders" gallery. Its CLSID is kept
# below so Invoke-ThisPcFolders can make sure it stays at the Windows default
# (hidden) even if an earlier version of this tool changed it.
$DesktopClsid  = '{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}'
$ThisPcClsids  = [ordered]@{
    'Documents' = '{d3162b92-9365-467a-956b-92703aca08af}'
    'Downloads' = '{088e3905-0323-4b02-9826-5d99428e115f}'
    'Music'     = '{3dfdf296-dbec-4fb4-81d1-6a3438bcf4de}'
    'Pictures'  = '{24ad3ad4-a569-4530-98e1-ab02f9417aa8}'
    'Videos'    = '{f86fa3ab-70d2-4fc7-9c99-fcbf05467f3a}'
}

# =====================================================================
#  GENERAL HELPERS
# =====================================================================
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

function Read-Password {
    param([string]$Prompt)
    Write-Host -NoNewline "  $Prompt"
    Write-Host -NoNewline ': '
    $buf = New-Object System.Text.StringBuilder
    while ($true) {
        $k = [Console]::ReadKey($true)
        if ($k.Key -eq 'Escape') { W ''; return $null }
        if ($k.Key -eq 'Enter') { W ''; return $buf.ToString() }
        if ($k.Key -eq 'Backspace') {
            if ($buf.Length -gt 0) { [void]$buf.Remove($buf.Length - 1, 1); Write-Host -NoNewline "`b `b" }
            continue
        }
        if (-not [char]::IsControl($k.KeyChar)) { [void]$buf.Append($k.KeyChar); Write-Host -NoNewline '*' }
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
        $ch = ([string]$k.KeyChar).ToUpper()
        foreach ($it in $Items) { if ($it.Split('|', 2)[0] -eq $ch) { W $ch; return $ch } }
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

function Test-IsAdmin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    return ([Security.Principal.WindowsPrincipal]$id).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Format-Arg([string]$s) {
    if ($s -match '[\s"]') { return '"' + ($s -replace '"', '\"') + '"' }
    return $s
}

function Invoke-Elevated {
    # Re-launches THIS SAME SCRIPT FILE (no copying anywhere) in a new admin
    # window with -Verb RunAs, and waits for it to finish.
    param([string[]]$ToolArgs)
    W 'Asking for administrator rights (an admin window will open)...' Yellow
    try {
        $parts = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', (Format-Arg $PSCommandPath)) +
                  ($ToolArgs | ForEach-Object { Format-Arg $_ }) + '-Elevated'
        Start-Process -FilePath 'powershell.exe' -ArgumentList ($parts -join ' ') -Verb RunAs -Wait | Out-Null
        return $true
    } catch {
        W 'Administrator rights were not granted. Nothing was changed.' Yellow
        return $false
    }
}

function Resolve-Bool([string]$s) {
    if ($s -match '^(1|true|yes|on)$') { return $true }
    if ($s -match '^(0|false|no|off)$') { return $false }
    return $true
}

function Get-PrefixFromMask([string]$mask) {
    $bits = 0
    foreach ($byte in ([System.Net.IPAddress]::Parse($mask)).GetAddressBytes()) {
        $bits += ([Convert]::ToString($byte, 2).ToCharArray() | Where-Object { $_ -eq '1' }).Count
    }
    return $bits
}

# =====================================================================
#  HELP
# =====================================================================
function Show-About {
    W ''
    W 'ABOUT' Yellow
    W '  Single file: copy Fresh-Setup-Toolkit.bat to any Windows 10/11'
    W '  PC and double-click it. Nothing else to copy, nothing to install.'
    W '  Network scan is a separate single file: Network-Scan.bat.'
    W ''
    W '  Omada route        Permanent route: 192.168.0.x via the Omada router'
    W '                     192.168.1.57. Needs admin (asks automatically).'
    W '  Spotlight desktop  Hides/restores the "Learn about this picture"'
    W '  icon               desktop icon, and a logon task that re-hides it'
    W '                     at every sign-in so Windows updates can''t'
    W '                     bring it back.'
    W '  Claude models      Sets the default model + effort in Claude'
    W '                     Desktop, and limits its picker to exactly 3'
    W '                     models. Backs up first; revert undoes it.'
    W '  This PC folders    Each option toggles - it switches whichever way'
    W '                     is not currently active:'
    W '                      - Main view: Documents/Downloads/Music/Pictures/'
    W '                        Videos inside the This PC window itself'
    W '                        (Desktop is left alone). Needs admin.'
    W '                      - Quick access sidebar: Documents/Music/'
    W '                        Pictures/Videos pinned/unpinned together. No'
    W '                        admin needed - this is the normal "Pin to'
    W '                        Quick access" action, just automated.'
    W '                      - Quick access sidebar: Desktop + Downloads'
    W '                        pinned/unpinned together, separately from the'
    W '                        other 4 folders.'
    W '                     Can also add a classic "This PC" desktop'
    W '                     shortcut.'
    W '  Web search         Removes/restores Bing web results from the'
    W '                     taskbar and Start search. Needs admin (asks'
    W '                     automatically).'
    W '  SMB user           Create/hide/unhide/delete a local account for'
    W '                     SMB sharing - but this tool never runs any of'
    W '                     it. It only PRINTS the exact command (net'
    W '                     user / reg add / reg delete) for you to run'
    W '                     yourself in your own elevated window. For'
    W '                     hide/unhide/delete, pick an existing account'
    W '                     from a list (from "net user") or type a'
    W '                     different name. Always uses the exact'
    W '                     name/password you type or pick.'
    W ''
    W '  Keys: press the number/letter, Enter = default, Esc = back.'
}

function Show-Parameters {
    W ''
    W 'PARAMETERS  (run a tool directly, without the menu)' Yellow
    W '  Fresh-Setup-Toolkit.bat <tool> <action> [options]'
    W ''
    W '  omada      add | remove | status'
    W '    -Network -Mask -Gateway -TestIP   (see defaults below)'
    W ''
    W '  spotlight  hide | restore | task-on | task-off'
    W ''
    W '  claude     set | revert | undo       (-Models -DefaultEffort -AlwaysUseDefault)'
    W ''
    W '  thispc     show | hide                  (This PC main view)'
    W '             qa-pin | qa-unpin             (Quick access: Documents/Music/Pictures/Videos)'
    W '             qa-pin-desk | qa-unpin-desk   (Quick access: Desktop + Downloads)'
    W '             shortcut                      (-IconIndex, desktop shortcut)'
    W ''
    W '  websearch  remove | restore'
    W ''
    W '  smbuser    create|hide|unhide|delete -User <name> [-Password <pw>]'
    W '             (-Password only needed for create. Prints a command - runs nothing.)'
    W ''
    W '  -h         show this help'
    W ''
    W 'DEFAULTS' Yellow
    W '  omada:  -Network 192.168.0.0  -Mask 255.255.255.0'
    W '          -Gateway 192.168.1.57 -TestIP 192.168.0.101'
    W "  claude: -Models $($Models -join ',')"
    W '          -DefaultEffort medium (low|medium|high|xhigh|max) -AlwaysUseDefault true'
    W '  thispc: -IconIndex 104  (imageres.dll icon array index, confirmed working)'
    W ''
    W '  Example: Fresh-Setup-Toolkit.bat omada add -Gateway 192.168.1.1'
}

# =====================================================================
#  OMADA ROUTE
# =====================================================================
function Get-OmadaSettings {
    $net = if ($Network) { $Network } else { '192.168.0.0' }
    return [PSCustomObject]@{
        Network = $net; Mask = $Mask; Gateway = $Gateway; TestIP = $TestIP
        Cidr    = "$net/$(Get-PrefixFromMask $Mask)"
    }
}

function Test-OmadaSaved($o) {
    try {
        return @(Get-NetRoute -DestinationPrefix $o.Cidr -AddressFamily IPv4 -PolicyStore PersistentStore -ErrorAction SilentlyContinue |
                 Where-Object NextHop -eq $o.Gateway).Count -gt 0
    } catch { return $false }
}

function Get-OmadaStateText($o) {
    $active = $false
    try {
        $active = @(Get-NetRoute -DestinationPrefix $o.Cidr -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                    Where-Object NextHop -eq $o.Gateway).Count -gt 0
    } catch {}
    $saved = Test-OmadaSaved $o
    if ($saved -and $active) { return 'ON (permanent, active now)' }
    if ($saved)              { return 'ON (permanent, waiting for the ' + (($o.Gateway -split '\.')[0..2] -join '.') + '.x network)' }
    if ($active)             { return 'ON (temporary - lost on reboot)' }
    return 'OFF'
}

function Invoke-Omada([string]$Act) {
    $o = Get-OmadaSettings
    if ($Act -eq 'status') {
        W "Route $($o.Cidr) via $($o.Gateway): $(Get-OmadaStateText $o)" Cyan
        return
    }

    if (-not (Test-IsAdmin)) {
        $ok = Invoke-Elevated @('omada', $Act, '-Network', $o.Network, '-Mask', $o.Mask, '-Gateway', $o.Gateway, '-TestIP', $o.TestIP)
        if ($ok) {
            W ''
            W "Route now: $(Get-OmadaStateText $o)" Cyan
            W '(Details and any extra steps are shown in the admin window.)' DarkGray
        }
        return
    }

    # Clear any old copy first. On a first run there's nothing to delete, and
    # route.exe prints "Element not found" for that - expected, not a real
    # problem, so this is wrapped in try/catch instead of relying on output
    # redirection (which does not reliably suppress it on every PowerShell
    # version when $ErrorActionPreference = 'Stop').
    try { route delete $o.Network mask $o.Mask *> $null } catch {}

    if ($Act -eq 'remove') {
        Start-Sleep -Milliseconds 300
        if (Test-OmadaSaved $o) { W 'Could not remove the route - it may be reserved by another program.' Red }
        else { W "Route $($o.Cidr) removed." Green }
        Show-ExtraSteps @()
        return
    }

    $steps = @()
    $gwPrefix = ($o.Gateway -split '\.')[0..2] -join '.'
    $onNet = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $_.IPAddress -like "$gwPrefix.*" }
    if (-not $onNet) {
        W "Note: this PC isn't on $gwPrefix.x right now. The route is saved anyway." Yellow
        $steps += "Connect this PC to the $gwPrefix.x network - the route then works by itself."
    }

    $addOutput = $null
    try { $addOutput = route -p add $o.Network mask $o.Mask $o.Gateway 2>&1 } catch { $addOutput = $_.Exception.Message }

    # Verify against the actual routing table rather than trusting the exit
    # code or captured text, since both can be unreliable across PowerShell
    # versions for this particular native command.
    Start-Sleep -Milliseconds 300
    if (Test-OmadaSaved $o) {
        W "Permanent route added: $($o.Cidr) via $($o.Gateway)" Green
    } else {
        W "Failed to add route: $addOutput" Red
        return
    }

    if ($onNet) {
        W "Testing $($o.TestIP) ..." Cyan
        if (Test-Connection $o.TestIP -Count 2 -Quiet) {
            W "OK - $($o.TestIP) replies." Green
        } else {
            W "No reply from $($o.TestIP)." Yellow
            $steps += "Check the Omada WAN IN / ACL rule allows $gwPrefix.x, and that $($o.TestIP) is switched on."
        }
    }
    Show-ExtraSteps $steps
}

# =====================================================================
#  SPOTLIGHT - icon and logon task are independent of each other
# =====================================================================
function Update-Desktop {
    if (-not ('TK.Shell' -as [type])) {
        Add-Type -Namespace TK -Name Shell -MemberDefinition '[DllImport("shell32.dll")] public static extern void SHChangeNotify(int e, uint f, IntPtr a, IntPtr b);'
    }
    [TK.Shell]::SHChangeNotify(0x08000000, 0, [IntPtr]::Zero, [IntPtr]::Zero)
}

function Test-SpotTaskExists {
    try { return [bool](Get-ScheduledTask -TaskName $SpotTask -ErrorAction SilentlyContinue) } catch { return $false }
}

function Get-SpotStateText {
    $v = $null
    try { $v = (Get-ItemProperty -Path $SpotKey -Name $SpotClsid -ErrorAction SilentlyContinue).$SpotClsid } catch {}
    $icon = if ($v -eq 1) { 'HIDDEN' } else { 'VISIBLE' }
    $t = if (Test-SpotTaskExists) { 'ON' } else { 'OFF' }
    return "Icon: $icon    Logon task: $t"
}

function Set-SpotHidden {
    if (-not (Test-Path $SpotKey)) { New-Item -Path $SpotKey -Force | Out-Null }
    New-ItemProperty -Path $SpotKey -Name $SpotClsid -Value 1 -PropertyType DWord -Force | Out-Null
    Update-Desktop
}

function Invoke-Spotlight([string]$Act) {
    switch ($Act) {
        'hide' {
            Set-SpotHidden
            W 'Spotlight desktop icon hidden (Spotlight wallpapers keep working).' Green
            Show-ExtraSteps @('If the icon is still visible, click the desktop and press F5.')
        }
        'restore' {
            Remove-ItemProperty -Path $SpotKey -Name $SpotClsid -ErrorAction SilentlyContinue
            Update-Desktop
            W 'Spotlight desktop icon restored.' Green
            if (Test-SpotTaskExists) {
                W "Note: the Spotlight logon task is still ON - it will hide this icon" Yellow
                W "again at your next sign-in. Turn it off from 'Spotlight logon task' if" Yellow
                W "you want the icon to stay visible." Yellow
            }
            Show-ExtraSteps @('If the icon does not show up, click the desktop and press F5.')
        }
        'task-on' {
            if (-not (Test-IsAdmin)) {
                $ok = Invoke-Elevated @('spotlight', 'task-on')
                if ($ok) { W ''; W (Get-SpotStateText) Cyan }
                return
            }
            # Only registers the scheduled task - it does NOT hide the icon
            # right now. Hiding immediately isn't needed for the task to
            # register correctly; the task's own reg.exe action is what
            # hides it, starting at the next sign-in.
            $regArgs = 'add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel" /v ' + $SpotClsid + ' /t REG_DWORD /d 1 /f'
            $user       = "$env:USERDOMAIN\$env:USERNAME"
            $taskAction = New-ScheduledTaskAction -Execute 'reg.exe' -Argument $regArgs
            $trigger    = New-ScheduledTaskTrigger -AtLogOn -User $user
            $settings   = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 1)
            Register-ScheduledTask -TaskName $SpotTask -Action $taskAction -Trigger $trigger -Settings $settings `
                -Description 'Keeps the Spotlight "Learn about this picture" desktop icon hidden' -Force | Out-Null
            W 'Logon task added - it hides the icon at every sign-in, starting next time.' Green
            W 'The task re-runs reg.exe directly; it does not depend on this file.' DarkGray
            W 'A console window may flash for a moment at sign-in - that is the task.' DarkGray
            Show-ExtraSteps @('To also hide the icon right now, use "Hide Spotlight desktop icon" separately.')
        }
        'task-off' {
            if (-not (Test-IsAdmin) -and (Test-SpotTaskExists)) {
                $ok = Invoke-Elevated @('spotlight', 'task-off')
                if ($ok) { W ''; W (Get-SpotStateText) Cyan }
                return
            }
            if (Test-SpotTaskExists) {
                Unregister-ScheduledTask -TaskName $SpotTask -Confirm:$false
                W 'Logon task removed.' Green
            } else {
                W 'No logon task found.' Yellow
            }
            Show-ExtraSteps @()
        }
    }
}

# =====================================================================
#  CLAUDE MODELS
# =====================================================================
function Get-BackupTimestamp([string]$FileName) {
    if ($FileName -match '\.(?:bak|replaced)-(\d{8})-(\d{6})$') {
        try { return [datetime]::ParseExact($Matches[1] + $Matches[2], 'yyyyMMddHHmmss', $null) } catch {}
    }
    return [datetime]::MinValue
}

function Get-ClaudeActiveConfig {
    if (-not (Test-Path -LiteralPath $ClaudeLib)) { return $null }
    $configs = @(Get-ChildItem -LiteralPath $ClaudeLib -File | Where-Object { $_.Extension -eq '.json' -and $_.Name -ne '_meta.json' })
    if ($configs.Count -eq 0) { return $null }
    $meta = Join-Path $ClaudeLib '_meta.json'
    if (Test-Path -LiteralPath $meta) {
        $metaText = Get-Content -LiteralPath $meta -Raw
        foreach ($c in $configs) { if ($metaText -match [regex]::Escape($c.BaseName)) { return $c } }
    }
    return $configs | Sort-Object LastWriteTime -Descending | Select-Object -First 1
}

function Get-ClaudeBackups {
    if (-not (Test-Path -LiteralPath $ClaudeLib)) { return @() }
    return @(Get-ChildItem -LiteralPath $ClaudeLib -File |
             Where-Object { $_.Name -match '\.json\.(bak|replaced)-\d{8}-\d{6}$' } |
             Sort-Object { Get-BackupTimestamp $_.Name } -Descending)
}

function Get-ClaudeStateText {
    $t = Get-ClaudeActiveConfig
    if (-not $t) { return 'Claude Desktop config not found.' }
    $cfg = Get-Content -LiteralPath $t.FullName -Raw | ConvertFrom-Json
    $m = '(app default)'; $e = '(app default)'
    if ($cfg.PSObject.Properties.Match('inferenceModels').Count -gt 0 -and @($cfg.inferenceModels).Count -gt 0) { $m = [string]@($cfg.inferenceModels)[0] }
    if ($cfg.PSObject.Properties.Match('defaultModelEffort').Count -gt 0) { $e = [string]$cfg.defaultModelEffort }
    return "Current default: $m at $e effort    Backups: $(@(Get-ClaudeBackups).Count)"
}

function Confirm-ClaudeClosed {
    if (-not (Get-Process -Name 'Claude' -ErrorAction SilentlyContinue)) { return $true }
    W 'Claude Desktop is still running. Quit it fully first (tray icon > Quit),' Yellow
    W 'otherwise it may overwrite the change when it closes.' Yellow
    if (-not $script:Interactive) { return $true }
    Write-Host -NoNewline '  Enter = continue anyway, Esc = back: '
    while ($true) {
        $k = Read-Key
        if ($k.Key -eq 'Escape') { W ''; return $false }
        if ($k.Key -eq 'Enter')  { W ''; return $true }
    }
}

function Get-ClaudeAfterSteps {
    param([string]$Check)
    return @('Quit Claude Desktop completely (tray icon > Quit).', 'Reopen Claude Desktop.', $Check)
}

function Invoke-Claude([string]$Act) {
    if (-not (Test-Path -LiteralPath $ClaudeLib)) {
        W "Claude Desktop config folder not found: $ClaudeLib" Red
        W 'Install Claude Desktop and open it once, then try again.'
        return
    }
    W "Config folder: $ClaudeLib" DarkGray

    if ($Act -eq 'set') {
        $target = Get-ClaudeActiveConfig
        if (-not $target) { W 'No config file found. Open Claude Desktop once to create one.' Red; return }
        if (-not (Confirm-ClaudeClosed)) { return }

        $backup = "$($target.FullName).bak-$(Get-Date -Format yyyyMMdd-HHmmss)"
        Copy-Item -LiteralPath $target.FullName -Destination $backup
        W "Backup saved: $(Split-Path $backup -Leaf)" Cyan

        $cfg = Get-Content -LiteralPath $target.FullName -Raw | ConvertFrom-Json
        $alwaysDefault = Resolve-Bool $AlwaysUseDefault
        foreach ($pair in @(@('inferenceModels', [object[]]$Models), @('defaultModelEffort', $DefaultEffort), @('alwaysStartWithDefaultModel', $alwaysDefault))) {
            if ($cfg.PSObject.Properties.Match($pair[0]).Count -gt 0) { $cfg.($pair[0]) = $pair[1] }
            else { $cfg | Add-Member -MemberType NoteProperty -Name $pair[0] -Value $pair[1] }
        }
        $json = $cfg | ConvertTo-Json -Depth 50
        [IO.File]::WriteAllText($target.FullName, $json, (New-Object Text.UTF8Encoding $false))

        W ''
        W '[OK] Claude Desktop config updated.' Green
        W "  Default model:  $($Models[0]) at $DefaultEffort effort"
        W "  Model picker limited to exactly $($Models.Count) models: $($Models -join ', ')"
        W "  Always start on default: $alwaysDefault"
        Show-ExtraSteps (Get-ClaudeAfterSteps "Open a new chat and check it starts on $($Models[0]) ($DefaultEffort).")
        W 'To undo: Claude models > Undo last change.' DarkGray
        return
    }

    # ----- revert / undo -----
    $backups = @(Get-ClaudeBackups)
    if ($backups.Count -eq 0) {
        W 'No backups found. Nothing to revert.' Yellow
        Show-ExtraSteps @()
        return
    }

    W ''
    W 'Backups (newest first):' Cyan
    W '  #    Date/time            Type           Size      Config' DarkGray
    for ($i = 0; $i -lt $backups.Count; $i++) {
        $b    = $backups[$i]
        $ts   = (Get-BackupTimestamp $b.Name).ToString('yyyy-MM-dd HH:mm:ss')
        $kind = if ($b.Name -match '\.replaced-') { 'before-revert' } else { 'before-set' }
        $size = if ($b.Length -ge 1KB) { '{0:N0} KB' -f ($b.Length / 1KB) } else { "$($b.Length) B" }
        $note = if ($i -eq 0) { '   <-- NEWEST' } else { '' }
        $line = '  [{0}]  {1}  {2,-13}  {3,-8}  {4}{5}' -f $i, $ts, $kind, $size, $b.Name.Substring(0, [Math]::Min(8, $b.Name.Length)), $note
        if ($i -eq 0) { W $line Green } else { W $line }
    }
    W ''

    if ($Act -eq 'undo') {
        $index = 0
        W "Restoring the newest backup [0]." Cyan
    } else {
        $choice = Read-Input 'Backup number to restore (Esc = back)' '0'
        if ($null -eq $choice) { W 'Cancelled. Nothing changed.' Yellow; return }
        $index = -1
        if (-not [int]::TryParse($choice, [ref]$index) -or $index -lt 0 -or $index -ge $backups.Count) {
            W "Invalid choice: $choice. Nothing changed." Red
            return
        }
    }

    if (-not (Confirm-ClaudeClosed)) { return }

    $picked   = $backups[$index]
    $original = Join-Path $ClaudeLib ($picked.Name -replace '\.(bak|replaced)-\d{8}-\d{6}$', '')
    if (Test-Path -LiteralPath $original) {
        $replaced = "$original.replaced-$(Get-Date -Format yyyyMMdd-HHmmss)"
        Copy-Item -LiteralPath $original -Destination $replaced
        W "Current config saved as: $(Split-Path $replaced -Leaf)" Cyan
    }
    Copy-Item -LiteralPath $picked.FullName -Destination $original -Force

    W ''
    W "[OK] Restored $(Split-Path $original -Leaf)" Green
    W "     from $($picked.Name)" DarkGray
    Show-ExtraSteps (Get-ClaudeAfterSteps 'Open a new chat and check the model/effort is what you expect.')
    W 'Changed your mind? Run Undo last change again - it restores the copy just saved.' DarkGray
}

# =====================================================================
#  THIS PC FOLDERS + SHORTCUT
# =====================================================================
function Get-SpecialFolderPath([string]$Name) {
    switch ($Name) {
        'Documents' { return [Environment]::GetFolderPath('MyDocuments') }
        'Music'     { return [Environment]::GetFolderPath('MyMusic') }
        'Pictures'  { return [Environment]::GetFolderPath('MyPictures') }
        'Videos'    { return [Environment]::GetFolderPath('MyVideos') }
        'Desktop'   { return [Environment]::GetFolderPath('Desktop') }
        'Downloads' {
            # No .NET constant for Downloads (it can be relocated), so ask the
            # shell for its real current path instead of assuming %USERPROFILE%\Downloads.
            try {
                $p = (New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path
                if ($p) { return $p }
            } catch {}
            return (Join-Path $env:USERPROFILE 'Downloads')
        }
    }
    return $null
}

# Quick access pin state/pin/unpin use the Shell.Application COM "pintohome"
# and "unpinfromhome" verbs - the same thing the right-click menu uses. This
# needs no admin rights and never touches the registry, so it's a completely
# separate code path from the This-PC-main-canvas NameSpace toggle above.
function Test-QuickAccessPinned([string]$Path) {
    try {
        $qa = (New-Object -ComObject Shell.Application).Namespace($QuickAccessNamespace)
        foreach ($item in $qa.Items()) { if ($item.Path -eq $Path) { return $true } }
    } catch {}
    return $false
}

function Invoke-QuickAccessPin([string]$Path, [string]$Act) {
    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return $false }
    try {
        if ($Act -eq 'pin') {
            if (Test-QuickAccessPinned $Path) { return $true }
            (New-Object -ComObject Shell.Application).Namespace($Path).Self.InvokeVerb('pintohome')
        } else {
            $qa = (New-Object -ComObject Shell.Application).Namespace($QuickAccessNamespace)
            foreach ($item in $qa.Items()) { if ($item.Path -eq $Path) { $item.InvokeVerb('unpinfromhome'); break } }
        }
        return $true
    } catch { return $false }
}

function Set-QuickAccessGroup([string[]]$Names, [string]$Act) {
    foreach ($name in $Names) {
        $p = Get-SpecialFolderPath $name
        if (-not $p) { W "  $name - could not resolve its folder path." Yellow; continue }
        if (Invoke-QuickAccessPin $p $Act) { W "  $name -> $(if ($Act -eq 'pin') { 'pinned' } else { 'unpinned' })" Green }
        else { W "  $name - failed (folder may not exist, or Quick access list is unavailable)." Red }
    }
    Show-ExtraSteps @('If the sidebar does not update right away, click any folder in Explorer once to refresh it.')
}

function Get-QuickAccessStateText([string[]]$Names) {
    $pinned = 0
    foreach ($name in $Names) {
        $p = Get-SpecialFolderPath $name
        if ($p -and (Test-QuickAccessPinned $p)) { $pinned++ }
    }
    return "Pinned to Quick access: $pinned of $($Names.Count)"
}

function Test-QuickAccessAllPinned([string[]]$Names) {
    foreach ($name in $Names) {
        $p = Get-SpecialFolderPath $name
        if (-not ($p -and (Test-QuickAccessPinned $p))) { return $false }
    }
    return $true
}

function Invoke-QuickAccessToggle([string[]]$Names) {
    # Switches the whole group together: if every folder in the group is
    # already pinned, unpin all of them; otherwise pin whichever are missing.
    $act = if (Test-QuickAccessAllPinned $Names) { 'unpin' } else { 'pin' }
    Set-QuickAccessGroup $Names $act
}

function Get-ThisPcStateText {
    $shown = 0
    foreach ($clsid in $ThisPcClsids.Values) {
        $k = "$NameSpaceRoot\$clsid"
        $v = $null
        try { $v = (Get-ItemProperty -Path $k -Name 'HiddenByDefault' -ErrorAction SilentlyContinue).HiddenByDefault } catch {}
        if ($v -eq 0) { $shown++ }
    }
    return "Shown in sidebar / This PC: $shown of $($ThisPcClsids.Count) folders"
}

function Test-ThisPcAllShown {
    foreach ($clsid in $ThisPcClsids.Values) {
        $k = "$NameSpaceRoot\$clsid"
        $v = $null
        try { $v = (Get-ItemProperty -Path $k -Name 'HiddenByDefault' -ErrorAction SilentlyContinue).HiddenByDefault } catch {}
        if ($v -ne 0) { return $false }
    }
    return $true
}

function Invoke-ThisPcFoldersToggle {
    $act = if (Test-ThisPcAllShown) { 'hide' } else { 'show' }
    Invoke-ThisPcFolders $act
}

function Invoke-ThisPcFolders([string]$Act) {
    if (-not (Test-IsAdmin)) {
        $ok = Invoke-Elevated @('thispc', $Act)
        if ($ok) { W ''; W (Get-ThisPcStateText) Cyan }
        return
    }
    $value = if ($Act -eq 'show') { 0 } else { 1 }
    foreach ($name in $ThisPcClsids.Keys) {
        $k = "$NameSpaceRoot\$($ThisPcClsids[$name])"
        try {
            if (-not (Test-Path $k)) { W "  Skipped $name - registry key not found on this PC." Yellow; continue }
            New-ItemProperty -Path $k -Name 'HiddenByDefault' -PropertyType DWord -Value $value -Force | Out-Null
            W "  $name -> $(if ($value -eq 0) { 'shown' } else { 'hidden' })" Green
        } catch { W "  $name - failed: $($_.Exception.Message)" Red }
    }
    # Keep Desktop at the Windows default (hidden) - it's not in $ThisPcClsids
    # on purpose (see the comment near $DesktopClsid above).
    try {
        $dk = "$NameSpaceRoot\$DesktopClsid"
        if (Test-Path $dk) { New-ItemProperty -Path $dk -Name 'HiddenByDefault' -PropertyType DWord -Value 1 -Force | Out-Null }
    } catch {}

    try { Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue } catch {}
    Start-Sleep -Milliseconds 500
    try { Start-Process explorer.exe } catch {}
    W ''
    if ($Act -eq 'show') { W 'Explorer restarted - the folders should now be in the sidebar and under This PC.' Green }
    else { W 'Explorer restarted - the folders are hidden from the sidebar and This PC again.' Green }
    Show-ExtraSteps @('If a folder is still missing, sign out and back in.')
}

function Invoke-ThisPcShortcut([int]$Idx) {
    $path = Join-Path ([Environment]::GetFolderPath('Desktop')) 'This PC.lnk'
    $ws   = New-Object -ComObject WScript.Shell
    $lnk  = $ws.CreateShortcut($path)
    $lnk.TargetPath    = Join-Path $env:WINDIR 'explorer.exe'
    $lnk.Arguments     = 'shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}'
    $lnk.IconLocation  = (Join-Path $env:WINDIR 'System32\imageres.dll') + ",$Idx"
    $lnk.Save()

    W 'Shortcut created on the Desktop: "This PC.lnk"' Green
    Show-ExtraSteps @()
}

# =====================================================================
#  WEB SEARCH IN TASKBAR
# =====================================================================
function Get-WebSearchStateText {
    $v = $null
    try { $v = (Get-ItemProperty -Path $WebSearchKey -Name 'DisableSearchBoxSuggestions' -ErrorAction SilentlyContinue).DisableSearchBoxSuggestions } catch {}
    return "Web results in taskbar/Start search: $(if ($v -eq 1) { 'OFF (removed)' } else { 'ON (default)' })"
}

function Invoke-WebSearch([string]$Act) {
    # HKCU\Software\Policies is locked by Windows to admins only, even though
    # it's inside your own user hive - a normal window gets "Access to the
    # registry key ... is denied". So this asks for admin like Omada/This PC.
    if (-not (Test-IsAdmin)) {
        $ok = Invoke-Elevated @('websearch', $Act)
        if ($ok) { W ''; W (Get-WebSearchStateText) Cyan }
        return
    }
    if ($Act -eq 'remove') {
        if (-not (Test-Path $WebSearchKey)) { New-Item -Path $WebSearchKey -Force | Out-Null }
        New-ItemProperty -Path $WebSearchKey -Name 'DisableSearchBoxSuggestions' -PropertyType DWord -Value 1 -Force | Out-Null
        W 'Web results removed from the taskbar/Start search.' Green
    } else {
        Remove-ItemProperty -Path $WebSearchKey -Name 'DisableSearchBoxSuggestions' -ErrorAction SilentlyContinue
        W 'Web results restored in the taskbar/Start search.' Green
    }
    Show-ExtraSteps @('Sign out and back in (or restart Explorer) for the change to show.')
}

# =====================================================================
#  SMB USER - PRINT-ONLY. Earlier versions executed hide/unhide
#  (registry write) and delete (Remove-LocalUser) directly and both got
#  flagged by Bitdefender, independently of each other and independently
#  of account creation. This toolkit no longer runs ANY of the four SMB
#  actions itself - it only prints the exact command for you to run
#  yourself in your own elevated window. Nothing here touches the
#  registry or any account, so there is nothing left for an AV heuristic
#  to react to.
# =====================================================================
function Get-SmbUserList {
    # Plain "net user" with no arguments is a read-only, extremely common
    # sysadmin command - just here to let you pick an existing name from a
    # list instead of typing it blind for hide/unhide/delete. If parsing
    # ever comes back empty (different Windows language, etc.) the menu
    # below falls back to typing the name directly, so this is never a
    # hard requirement.
    try {
        $lines = @(net user 2>$null)
        $dashIdx = -1
        for ($i = 0; $i -lt $lines.Count; $i++) { if ($lines[$i] -match '^-+$') { $dashIdx = $i; break } }
        if ($dashIdx -lt 0) { return @() }
        $names = @()
        for ($i = $dashIdx + 1; $i -lt $lines.Count; $i++) {
            $line = $lines[$i]
            if (-not $line -or $line.Trim() -eq '' -or $line -match 'command completed') { continue }
            $names += ($line -split '\s{2,}' | Where-Object { $_.Trim() -ne '' })
        }
        return @($names | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique)
    } catch { return @() }
}

function Select-SmbUser {
    param([string]$Prompt = 'Exact account name')
    $names = @(Get-SmbUserList)
    if ($names.Count -eq 0) { return Read-Input "$Prompt (Esc = back)" '' }
    $items = @()
    for ($i = 0; $i -lt [Math]::Min($names.Count, 9); $i++) { $items += "$($i + 1)|$($names[$i])" }
    $items += 'O|Type a different name'
    $c = Show-Menu -Title 'Choose an account' -Items $items
    if ($null -eq $c) { return $null }
    if ($c -eq 'O') { return Read-Input "$Prompt (Esc = back)" '' }
    return $names[[int]$c - 1]
}

function Show-SmbCommand([string]$Act, [string]$User, [string]$Password) {
    W ''
    W 'Nothing was changed. Copy the command below into an elevated' Yellow
    W 'Command Prompt or PowerShell window and press Enter:' Yellow
    W ''
    switch ($Act) {
        'create' {
            W "net user ""$User"" ""$Password"" /add" Cyan
            W "wmic useraccount where ""name='$User'"" set PasswordExpires=false" Cyan
        }
        'hide' {
            W "reg add ""HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\SpecialAccounts\UserList"" /v ""$User"" /t REG_DWORD /d 0 /f" Cyan
        }
        'unhide' {
            W "reg delete ""HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\SpecialAccounts\UserList"" /v ""$User"" /f" Cyan
        }
        'delete' {
            W "net user ""$User"" /delete" Cyan
        }
    }
    Show-ExtraSteps @('Right-click Start > Terminal (Admin), or Command Prompt (Admin), paste the command above, and press Enter.')
}

# =====================================================================
#  MENUS
# =====================================================================
function Menu-Omada {
    while ($true) {
        $o = Get-OmadaSettings
        $c = Show-Menu -Title 'Omada route' -Items @('1|Add route (permanent)', '2|Remove route') -Info {
            W "  $($o.Cidr) via $($o.Gateway)  -  now: $(Get-OmadaStateText $o)" DarkGray
        }
        if ($null -eq $c) { return }
        W ''
        switch ($c) { '1' { Invoke-Safe { Invoke-Omada 'add' } } '2' { Invoke-Safe { Invoke-Omada 'remove' } } }
        Wait-Back
    }
}

function Menu-Spotlight {
    while ($true) {
        $c = Show-Menu -Title 'Spotlight desktop icon' -Items @(
            '1|Hide Spotlight desktop icon',
            '2|Restore Spotlight desktop icon',
            '3|Add logon task (hides the icon at every sign-in)',
            '4|Remove logon task') -Info { W "  $(Get-SpotStateText)" DarkGray }
        if ($null -eq $c) { return }
        W ''
        $act = @{ '1' = 'hide'; '2' = 'restore'; '3' = 'task-on'; '4' = 'task-off' }[$c]
        Invoke-Safe { Invoke-Spotlight $act }
        Wait-Back
    }
}

function Menu-Claude {
    while ($true) {
        $c = Show-Menu -Title 'Claude models' -Items @(
            "1|Set default ($($Models[0]), $DefaultEffort effort)",
            '2|Revert - choose a backup',
            '3|Undo last change (restore newest backup)') -Info {
            W "  $(Get-ClaudeStateText)" DarkGray
            W "  'Set default' limits the model picker to exactly $($Models.Count) models:" DarkGray
            W "  $($Models -join ', ')" DarkGray
        }
        if ($null -eq $c) { return }
        W ''
        $act = @{ '1' = 'set'; '2' = 'revert'; '3' = 'undo' }[$c]
        Invoke-Safe { Invoke-Claude $act }
        Wait-Back
    }
}

# Downloads is deliberately NOT in Group1 - it's handled together with
# Desktop in Group2 instead, per Ahmed's request.
$QuickAccessGroup1 = @('Documents', 'Music', 'Pictures', 'Videos')
# Quick access sorts its pinned folders alphabetically when nothing has been
# manually drag-reordered, and "Desktop" < "Downloads" alphabetically - so
# listing Desktop first here already puts it above Downloads with no extra
# reordering trick needed.
$QuickAccessGroup2 = @('Desktop', 'Downloads')

function Menu-ThisPc {
    while ($true) {
        $c = Show-Menu -Title 'This PC folders' -Items @(
            '1|This PC main view: show or hide Documents/Downloads/Music/Pictures/Videos',
            '2|Quick access sidebar: un/pin Documents/Music/Pictures/Videos',
            '3|Quick access sidebar: un/pin Desktop + Downloads only',
            '4|Add "This PC" shortcut to the Desktop') -Info {
            W "  Main view - $(Get-ThisPcStateText)" DarkGray
            W "  Quick access (Documents/Music/Pictures/Videos) - $(Get-QuickAccessStateText $QuickAccessGroup1)" DarkGray
            W "  Quick access (Desktop/Downloads)                - $(Get-QuickAccessStateText $QuickAccessGroup2)" DarkGray
        }
        if ($null -eq $c) { return }
        W ''
        switch ($c) {
            '1' { Invoke-Safe { Invoke-ThisPcFoldersToggle }; Wait-Back }
            '2' { Invoke-Safe { Invoke-QuickAccessToggle $QuickAccessGroup1 }; Wait-Back }
            '3' { Invoke-Safe { Invoke-QuickAccessToggle $QuickAccessGroup2 }; Wait-Back }
            '4' { Invoke-Safe { Invoke-ThisPcShortcut $IconIndex }; Wait-Back }
        }
    }
}

function Menu-WebSearch {
    while ($true) {
        $c = Show-Menu -Title 'Web search in taskbar' -Items @('1|Remove web results', '2|Restore web results') -Info { W "  $(Get-WebSearchStateText)" DarkGray }
        if ($null -eq $c) { return }
        W ''
        Invoke-Safe { Invoke-WebSearch (@{ '1' = 'remove'; '2' = 'restore' }[$c]) }
        Wait-Back
    }
}

function Menu-SmbUser {
    while ($true) {
        $c = Show-Menu -Title 'SMB user (prints a command - nothing is run for you)' -Items @(
            '1|Create an account',
            '2|Hide an account from the sign-in screen',
            '3|Unhide an account',
            '4|Delete an account') -Info { W '  This only prints a command. It never changes anything by itself.' DarkGray }
        if ($null -eq $c) { return }
        W ''
        $pw = ''
        if ($c -eq '1') {
            # Create is for a NEW account, so there's nothing existing to
            # pick from - always type the name here.
            $u = Read-Input 'Exact account name for the new account (Esc = back)' ''
            if ($null -eq $u -or $u -eq '') { W 'Cancelled.' Yellow; Wait-Back; continue }
            $pw = Read-Password 'Password for the new account (Esc = back)'
            if ($null -eq $pw -or $pw -eq '') { W 'Cancelled.' Yellow; Wait-Back; continue }
        } else {
            # Hide/unhide/delete act on an EXISTING account - offer to pick
            # one from "net user" instead of typing it blind.
            $u = Select-SmbUser
            if ($null -eq $u -or $u -eq '') { W 'Cancelled.' Yellow; Wait-Back; continue }
        }
        $act = @{ '1' = 'create'; '2' = 'hide'; '3' = 'unhide'; '4' = 'delete' }[$c]
        Invoke-Safe { Show-SmbCommand $act $u $pw }
        Wait-Back
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
        $c = Show-Menu -Title 'Fresh Setup Toolkit' -EscLabel 'Exit' -Items @(
            '1|Omada route            add / remove',
            '2|Spotlight desktop icon hide/restore icon, add/remove logon task',
            '3|Claude models          set / revert (3 fixed models)',
            '4|This PC folders        main view + Quick access + shortcut',
            '5|Web search taskbar     remove / restore',
            '6|SMB user               create/hide/unhide/delete (prints commands only)',
            'H|Help')
        if ($null -eq $c) { return }
        switch ($c) {
            '1' { Menu-Omada }
            '2' { Menu-Spotlight }
            '3' { Menu-Claude }
            '4' { Menu-ThisPc }
            '5' { Menu-WebSearch }
            '6' { Menu-SmbUser }
            'H' { Menu-Help }
        }
    }
}

# =====================================================================
#  ENTRY POINT
# =====================================================================
if ($Help -or $Tool -match '^(help|/\?|/h|--help)$') {
    # Note: "-?" is NOT in this list on purpose - PowerShell intercepts
    # that exact token for ANY script before it ever reaches here and
    # prints its own generic parameter-syntax line instead, regardless of
    # what this script does. Confirmed by direct testing. -h, help, /?,
    # /h, and --help all work; -? never will, so don't advertise it.
    Show-About
    Show-Parameters
    W ''
    exit 0
}

if ($Tool) {
    # Command-line mode: run one action and exit
    $ok = $true
    switch ($Tool.ToLower()) {
        'omada' {
            if ($Action -in 'add', 'remove', 'status') { Invoke-Safe { Invoke-Omada $Action.ToLower() } }
            else { W 'Usage: omada add | remove | status   (-h for help)' Red; $ok = $false }
        }
        'spotlight' {
            if ($Action -in 'hide', 'restore', 'task-on', 'task-off') { Invoke-Safe { Invoke-Spotlight $Action.ToLower() } }
            else { W 'Usage: spotlight hide | restore | task-on | task-off   (-h for help)' Red; $ok = $false }
        }
        'claude' {
            if ($Action -in 'set', 'revert', 'undo') { Invoke-Safe { Invoke-Claude $Action.ToLower() } }
            else { W 'Usage: claude set | revert | undo   (-h for help)' Red; $ok = $false }
        }
        'thispc' {
            if ($Action -in 'show', 'hide') { Invoke-Safe { Invoke-ThisPcFolders $Action.ToLower() } }
            elseif ($Action -eq 'shortcut') { Invoke-Safe { Invoke-ThisPcShortcut $IconIndex } }
            elseif ($Action -eq 'qa-pin')        { Invoke-Safe { Set-QuickAccessGroup $QuickAccessGroup1 'pin' } }
            elseif ($Action -eq 'qa-unpin')      { Invoke-Safe { Set-QuickAccessGroup $QuickAccessGroup1 'unpin' } }
            elseif ($Action -eq 'qa-pin-desk')   { Invoke-Safe { Set-QuickAccessGroup $QuickAccessGroup2 'pin' } }
            elseif ($Action -eq 'qa-unpin-desk') { Invoke-Safe { Set-QuickAccessGroup $QuickAccessGroup2 'unpin' } }
            else { W 'Usage: thispc show | hide | shortcut | qa-pin | qa-unpin | qa-pin-desk | qa-unpin-desk   (-h for help)' Red; $ok = $false }
        }
        'websearch' {
            if ($Action -in 'remove', 'restore') { Invoke-Safe { Invoke-WebSearch $Action.ToLower() } }
            else { W 'Usage: websearch remove | restore   (-h for help)' Red; $ok = $false }
        }
        'smbuser' {
            if ($Action -in 'create', 'hide', 'unhide', 'delete' -and $User -and ($Action -ne 'create' -or $Password)) {
                Invoke-Safe { Show-SmbCommand $Action.ToLower() $User $Password }
            } else { W 'Usage: smbuser create|hide|unhide|delete -User <name> [-Password <pw>]   (-h for help)' Red; $ok = $false }
        }
        default { W "Unknown tool '$Tool'. Run with -h for help." Red; $ok = $false }
    }
    if ($Elevated) { Wait-Back 'Press any key to close this admin window' }
    if ($ok) { exit 0 } else { exit 1 }
}

# No arguments: interactive menu (double-click)
$script:Interactive = $true
Menu-Main
Clear-Host
exit 0
