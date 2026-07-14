#Requires -Version 5.1
<#
.SYNOPSIS
    EDWing multibox launcher - launches N sandboxed Elite Dangerous clients and arranges them.

.DESCRIPTION
    For each Sandboxie box / CMDR it: seeds min-ed credentials, launches the client,
    waits (bounded, all boxes concurrently) for its window, then arranges the windows:

      stacked  all clients borderless full-size on the primary monitor; you or a bot
               bring the wanted CMDR forward and act, then cycle (mode 2, default).
               Requires each client in Borderless/Windowed display mode (NOT exclusive
               fullscreen - Windows only allows one exclusive owner per monitor).
      tiled    primary CMDR on the main monitor, the rest tiled on a side monitor,
               rects derived from the LIVE monitor list (mode 1).

    A client that doesn't appear in time (e.g. stuck on a Frontier re-auth prompt) is
    flagged and SKIPPED - it never hangs the other three.

    Local overrides live in wing.conf.ps1 (gitignored). See example_configs/.

.EXAMPLE
    .\Get-Wing.ps1                 # stacked (mode 2)
    .\Get-Wing.ps1 -Mode tiled     # classic tiled layout (mode 1)
    .\Get-Wing.ps1 -WhatIf         # seed-report + no launch (dry preview)
#>
[CmdletBinding()]
param(
    [ValidateSet('stacked','tiled')][string]$Mode,
    [switch]$NoCreds,
    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'
$here = $PSScriptRoot

# --- Shared libraries ---
. "$here\WingLib.ps1"     # window identity / focus / positioning primitives
. "$here\WingCreds.ps1"   # min-ed credential + settings seeding

# --- Default configuration (override any of these in wing.conf.ps1) ---
$config = @{
    launchEliteDangerous           = $true
    windowMode                     = 'stacked'   # 'stacked' (mode 2) | 'tiled' (mode 1)
    pgEntry                        = $true        # run per-CMDR menu click sets (needs captured coords)
    launchEDMC                     = $false       # companions normally auto-run via Sandboxie RunCommand
    SeedCredentials                = $true
    StopCustomServicesAndProcesses = $false

    WindowTimeoutSec               = 120          # bounded wait (all boxes concurrently), then flag & continue
    WindowPollMs                   = 2000
    EliteSettleSeconds             = 7            # let renderers settle before positioning
    PgReadyDelaySeconds            = 20           # TODO: replace with a per-CMDR journal readiness signal

    SandboxRoot                    = "C:\Sandbox\$env:USERNAME"
    CredBackupDir                  = "$here\cred_backup"

    # Per-CMDR menu click sets, in PRIMARY-monitor coordinates (mode 2 focuses each
    # window first). Capture with tools\Get-CursorPos.ps1 and set these in wing.conf.ps1.
    # Empty => that CMDR's clicks are skipped (never fires blind clicks at wrong pixels).
    CmdrClickSets                  = @{}
}

# Commander boxes. Order = launch/layout order; first entry is the "primary" CMDR
# (main monitor in tiled mode). Reordering is safe for LAYOUT - credential identity is
# anchored by $cmdrProfiles below, NOT by array position.
$cmdrNames = @(
    'CMDRDuvrazh',
    'CMDRBistronaut',
    'CMDRTristronaut',
    'CMDRQuadstronaut'
)

# CMDR box -> min-ed profile label. Identity-anchored so reordering $cmdrNames for layout
# can never log a box in as a different real Frontier account (review rank 3).
$cmdrProfiles = @{
    'CMDRDuvrazh'      = 'Account1'
    'CMDRBistronaut'   = 'Account2'
    'CMDRTristronaut'  = 'Account3'
    'CMDRQuadstronaut' = 'Account4'
}

# Executable paths (override in wing.conf.ps1).
$sandboxieStart = 'C:\Users\Quadstronaut\scoop\apps\sandboxie-plus-np\current\Start.exe'
$minEDLauncher  = 'G:\SteamLibrary\steamapps\common\Elite Dangerous\MinEdLauncher.exe'
$edmc_path      = 'G:\EliteApps\EDMarketConnector\EDMarketConnector.exe'

# --- Apply local overrides, then command-line params (params win) ---
$wingConf = Join-Path $here 'wing.conf.ps1'
if (Test-Path $wingConf) { Write-Host "Loading $wingConf"; . $wingConf }
if ($Mode)    { $config.windowMode = $Mode }
if ($NoCreds) { $config.SeedCredentials = $false }

# --- Validate executables (only what's enabled) ---
# NB: print the itemized list with Write-Host, not Write-Error - $ErrorActionPreference
# is 'Stop', so a Write-Error here would terminate before the list prints (review rank 11).
$missing = @()
if (-not (Test-Path $sandboxieStart)) { $missing += "Sandboxie Start.exe: $sandboxieStart" }
if ($config.launchEliteDangerous -and -not (Test-Path $minEDLauncher)) { $missing += "MinEdLauncher: $minEDLauncher" }
if ($config.launchEDMC -and -not (Test-Path $edmc_path)) { $missing += "EDMC: $edmc_path" }
if ($missing.Count) {
    Write-Host "Missing required executables:" -ForegroundColor Red
    $missing | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
    return
}

# --- Tiled layout helper (mode 1): primary CMDR on main monitor, rest on a side monitor,
#     with rects derived from the LIVE monitor list instead of baked pixel constants. ---
function Set-WingTiledLayout {
    param(
        [Parameter(Mandatory)][object[]]$Ready,
        [Parameter(Mandatory)][string]$PrimaryCmdr
    )
    $layout  = Get-WingMonitorLayout
    $primary = $layout | Where-Object Primary | Select-Object -First 1
    $side    = $layout | Where-Object { -not $_.Primary } | Sort-Object X | Select-Object -First 1  # leftmost non-primary

    $primWin = $Ready | Where-Object { $_.Box -eq $PrimaryCmdr } | Select-Object -First 1
    if ($primWin) {
        Set-CmdrWindowRect -Hwnd $primWin.Hwnd -X $primary.X -Y $primary.Y -Width $primary.Width -Height $primary.Height | Out-Null
        Write-Host "  $($primWin.Box) -> main monitor (full)"
    }

    $others = @($Ready | Where-Object { $_.Box -ne $PrimaryCmdr })
    if (-not $side) { Write-Warning "  No side monitor found; alt clients left unplaced."; return }
    if (-not $others) { return }
    $rowH = [int]($side.Height / $others.Count)
    for ($k = 0; $k -lt $others.Count; $k++) {
        $y = $side.Y + $k * $rowH
        Set-CmdrWindowRect -Hwnd $others[$k].Hwnd -X $side.X -Y $y -Width $side.Width -Height $rowH -Borderless | Out-Null
        Write-Host ("  {0} -> side {1},{2} {3}x{4}" -f $others[$k].Box, $side.X, $y, $side.Width, $rowH)
    }
}

# --- Optionally free resources ---
if ($config.StopCustomServicesAndProcesses -and -not $WhatIf) {
    Get-Process -Name '*discord*','*ollama*' -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}

# --- Credential phase ---
if ($config.SeedCredentials) {
    Write-Host "`n=== Credentials ===" -ForegroundColor Cyan
    Invoke-WingCredentialSeeding -CmdrNames $cmdrNames -SandboxRoot $config.SandboxRoot `
        -CredBackupDir $config.CredBackupDir -ProfileMap $cmdrProfiles -WhatIfSeed:$WhatIf | Out-Null
}

if ($WhatIf) { Write-Host "`n-WhatIf: stopping before launch." -ForegroundColor Yellow; return }

# --- Launch phase ---
if ($config.launchEliteDangerous) {
    Write-Host "`n=== Launching $($cmdrNames.Count) sandboxed clients ($($config.windowMode)) ===" -ForegroundColor Cyan
    for ($i = 0; $i -lt $cmdrNames.Count; $i++) {
        $box         = $cmdrNames[$i]
        $profileName = if ($cmdrProfiles.ContainsKey($box)) { $cmdrProfiles[$box] } else { "Account$($i + 1)" }
        $sbArgs      = "/box:$box `"$minEDLauncher`" /frontier $profileName /edo /autorun /autoquit /skipInstallPrompt"
        Start-Process -FilePath $sandboxieStart -ArgumentList $sbArgs
        Write-Host "  launched $box ($profileName)"
        if ($config.launchEDMC) {
            Start-Process -FilePath $sandboxieStart -ArgumentList "/box:$box `"$edmc_path`""
        }
    }
}

# --- Detect (bounded, all boxes concurrently) + arrange ---
if ($config.launchEliteDangerous) {
    Write-Host "`n=== Waiting for windows (<= $($config.WindowTimeoutSec)s, all boxes) ===" -ForegroundColor Cyan
    $result   = Wait-AllCmdrWindows -BoxNames $cmdrNames -TimeoutSec $config.WindowTimeoutSec -PollMs $config.WindowPollMs
    $ready    = @($result.Ready)
    $notReady = @($result.NotReady)
    foreach ($b in $notReady) {
        Write-Warning "  $b did not appear in $($config.WindowTimeoutSec)s - likely stuck at login / needs a re-auth code. Continuing with the rest."
    }

    if ($ready) {
        Start-Sleep -Seconds $config.EliteSettleSeconds
        Write-Host "`n=== Arranging ($($config.windowMode)) ===" -ForegroundColor Cyan
        switch ($config.windowMode) {
            'stacked' {
                Write-Host "  NOTE: mode 2 assumes each client is in Borderless/Windowed display mode (not exclusive fullscreen)." -ForegroundColor DarkYellow
                $p = Get-WingPrimaryRect
                Write-Host "  stacking $($ready.Count) clients on primary ($($p.Width)x$($p.Height))"
                foreach ($w in $ready) {
                    $ok = Set-CmdrWindowRect -Hwnd $w.Hwnd -X $p.X -Y $p.Y -Width $p.Width -Height $p.Height -Borderless
                    Write-Host ("    {0}: {1}" -f $w.Box, $(if ($ok) { 'placed' } else { 'FAILED' }))
                }
            }
            'tiled' { Set-WingTiledLayout -Ready $ready -PrimaryCmdr $cmdrNames[0] }
        }
    }

    # --- Menu automation (cutscene skip + Continue/PG/Launch), mode-2 style ---
    # For each ready CMDR: focus its window, then run its captured click set. Coords are
    # display-specific; a box with no captured set is skipped (never fires blind clicks).
    # Before EACH click we re-verify the intended window is still foreground and abort that
    # CMDR's sequence if another window stole focus (review rank 5) - critical in stacked
    # mode where all windows share the same rect.
    if ($config.pgEntry -and $ready) {
        Write-Host "`n=== Menu automation ===" -ForegroundColor Cyan
        . "$here\clicker_scripts\MouseUtil.ps1"
        Start-Sleep -Seconds $config.PgReadyDelaySeconds   # TODO: gate per-CMDR on a journal readiness signal
        foreach ($w in $ready) {
            $set = $config.CmdrClickSets[$w.Box]
            if (-not $set) {
                Write-Warning "  $($w.Box): no click set captured - skipping (see tools\Get-CursorPos.ps1)"
                continue
            }
            if (-not (Set-CmdrForeground -Hwnd $w.Hwnd)) {
                Write-Warning "  $($w.Box): could not bring to foreground - skipping its clicks"
                continue
            }
            Start-Sleep -Milliseconds 400
            $aborted = $false
            foreach ($c in $set) {
                if ((Get-ForegroundHwnd) -ne $w.Hwnd) {
                    Write-Warning "  $($w.Box): foreground was stolen mid-sequence - aborting its remaining clicks"
                    $aborted = $true; break
                }
                Invoke-ClickAction -X $c.X -Y $c.Y -ClickType $(if ($c.ClickType) { $c.ClickType } else { 'Double' })
            }
            if (-not $aborted) { Write-Host "  $($w.Box): ran $($set.Count) clicks" }
        }
    }

    Write-Host "`n=== Summary ===" -ForegroundColor Cyan
    Write-Host ("  ready:  {0}" -f (($ready.Box) -join ', '))
    if ($notReady) { Write-Host ("  needs attention (login/re-auth): {0}" -f ($notReady -join ', ')) -ForegroundColor Yellow }
    Write-Host "`nTip: bring a CMDR forward any time with:  .\tools\Show-Cmdr.ps1 CMDRBistronaut" -ForegroundColor DarkGray
}
