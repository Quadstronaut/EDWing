# WingCreds.ps1 - min-ed-launcher credential + settings seeding for EDWing sandboxes.
# Dot-source from the launcher:  . "$PSScriptRoot\WingCreds.ps1"
#
# How min-ed-launcher (v0.12.x) actually works, verified on this machine:
#   - It reads settings.json (launcher settings: filterOverrides, autoUpdate, ...).
#   - Per profile label passed as `/frontier <Profile>`, first run does an interactive
#     Frontier login and caches a DPAPI-encrypted `.frontier-<profile>.cred` for
#     auto-login next time. There is NO plaintext ini; the profile is just a label.
#   - Both files live under %LOCALAPPDATA%\min-ed-launcher\.
#
# Sandboxie virtualizes %LOCALAPPDATA%, so each box gets its OWN copy under the box's
# virtual AppData, and it's lost if the box is ever wiped. DPAPI blobs decrypt only
# under the same Windows user + machine - which a Sandboxie box still is - so a token
# minted natively (unsandboxed) can be seeded into a box and should still decrypt.
#
# This module: pins an explicit CMDR<->profile map, reports each box's credential
# state, and seeds a box that lacks a token from the native token first, else a local
# backup - never clobbering a token the box already has.

#Requires -Version 5.1

# CMDR name -> min-ed profile label. Positional (index i -> Account(i+1)), but returned
# as an explicit map so a later $cmdrNames reorder can't silently swap logins.
function Get-CmdrProfileMap {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string[]]$CmdrNames)
    $map = @()
    for ($i = 0; $i -lt $CmdrNames.Count; $i++) {
        $map += [pscustomobject]@{ Box = $CmdrNames[$i]; Profile = "Account$($i + 1)" }
    }
    return $map
}

function Get-CredFileName {
    param([Parameter(Mandatory)][string]$Profile)
    ".frontier-$($Profile.ToLower()).cred"
}

function Get-NativeMinEdDir {
    Join-Path $env:LOCALAPPDATA 'min-ed-launcher'
}

# min-ed's dir inside a box. Two Sandboxie virtualization conventions exist in the wild;
# probe both (autohonk.py already knows this), prefer whichever exists, else convention #1
# for creation. Verified on this machine: convention #1 (user\current\AppData\Local).
function Get-BoxMinEdDir {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SandboxRoot,
        [Parameter(Mandatory)][string]$Box
    )
    $c1 = Join-Path $SandboxRoot "$Box\user\current\AppData\Local\min-ed-launcher"
    $c2 = Join-Path $SandboxRoot "$Box\drive\C\Users\$env:USERNAME\AppData\Local\min-ed-launcher"
    if (Test-Path $c1) { return $c1 }
    if (Test-Path $c2) { return $c2 }
    return $c1   # default for first-time creation
}

# Per-box report: does the box have a token / settings, is a native token or a backup
# available to seed from, and what's the recommended action.
function Test-CmdrCredentialState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$CmdrNames,
        [Parameter(Mandatory)][string]$SandboxRoot,
        [string]$CredBackupDir
    )
    $nativeDir = Get-NativeMinEdDir
    foreach ($entry in (Get-CmdrProfileMap -CmdrNames $CmdrNames)) {
        $cred      = Get-CredFileName -Profile $entry.Profile
        $boxDir    = Get-BoxMinEdDir -SandboxRoot $SandboxRoot -Box $entry.Box
        $boxCred   = Join-Path $boxDir $cred
        $boxJson   = Join-Path $boxDir 'settings.json'
        $nativeCred = Join-Path $nativeDir $cred
        $backupCred = if ($CredBackupDir) { Join-Path $CredBackupDir $cred } else { $null }

        $hasBox    = Test-Path $boxCred
        $hasNative = Test-Path $nativeCred
        $hasBackup = $backupCred -and (Test-Path $backupCred)

        $action =
            if     ($hasBox)    { 'ok (box has token)' }
            elseif ($hasNative) { 'seed from native' }
            elseif ($hasBackup) { 'seed from backup' }
            else                { 'interactive login required' }

        [pscustomobject]@{
            Box         = $entry.Box
            Profile     = $entry.Profile
            BoxToken    = [bool]$hasBox
            BoxSettings = [bool](Test-Path $boxJson)
            Native      = [bool]$hasNative
            Backup      = [bool]$hasBackup
            Action      = $action
        }
    }
}

# Mirror each box's current token into the gitignored backup dir (safety net for wipes).
function Backup-CmdrCredentials {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$CmdrNames,
        [Parameter(Mandatory)][string]$SandboxRoot,
        [Parameter(Mandatory)][string]$CredBackupDir
    )
    if (-not (Test-Path $CredBackupDir)) { New-Item -ItemType Directory -Path $CredBackupDir -Force | Out-Null }
    foreach ($entry in (Get-CmdrProfileMap -CmdrNames $CmdrNames)) {
        $cred    = Get-CredFileName -Profile $entry.Profile
        $boxCred = Join-Path (Get-BoxMinEdDir -SandboxRoot $SandboxRoot -Box $entry.Box) $cred
        if (Test-Path $boxCred) {
            Copy-Item $boxCred (Join-Path $CredBackupDir $cred) -Force
            Write-Host "  backed up $($entry.Box) ($($entry.Profile))"
        }
    }
}

# Seed a box that LACKS a token, preferring the native token, else a backup. Never
# overwrites an existing box token. Also seeds settings.json if the box lacks one.
# Returns the post-seed state objects.
function Restore-CmdrCredentials {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$CmdrNames,
        [Parameter(Mandatory)][string]$SandboxRoot,
        [string]$CredBackupDir,
        [switch]$WhatIfSeed   # report-only, make no changes
    )
    if (-not (Test-Path $SandboxRoot)) {
        Write-Warning "SandboxRoot '$SandboxRoot' does not exist - no boxes to seed. Check your Sandboxie install path."
        return @()
    }
    $nativeDir = Get-NativeMinEdDir
    $nativeJson = Join-Path $nativeDir 'settings.json'

    foreach ($entry in (Get-CmdrProfileMap -CmdrNames $CmdrNames)) {
        $cred     = Get-CredFileName -Profile $entry.Profile
        $boxDir   = Get-BoxMinEdDir -SandboxRoot $SandboxRoot -Box $entry.Box
        $boxCred  = Join-Path $boxDir $cred
        $boxJson  = Join-Path $boxDir 'settings.json'

        # --- token ---
        if (Test-Path $boxCred) {
            Write-Host "  $($entry.Box): token present - leaving as-is"
        }
        else {
            $src = $null; $srcLabel = $null
            $nativeCred = Join-Path $nativeDir $cred
            $backupCred = if ($CredBackupDir) { Join-Path $CredBackupDir $cred } else { $null }
            if (Test-Path $nativeCred)                         { $src = $nativeCred; $srcLabel = 'native' }
            elseif ($backupCred -and (Test-Path $backupCred))  { $src = $backupCred; $srcLabel = 'backup' }

            if ($src) {
                if ($WhatIfSeed) {
                    Write-Host "  $($entry.Box): WOULD seed token from $srcLabel" -ForegroundColor Yellow
                }
                else {
                    if (-not (Test-Path $boxDir)) { New-Item -ItemType Directory -Path $boxDir -Force | Out-Null }
                    Copy-Item $src $boxCred -Force
                    Write-Host "  $($entry.Box): seeded token from $srcLabel" -ForegroundColor Green
                }
            }
            else {
                Write-Warning "  $($entry.Box) ($($entry.Profile)): no native or backup token - needs an interactive first login"
            }
        }

        # --- settings.json (only if absent; don't clobber a box's own) ---
        if (-not (Test-Path $boxJson) -and (Test-Path $nativeJson)) {
            if ($WhatIfSeed) {
                Write-Host "  $($entry.Box): WOULD seed settings.json from native" -ForegroundColor Yellow
            }
            else {
                if (-not (Test-Path $boxDir)) { New-Item -ItemType Directory -Path $boxDir -Force | Out-Null }
                Copy-Item $nativeJson $boxJson -Force
                Write-Host "  $($entry.Box): seeded settings.json from native"
            }
        }
    }
    return (Test-CmdrCredentialState -CmdrNames $CmdrNames -SandboxRoot $SandboxRoot -CredBackupDir $CredBackupDir)
}

# Orchestrator the launcher calls: report -> back up existing -> seed the gaps -> report.
function Invoke-WingCredentialSeeding {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$CmdrNames,
        [Parameter(Mandatory)][string]$SandboxRoot,
        [Parameter(Mandatory)][string]$CredBackupDir,
        [switch]$WhatIfSeed
    )
    Write-Host "Credential state (before):" -ForegroundColor Cyan
    Test-CmdrCredentialState -CmdrNames $CmdrNames -SandboxRoot $SandboxRoot -CredBackupDir $CredBackupDir |
        Format-Table -AutoSize | Out-String | Write-Host

    if (-not $WhatIfSeed) {
        Write-Host "Backing up existing box tokens:" -ForegroundColor Cyan
        Backup-CmdrCredentials -CmdrNames $CmdrNames -SandboxRoot $SandboxRoot -CredBackupDir $CredBackupDir
    }

    Write-Host "Seeding gaps (native -> box, then backup -> box):" -ForegroundColor Cyan
    $after = Restore-CmdrCredentials -CmdrNames $CmdrNames -SandboxRoot $SandboxRoot -CredBackupDir $CredBackupDir -WhatIfSeed:$WhatIfSeed

    Write-Host "Credential state (after):" -ForegroundColor Cyan
    $after | Format-Table -AutoSize | Out-String | Write-Host

    $needLogin = $after | Where-Object { -not $_.BoxToken }
    if ($needLogin) {
        Write-Host ("CMDRs still needing an interactive first login: {0}" -f (($needLogin.Box) -join ', ')) -ForegroundColor Yellow
    } else {
        Write-Host "All boxes have a token (decryption still to be confirmed on first launch)." -ForegroundColor Green
    }
    return $after
}
