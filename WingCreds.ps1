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
# This module: pins an explicit CMDR<->profile map (identity-anchored, NOT array
# position), reports each box's credential state, and seeds a box that lacks a valid
# token from the native token first, else a local backup - never clobbering a token
# the box already has, and never letting a corrupt/short token masquerade as valid.

#Requires -Version 5.1

# A real min-ed token is ~962 bytes; anything shorter is truncated/corrupt (review rank 14).
$script:MinCredBytes = 100

# CMDR name -> min-ed profile label. Pass an explicit identity map (@{CMDRDuvrazh='Account1';...})
# so reordering $cmdrNames for layout can NEVER remap which real account a box logs in as
# (review rank 3). Falls back to positional only if no map is supplied, and warns per gap.
function Get-CmdrProfileMap {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$CmdrNames,
        [hashtable]$ProfileMap
    )
    $out = @()
    for ($i = 0; $i -lt $CmdrNames.Count; $i++) {
        $box = $CmdrNames[$i]
        if ($ProfileMap -and $ProfileMap.ContainsKey($box)) {
            $profile = $ProfileMap[$box]
        }
        else {
            $profile = "Account$($i + 1)"
            if ($ProfileMap) { Write-Warning "No profile pinned for $box; falling back to positional '$profile'" }
        }
        $out += [pscustomobject]@{ Box = $box; Profile = $profile }
    }
    return $out
}

function Get-CredFileName {
    param([Parameter(Mandatory)][string]$Profile)
    ".frontier-$($Profile.ToLower()).cred"
}

function Get-NativeMinEdDir {
    Join-Path $env:LOCALAPPDATA 'min-ed-launcher'
}

# A cred file counts as valid only if it exists AND is at least MinCredBytes (review rank 14).
function Test-CredValid {
    param([string]$Path)
    if (-not $Path -or -not (Test-Path $Path)) { return $false }
    return ((Get-Item $Path).Length -ge $script:MinCredBytes)
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

# Per-box report: does the box have a VALID token / settings, is a native token or a backup
# available to seed from, and what's the recommended action.
function Test-CmdrCredentialState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$CmdrNames,
        [Parameter(Mandatory)][string]$SandboxRoot,
        [string]$CredBackupDir,
        [hashtable]$ProfileMap
    )
    $nativeDir = Get-NativeMinEdDir
    foreach ($entry in (Get-CmdrProfileMap -CmdrNames $CmdrNames -ProfileMap $ProfileMap)) {
        $cred      = Get-CredFileName -Profile $entry.Profile
        $boxDir    = Get-BoxMinEdDir -SandboxRoot $SandboxRoot -Box $entry.Box
        $boxCred   = Join-Path $boxDir $cred
        $boxJson   = Join-Path $boxDir 'settings.json'
        $nativeCred = Join-Path $nativeDir $cred
        $backupCred = if ($CredBackupDir) { Join-Path $CredBackupDir $cred } else { $null }

        $hasBox    = Test-CredValid $boxCred
        $hasNative = Test-CredValid $nativeCred
        $hasBackup = Test-CredValid $backupCred

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

# Mirror each box's current VALID token into the gitignored backup dir (safety net for
# wipes). Never overwrite an existing backup with a shorter/corrupt box file (review rank 14).
# One box's filesystem error is warned and skipped, never fatal (review rank 4).
function Backup-CmdrCredentials {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$CmdrNames,
        [Parameter(Mandatory)][string]$SandboxRoot,
        [Parameter(Mandatory)][string]$CredBackupDir,
        [hashtable]$ProfileMap
    )
    if (-not (Test-Path $CredBackupDir)) { New-Item -ItemType Directory -Path $CredBackupDir -Force | Out-Null }
    foreach ($entry in (Get-CmdrProfileMap -CmdrNames $CmdrNames -ProfileMap $ProfileMap)) {
        try {
            $cred    = Get-CredFileName -Profile $entry.Profile
            $boxCred = Join-Path (Get-BoxMinEdDir -SandboxRoot $SandboxRoot -Box $entry.Box) $cred
            if (-not (Test-CredValid $boxCred)) { continue }
            $backup = Join-Path $CredBackupDir $cred
            if ((Test-Path $backup) -and ((Get-Item $boxCred).Length -lt (Get-Item $backup).Length)) {
                Write-Warning "  $($entry.Box): box token is smaller than the existing backup - not overwriting"
                continue
            }
            Copy-Item $boxCred $backup -Force
            Write-Host "  backed up $($entry.Box) ($($entry.Profile))"
        }
        catch {
            Write-Warning "  $($entry.Box): backup failed ($($_.Exception.Message)) - skipping"
        }
    }
}

# Seed a box that LACKS a valid token, preferring the native token, else a backup. Never
# overwrites an existing valid box token. Also seeds settings.json if the box lacks one.
# One box's error is warned and skipped. Returns the post-seed state objects.
function Restore-CmdrCredentials {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$CmdrNames,
        [Parameter(Mandatory)][string]$SandboxRoot,
        [string]$CredBackupDir,
        [hashtable]$ProfileMap,
        [switch]$WhatIfSeed   # report-only, make no changes
    )
    if (-not (Test-Path $SandboxRoot)) {
        Write-Warning "SandboxRoot '$SandboxRoot' does not exist - no boxes to seed. Check your Sandboxie install path."
        return @()
    }
    $nativeDir  = Get-NativeMinEdDir
    $nativeJson = Join-Path $nativeDir 'settings.json'

    foreach ($entry in (Get-CmdrProfileMap -CmdrNames $CmdrNames -ProfileMap $ProfileMap)) {
        try {
            $cred    = Get-CredFileName -Profile $entry.Profile
            $boxDir  = Get-BoxMinEdDir -SandboxRoot $SandboxRoot -Box $entry.Box
            $boxCred = Join-Path $boxDir $cred
            $boxJson = Join-Path $boxDir 'settings.json'

            # --- token ---
            if (Test-CredValid $boxCred) {
                Write-Host "  $($entry.Box): token present - leaving as-is"
            }
            else {
                $src = $null; $srcLabel = $null
                $nativeCred = Join-Path $nativeDir $cred
                $backupCred = if ($CredBackupDir) { Join-Path $CredBackupDir $cred } else { $null }
                if (Test-CredValid $nativeCred)     { $src = $nativeCred; $srcLabel = 'native' }
                elseif (Test-CredValid $backupCred) { $src = $backupCred; $srcLabel = 'backup' }

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
                    Write-Warning "  $($entry.Box) ($($entry.Profile)): no valid native or backup token - needs an interactive first login"
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
        catch {
            Write-Warning "  $($entry.Box): seeding failed ($($_.Exception.Message)) - box will need an interactive login"
        }
    }
    return (Test-CmdrCredentialState -CmdrNames $CmdrNames -SandboxRoot $SandboxRoot -CredBackupDir $CredBackupDir -ProfileMap $ProfileMap)
}

# Orchestrator the launcher calls: report -> back up existing -> seed the gaps -> report.
function Invoke-WingCredentialSeeding {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$CmdrNames,
        [Parameter(Mandatory)][string]$SandboxRoot,
        [Parameter(Mandatory)][string]$CredBackupDir,
        [hashtable]$ProfileMap,
        [switch]$WhatIfSeed
    )
    Write-Host "Credential state (before):" -ForegroundColor Cyan
    Test-CmdrCredentialState -CmdrNames $CmdrNames -SandboxRoot $SandboxRoot -CredBackupDir $CredBackupDir -ProfileMap $ProfileMap |
        Format-Table -AutoSize | Out-String | Write-Host

    if (-not $WhatIfSeed) {
        Write-Host "Backing up existing box tokens:" -ForegroundColor Cyan
        Backup-CmdrCredentials -CmdrNames $CmdrNames -SandboxRoot $SandboxRoot -CredBackupDir $CredBackupDir -ProfileMap $ProfileMap
    }

    Write-Host "Seeding gaps (native -> box, then backup -> box):" -ForegroundColor Cyan
    $after = Restore-CmdrCredentials -CmdrNames $CmdrNames -SandboxRoot $SandboxRoot -CredBackupDir $CredBackupDir -ProfileMap $ProfileMap -WhatIfSeed:$WhatIfSeed

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
