#Requires -Version 5.1

# Load external config if present (overrides defaults below)
$wingConfPath = Join-Path -Path $PSScriptRoot -ChildPath 'wing.conf.ps1'

# Default configuration
$config = @{
    launchEliteDangerous           = $true
    skipIntro                      = $true
    pgEntry                        = $true
    launchEDEB                     = $false
    launchEDMC                     = $false
    pythonPath                     = 'C:\Users\Quadstronaut\scoop\apps\python\current\python.exe'
    WindowPollInterval             = 3333
    ProcessWaitInterval            = 3333
    WindowMoveRetryInterval        = 3333
    MaxRetries                     = 3
    StopCustomServicesAndProcesses = $false
    EliteWindowSettleSeconds       = 7
    # Credential seeding: back up and restore MinEdLauncher .cred files across sandbox wipes
    SeedCredentials                = $true
    SandboxRoot                    = "C:\Sandbox\$env:USERNAME"
    CredBackupDir                  = "$PSScriptRoot\cred_backup"
}

# Commander configuration
$cmdrNames = @(
    "CMDRDuvrazh",
    "CMDRBistronaut",
    "CMDRTristronaut",
    "CMDRQuadstronaut"
)

# Executable paths
$edmc_path = 'G:\EliteApps\EDMarketConnector\EDMarketConnector.exe'
$sandboxieStart = 'C:\Users\Quadstronaut\scoop\apps\sandboxie-plus-np\current\Start.exe'
$minEDLauncher = 'G:\SteamLibrary\steamapps\common\Elite Dangerous\MinEdLauncher.exe'
$edebLauncher = 'G:\EliteApps\EDEB\Elite Dangerous Exploration Buddy.exe'

# Apply external config overrides
if (Test-Path $wingConfPath) {
    Write-Host "Loading config from $wingConfPath"
    . $wingConfPath
}

# Alt commanders (all except the first, who plays on the primary monitor)
$eliteDangerousCmdrs = $cmdrNames | Select-Object -Skip 1

# --- Win32 API Types (loaded once at script scope) ---
if (-not ([System.Management.Automation.PSTypeName]'User32').Type) {
    Add-Type -TypeDefinition @"
        using System;
        using System.Runtime.InteropServices;
        public class User32 {
            [DllImport("user32.dll")]
            public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

            [DllImport("user32.dll")]
            public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
        }
"@
}

# --- Functions ---

function Set-WindowPosition {
    param(
        [Parameter(Mandatory)]
        [int]$X,

        [Parameter(Mandatory)]
        [int]$Y,

        [Parameter(Mandatory)]
        [ValidateRange(0, [int]::MaxValue)]
        [int]$Width,

        [Parameter(Mandatory)]
        [ValidateRange(0, [int]::MaxValue)]
        [int]$Height,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ProcessName,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$WindowTitle,

        [switch]$Maximize
    )

    $SWP_NOZORDER = 4
    $SW_MAXIMIZE = 3

    $process = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowTitle -like "*$WindowTitle*" }

    if (-not $process) {
        Write-Warning "Process '$ProcessName' with title '$WindowTitle' not found. Waiting..."
        return $false
    }

    $handle = $process.MainWindowHandle
    Start-Sleep -Milliseconds 100

    if ($Maximize) {
        $result = [User32]::ShowWindowAsync($handle, $SW_MAXIMIZE)
        if ($result) {
            Write-Host "Maximized window for '$WindowTitle'"
        }
        else {
            Write-Warning "Failed to maximize window for '$WindowTitle'"
            return $false
        }
    }
    else {
        $result = [User32]::SetWindowPos($handle, [IntPtr]::Zero, $X, $Y, $Width, $Height, $SWP_NOZORDER)
        if ($result) {
            Write-Host "Positioned '$WindowTitle' at X=$X Y=$Y ${Width}x${Height}"
        }
        else {
            Write-Warning "Failed to position window for '$WindowTitle'"
            return $false
        }
    }
    return $true
}

# --- Credential Seeding Functions ---

function Get-SandboxCredPath {
    param([string]$BoxName, [string]$ProfileName)
    Join-Path $config.SandboxRoot "$BoxName\user\current\AppData\Local\min-ed-launcher\.frontier-$($ProfileName.ToLower()).cred"
}

function Backup-SandboxCredentials {
    if (-not (Test-Path $config.CredBackupDir)) {
        New-Item -ItemType Directory -Path $config.CredBackupDir -Force | Out-Null
    }
    for ($i = 0; $i -lt $cmdrNames.Count; $i++) {
        $profileName = "Account$($i + 1)"
        $credFile = Get-SandboxCredPath -BoxName $cmdrNames[$i] -ProfileName $profileName
        $backupFile = Join-Path $config.CredBackupDir ".frontier-$($profileName.ToLower()).cred"
        if (Test-Path $credFile) {
            Copy-Item -Path $credFile -Destination $backupFile -Force
            Write-Host "Backed up credentials for $($cmdrNames[$i]) ($profileName)"
        }
    }
}

function Restore-SandboxCredentials {
    for ($i = 0; $i -lt $cmdrNames.Count; $i++) {
        $profileName = "Account$($i + 1)"
        $backupFile = Join-Path $config.CredBackupDir ".frontier-$($profileName.ToLower()).cred"
        if (-not (Test-Path $backupFile)) {
            Write-Warning "No backed-up credentials for $($cmdrNames[$i]) ($profileName) - will need manual login"
            continue
        }
        $credFile = Get-SandboxCredPath -BoxName $cmdrNames[$i] -ProfileName $profileName
        $credDir = Split-Path $credFile -Parent
        if (-not (Test-Path $credDir)) {
            New-Item -ItemType Directory -Path $credDir -Force | Out-Null
        }
        Copy-Item -Path $backupFile -Destination $credFile -Force
        Write-Host "Restored credentials for $($cmdrNames[$i]) ($profileName)"
    }
}

# --- Build Window Configuration List ---
$windowConfigurations = @()

if ($config.launchEliteDangerous) {
    $eliteWindows = @(
        @{ Name = $eliteDangerousCmdrs[0]; X = -1080; Y = -387; Width = 800; Height = 600; Moved = $false; RetryCount = 0 },
        @{ Name = $eliteDangerousCmdrs[1]; X = -1080; Y = 213;  Width = 800; Height = 600; Moved = $false; RetryCount = 0 },
        @{ Name = $eliteDangerousCmdrs[2]; X = -1080; Y = 813;  Width = 800; Height = 600; Moved = $false; RetryCount = 0 }
    )
    $eliteWindows | ForEach-Object {
        $_.ProcessName = "EliteDangerous64"
        $_.Maximize = $false
    }
    $windowConfigurations += $eliteWindows
}

if ($config.launchEDEB) {
    if (-not (Test-Path $edebLauncher)) {
        Write-Warning "EDEB launcher not found at $edebLauncher - skipping"
    }
    else {
        & $edebLauncher
        $edebWindows = @(
            @{ ProcessName = "Elite Dangerous Exploration Buddy"; Name = $cmdrNames[0]; X = 0; Y = 0; Width = 800; Height = 600; Maximize = $true; Moved = $false; RetryCount = 0 }
        )
        $windowConfigurations += $edebWindows
    }
}

if ($config.launchEDMC) {
    $edmcWindows = @(
        @{ Name = $cmdrNames[0]; X = 100;  Y = 100;  Width = 300; Height = 600; Moved = $false; RetryCount = 0 },
        @{ Name = $cmdrNames[1]; X = -280; Y = -387; Width = 300; Height = 600; Moved = $false; RetryCount = 0 },
        @{ Name = $cmdrNames[2]; X = -280; Y = 213;  Width = 300; Height = 600; Moved = $false; RetryCount = 0 },
        @{ Name = $cmdrNames[3]; X = -280; Y = 813;  Width = 300; Height = 600; Moved = $false; RetryCount = 0 }
    )
    $edmcWindows | ForEach-Object {
        $_.ProcessName = "EDMarketConnector"
        $_.Maximize = $false
    }
    $windowConfigurations += $edmcWindows
}

# --- Validate Required Executables ---
$missingPaths = @()
if (-not (Test-Path $sandboxieStart)) { $missingPaths += "Sandboxie: $sandboxieStart" }
if ($config.launchEliteDangerous -and -not (Test-Path $minEDLauncher)) { $missingPaths += "MinEdLauncher: $minEDLauncher" }
if ($config.launchEDMC -and -not (Test-Path $edmc_path)) { $missingPaths += "EDMC: $edmc_path" }

if ($missingPaths.Count -gt 0) {
    Write-Error "Missing required executables:"
    $missingPaths | ForEach-Object { Write-Host "  - $_" }
    return
}

# --- Stop Custom Processes ---
if ($config.StopCustomServicesAndProcesses) {
    $processesToStop = @("*battle.net*", "*epic*", "*gog*", "*steam*", "*discord*", "*ollama*")
    Get-Process -Name $processesToStop -ErrorAction SilentlyContinue | Stop-Process -Force
    Get-ScheduledTask 'Syncthing' -ErrorAction SilentlyContinue | Stop-ScheduledTask
    Get-Service -Name "Everything", "Filezilla Server" -ErrorAction SilentlyContinue | Stop-Service
}

# --- Credential Seeding Phase ---
if ($config.SeedCredentials) {
    # Back up any existing .cred files before they could be lost
    Backup-SandboxCredentials
    # Restore into sandbox filesystem (creates dirs if sandbox was wiped)
    Restore-SandboxCredentials
}

# --- Launch Phase ---
Write-Host "Starting Elite Dangerous multibox"

for ($i = 0; $i -lt $cmdrNames.Count; $i++) {
    $boxName = $cmdrNames[$i]
    $arguments = "/box:$boxName `"$minEDLauncher`" /frontier Account$($i+1) /edo /autorun /autoquit /skipInstallPrompt"
    Start-Process -FilePath $sandboxieStart -ArgumentList $arguments
    Write-Host "Launched $boxName in sandbox"

    if ($config.launchEDMC) {
        $arguments = "/box:$boxName `"$edmc_path`""
        Start-Process -FilePath $sandboxieStart -ArgumentList $arguments
        Write-Host "Launched EDMC in sandbox $boxName"
    }
}

# --- Window Detection Phase ---
if ($windowConfigurations.Count -eq 0) {
    Write-Host "No windows to manage - skipping detection and positioning"
}
else {
    Write-Host "`nWaiting for application windows to load..."

    $previousCount = -1
    do {
        $windowsFoundCount = 0
        foreach ($window in $windowConfigurations) {
            $process = Get-Process -Name $window.ProcessName -ErrorAction SilentlyContinue |
                Where-Object { $_.MainWindowTitle -like "*$($window.Name)*" }
            if ($process) { $windowsFoundCount++ }
        }

        if ($windowsFoundCount -ne $previousCount) {
            Write-Host "Found $windowsFoundCount of $($windowConfigurations.Count)"
            $previousCount = $windowsFoundCount
        }

        if ($windowsFoundCount -lt $windowConfigurations.Count) {
            Start-Sleep -Milliseconds $config.WindowPollInterval
        }
    } until ($windowsFoundCount -eq $windowConfigurations.Count)

    # --- Window Positioning Phase ---
    Write-Host "`nAll windows detected. Beginning positioning..."

    $eliteSettled = $false
    do {
        foreach ($window in $windowConfigurations) {
            if ($window.Moved) { continue }

            # Wait for process to appear
            do {
                $process = Get-Process -Name $window.ProcessName -ErrorAction SilentlyContinue |
                    Where-Object { $_.MainWindowTitle -like "*$($window.Name)*" }
                if (-not $process) {
                    Start-Sleep -Milliseconds $config.ProcessWaitInterval
                }
            } while (-not $process)

            # Elite instances need extra time to finish loading their renderer
            if ($window.ProcessName -eq "EliteDangerous64" -and -not $eliteSettled) {
                Write-Host "Waiting $($config.EliteWindowSettleSeconds)s for Elite windows to settle..."
                Start-Sleep -Seconds $config.EliteWindowSettleSeconds
                $eliteSettled = $true
            }

            # Attempt positioning with retry logic
            if ($window.RetryCount -lt $config.MaxRetries) {
                $positioned = Set-WindowPosition `
                    -ProcessName $window.ProcessName `
                    -WindowTitle $window.Name `
                    -X $window.X -Y $window.Y `
                    -Width $window.Width -Height $window.Height `
                    -Maximize:([bool]$window.Maximize)

                if ($positioned) {
                    $window.Moved = $true
                }
                else {
                    $window.RetryCount++
                    Write-Host "Retry $($window.RetryCount)/$($config.MaxRetries) for $($window.Name)"
                }
            }
            else {
                Write-Warning "Failed to position $($window.Name) after $($config.MaxRetries) attempts. Skipping."
                $window.Moved = $true
            }
        }

        Start-Sleep -Milliseconds $config.WindowMoveRetryInterval
    } until (($windowConfigurations | Where-Object { -not $_.Moved }).Count -eq 0)

    Write-Host "`nWindow positioning complete!"
}

# --- Post-Launch Automation ---
if ($config.skipIntro) {
    $skipIntro_scriptPath = Join-Path -Path $PSScriptRoot -ChildPath 'clicker_scripts\cutscene.ps1'
    & $skipIntro_scriptPath
    $introSkipped = $true
}
else {
    $introSkipped = $false
}

if ($config.pgEntry -and $introSkipped) {
    Start-Sleep -Seconds 20
    $pgEntry_scriptPath = Join-Path -Path $PSScriptRoot -ChildPath 'clicker_scripts\continue-pg.ps1'
    & $pgEntry_scriptPath
}
