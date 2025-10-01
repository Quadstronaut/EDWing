# Configuration constants
$config = @{
    launchEliteDangerous    = $true # Elite - Dangerous (CLIENT)
    skipIntro               = $true # Launch_app skip intro
    pgEntry                 = $true # Launch a script to select Continue and PG
    launchEDEB              = $false # Elite Dangerous Exploration Buddy
    launchEDMC              = $true # Elite Dangerous Market Connector
    pythonPath              = 'C:\Users\Quadstronaut\scoop\apps\python\current\python.exe'
    WindowPollInterval      = 333 # milliseconds between window detection checks
    ProcessWaitInterval     = 500 # milliseconds between process checks
    WindowMoveRetryInterval = 500 # milliseconds between window positioning attempts
    MaxRetries              = 3 # maximum attempts to position each window
}

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
    
    # C# code to call the Windows API functions
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
    
    # Flag values for the SetWindowPos function
    $SWP_NOZORDER = 4
    
    # Constants for ShowWindowAsync
    $SW_MAXIMIZE = 3

    # Find the process by name and window title
    $process = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -like "*$WindowTitle*" }
    
    if ($process) {
        $handle = $process.MainWindowHandle

        # Add small delay to ensure window is fully ready for manipulation
        Start-Sleep -Milliseconds 100

        # If the Maximize switch is used, use the ShowWindowAsync function
        if ($Maximize) {
            $result = [User32]::ShowWindowAsync($handle, $SW_MAXIMIZE)
            if ($result) {
                Write-Host "Maximized window for '$WindowTitle' (Process: '$ProcessName')"
            }
            else {
                Write-Warning "Failed to maximize window for '$WindowTitle' (Process: '$ProcessName')"
                return $false
            }
        }
        else {
            # Otherwise, use the standard SetWindowPos to move and resize
            $result = [User32]::SetWindowPos($handle, [IntPtr]::Zero, $X, $Y, $Width, $Height, $SWP_NOZORDER)
            if ($result) {
                Write-Host "Moved and resized window for '$WindowTitle' (Process: '$ProcessName') to X=$X, Y=$Y, W=$Width, H=$Height"
            }
            else {
                Write-Warning "Failed to position window for '$WindowTitle' (Process: '$ProcessName')"
                return $false
            }
        }
        
        return $true
    }
    else {
        # The window wasn't found, which is a normal state while we wait for it to load
        Write-Warning "Process '$ProcessName' with title '$WindowTitle' not found. Waiting..."
        return $false
    }
} ### END Set-WindowPosition

# Commander configuration
$cmdrNames = @(
    "CMDRDuvrazh",
    "CMDRBistronaut",
    "CMDRTristronaut",
    "CMDRQuadstronaut"
)

# Alt Elite Dangerous commander names (excluding the first CMDR from the list)
$eliteDangerousCmdrs = $cmdrNames | Select-Object -Skip 1

# Executable paths - centralized for easier maintenance
$sandboxieStart = 'C:\Users\Quadstronaut\scoop\apps\sandboxie-plus-np\current\Start.exe'

# Create a single list of all windows to manage
if ($config.launchEliteDangerous) {
    # Elite: Dangerous Odyssey configuration
    $minEDLauncher = 'G:\SteamLibrary\steamapps\common\Elite Dangerous\MinEdLauncher.exe'

    # Define the Elite Dangerous accounts to move and their target coordinates/dimensions
    $eliteWindows = @(
        @{ Name = $eliteDangerousCmdrs[0]; X = -1080; Y = -387; Width = 800; Height = 600; Moved = $false; RetryCount = 0 },
        @{ Name = $eliteDangerousCmdrs[1]; X = -1080; Y = 213; Width = 800; Height = 600; Moved = $false; RetryCount = 0 },
        @{ Name = $eliteDangerousCmdrs[2]; X = -1080; Y = 813; Width = 800; Height = 600; Moved = $false; RetryCount = 0 }
    )
    # Assign process names to window configurations
    $eliteWindows | ForEach-Object { $_.ProcessName = "EliteDangerous64" }
    $windowConfigurations += $eliteWindows
}
if ($config.launchEDEB) {
    # Elite Dangerous Exploration Buddy configuration
    $edebLauncher = "G:\EliteApps\EDEB\Elite Dangerous Exploration Buddy.exe"
    & $edebLauncher

    # Window configuration
    $edebWindows = @(
        @{ ProcessName = "Elite Dangerous Exploration Buddy"; Name = $cmdrNames[0]; Maximize = $true; Moved = $false; RetryCount = 0 }
    )

    $windowConfigurations += $edebWindows
}
if ($config.launchEDMC) {
    # Define the EDMC accounts to move and their target coordinates/dimensions
    $edmcWindows = @(
        @{ Name = $cmdrNames[0]; X = 100; Y = 100; Width = 300; Height = 600; Moved = $false; RetryCount = 0 },
        @{ Name = $cmdrNames[1]; X = -280; Y = -387; Width = 300; Height = 600; Moved = $false; RetryCount = 0 },
        @{ Name = $cmdrNames[2]; X = -280; Y = 213; Width = 300; Height = 600; Moved = $false; RetryCount = 0 },
        @{ Name = $cmdrNames[3]; X = -280; Y = 813; Width = 300; Height = 600; Moved = $false; RetryCount = 0 }
    )
    # Assign process names to window configurations
    $edmcWindows | ForEach-Object { $_.ProcessName = "EDMarketConnector" }
    $windowConfigurations += $edmcWindows 
}

# Validate that all required executables exist
$sbsTrue = Test-Path $sandboxieStart
$edmlTrue = Test-Path $minEDLauncher

$all_apps_are_go = $sbsTrue -and $edmlTrue

if ($all_apps_are_go) {
    Write-Host "Starting Elite Dangerous multibox"

    # Launch all four Elite Dangerous instances simultaneously
    for ($i = 0; $i -lt $cmdrNames.Count; $i++) {
        $arguments = "/box:$($cmdrNames[$i]) `"$minEDLauncher`" /frontier Account$($i+1) /edo /autorun /autoquit /skipInstallPrompt"
        Start-Process -FilePath $sandboxieStart -ArgumentList $arguments
        Write-Host "Launched $($cmdrNames[$i]) in sandbox"
    }

    # --- WINDOW DETECTION PHASE ---
    Write-Host "`nWaiting for application windows to load..."
    
    $previousCount = -1
    do {
        $windowsFoundCount = 0
        foreach ($window in $windowConfigurations) {
            $process = Get-Process -Name $window.ProcessName -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -like "*$($window.Name)*" }
            if ($process) {
                $windowsFoundCount++
            }
        }
    
        # Update display only when count changes to reduce console spam
        if ($windowsFoundCount -ne $previousCount) {
            Write-Host "Found $windowsFoundCount of $($windowConfigurations.Count)"
            $previousCount = $windowsFoundCount
        }

        # Only sleep if we haven't found all the windows yet
        if ($windowsFoundCount -lt $windowConfigurations.Count) {
            Start-Sleep -Milliseconds $config.WindowPollInterval
        }
    
    } until ($windowsFoundCount -eq $windowConfigurations.Count)
    
    # --- WINDOW POSITIONING PHASE ---
    Write-Host "`nAll windows detected. Beginning positioning..."
    
    $first_wait = $true
    do {
        foreach ($window in $windowConfigurations) {
            # Only try to move windows that haven't been successfully positioned yet
            if ($window.Moved -eq $false) {
                # Wait for process to be ready with correct window title
                do {
                    $process = Get-Process -Name $window.ProcessName -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -like "*$($window.Name)*" }
                    if (-not $process) {
                        Start-Sleep -Milliseconds $config.ProcessWaitInterval
                    }
                } while (-not $process)
                
                # Special handling for Elite Dangerous instances - they need extra time to fully load
                if ($window.ProcessName -eq "EliteDangerous64") {
                    Write-Host "Found Elite Dangerous instance for $($window.Name)."
                    if ($first_wait) {
                        # Unfortunately the Get-Random -Maximum parameter is non-inclusive. So, let the games begin.
                        $seven = 7
                        $eleven = 12
                        Get-Random -Minimum $seven -Maximum $eleven -Verbose
                        Start-Sleep -Seconds $seven
                        $first_wait = $false
                    }
                }

                # Attempt to position the window with retry logic
                $positioned = $false
                if ($window.RetryCount -lt $config.MaxRetries) {
                    $positioned = Set-WindowPosition -ProcessName $window.ProcessName -WindowTitle $window.Name -X $window.X -Y $window.Y -Width $window.Width -Height $window.Height -Maximize:$window.Maximize
                    
                    if ($positioned) {
                        $window.Moved = $true
                    }
                    else {
                        $window.RetryCount++
                        Write-Host "Retry attempt $($window.RetryCount)/$($config.MaxRetries) for $($window.Name)"
                    }
                }
                else {
                    # Max retries reached - mark as moved to prevent infinite loop
                    Write-Warning "Failed to position window $($window.Name) after $($config.MaxRetries) attempts. Skipping."
                    $window.Moved = $true
                }
            }
        }
        
        # Brief pause before checking again to prevent high CPU usage
        Start-Sleep -Milliseconds $config.WindowMoveRetryInterval
        
    } until (($windowConfigurations | Where-Object { $_.Moved -eq $false }).Count -eq 0)

    Write-Host "`nWindow positioning complete!"
}
else {
    Write-Error 'Could not find one or more required executables. Check your paths:'
    Write-Host "Sandboxie Start: $sandboxieStart (Exists: $sbsTrue)"
    Write-Host "Elite Dangerous Launcher: $minEDLauncher (Exists: $edmlTrue)"
}

if ($config.skipIntro) {
    $scriptA_path = Join-Path -Path $PSScriptRoot -ChildPath 'clicker_scripts\cutscene.ps1'
    & $scriptA_path
    $introSkipped = $true
}
else {
    $introSkipped = $false
}

if ($config.pgEntry -and $introSkipped) {
    Start-Sleep -seconds 20 # 20 seconds was based on repeated runs during testing - it's tuned to the dev rig
    $scriptB_path = Join-Path -Path $PSScriptRoot -ChildPath 'C:\Users\Quadstronaut\Documents\Git\EDWing\clicker_scripts\continue-pg.ps1'
    & $scriptB_path
}

# Start-Process -WorkingDirectory $PSScriptRoot -FilePath $config.pythonPath -ArgumentList 'input_broadcast.py'