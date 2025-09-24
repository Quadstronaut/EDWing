# Configuration constants
$config = @{
    WindowPollInterval      = 333      # milliseconds between window detection checks
    ProcessWaitInterval     = 500     # milliseconds between process checks
    WindowMoveRetryInterval = 500 # milliseconds between window positioning attempts
    MaxRetries              = 3                # maximum attempts to position each window
    launchAdditional        = $false     # launch additional non-core apps (EDEB,SRVS,etc)
    skipIntro               = $true             # Launch a py app to skip the opening cutscene
    pythonPath              = 'C:\Users\Quadstronaut\scoop\apps\python\current\python.exe'
    autoloadPY              = 'C:\Users\Quadstronaut\Documents\Git\EDWing\autoload\autoload.py'
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
#$edhm_uiLauncher = 'C:\Users\Quadstronaut\AppData\Local\EDHM-UI-V3\EDHM-UI-V3.exe'
$edminLauncher = 'G:\SteamLibrary\steamapps\common\Elite Dangerous\MinEdLauncher.exe'

# Define the Elite Dangerous accounts to move and their target coordinates/dimensions
$eliteWindows = @(
    @{ Name = $eliteDangerousCmdrs[0]; X = -1080; Y = -387; Width = 800; Height = 600; Moved = $false; RetryCount = 0 },
    @{ Name = $eliteDangerousCmdrs[1]; X = -1080; Y = 213; Width = 800; Height = 600; Moved = $false; RetryCount = 0 },
    @{ Name = $eliteDangerousCmdrs[2]; X = -1080; Y = 813; Width = 800; Height = 600; Moved = $false; RetryCount = 0 }
)

# Define the EDMC accounts to move and their target coordinates/dimensions
$edmcWindows = @(
    @{ Name = $cmdrNames[0]; X = 100; Y = 100; Width = 300; Height = 600; Moved = $false; RetryCount = 0 },
    @{ Name = $cmdrNames[1]; X = -280; Y = -387; Width = 300; Height = 600; Moved = $false; RetryCount = 0 },
    @{ Name = $cmdrNames[2]; X = -280; Y = 213; Width = 300; Height = 600; Moved = $false; RetryCount = 0 },
    @{ Name = $cmdrNames[3]; X = -280; Y = 813; Width = 300; Height = 600; Moved = $false; RetryCount = 0 }
)

<# Elite Dangerous Exploration Buddy configuration
$edebLauncher = "G:\EliteApps\EDEB\Elite Dangerous Exploration Buddy.exe"
$edebWindows = @(
    @{ ProcessName = "Elite Dangerous Exploration Buddy"; Name = $cmdrNames[0]; Maximize = $true; Moved = $false; RetryCount = 0 }
)#>

# Assign process names to window configurations
$eliteWindows | ForEach-Object { $_.ProcessName = "EliteDangerous64" }
$edmcWindows | ForEach-Object { $_.ProcessName = "EDMarketConnector" }

# Create a single list of all windows to manage
$windowConfigurations = $edmcWindows + $eliteWindows

# Validate that all required executables exist
$sbsTrue = Test-Path $sandboxieStart
$edmlTrue = Test-Path $edminLauncher

$all_apps_are_go = $sbsTrue -and $edmlTrue

if ($all_apps_are_go) {
    Write-Host "Starting Elite Dangerous multibox"

    # Launch all four Elite Dangerous instances simultaneously
    for ($i = 0; $i -lt $cmdrNames.Count; $i++) {
        $arguments = "/box:$($cmdrNames[$i]) `"$edminLauncher`" /frontier Account$($i+1) /edo /autorun /autoquit /skipInstallPrompt"
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
    Write-Host "Elite Dangerous Launcher: $edminLauncher (Exists: $edmlTrue)"
}

if ($config.skipIntro) {
    try {
        # Load the required .NET assembly for cursor position and other UI elements.
        # This must be done at the beginning of the script.
        try {
            Add-Type -AssemblyName System.Windows.Forms
        }
        catch {
            Write-Host "Error: Failed to load System.Windows.Forms assembly." -ForegroundColor Red
            Write-Host "Please ensure you are running this on a Windows environment with .NET installed." -ForegroundColor Red
            return
        }

        # Define the mouse event constants for the Windows API.
        $MOUSEEVENTF_LEFTDOWN = 0x0002
        $MOUSEEVENTF_LEFTUP = 0x0004

        # Use PInvoke to import the `mouse_event` function from user32.dll.
        # This function is used to send mouse events directly to the operating system.
        $Signature = @'
[DllImport("user32.dll")]
public static extern void mouse_event(
    int dwFlags,
    int dx,
    int dy,
    int cButtons,
    int dwExtraInfo
);
'@

        # Add the C# type definition to PowerShell.
        Add-Type -MemberDefinition $Signature -Namespace Win32 -Name MouseAPI

        function Set-dblClickAtXY {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory = $true)]
                [int]$X,

                [Parameter(Mandatory = $true)]
                [int]$Y
            )

            try {
                # 1. Move the cursor to the specified coordinates.
                Write-Host "Moving cursor to X: $X, Y: $Y..."
                [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($X, $Y)

                # 2. Perform the first click.
                # Recommended delay between key down and key up is 20-50ms.
                # This ensures the click registers as a valid event.
                Write-Host "  -> Performing first click."
                [Win32.MouseAPI]::mouse_event($MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                Start-Sleep -Milliseconds 50
                [Win32.MouseAPI]::mouse_event($MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

                # 3. Pause for the double-click interval.
                # Standard double-click interval is 50-200ms. 100ms is a safe value.
                Start-Sleep -Milliseconds 100

                # 4. Perform the second click.
                Write-Host "  -> Performing second click."
                [Win32.MouseAPI]::mouse_event($MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                Start-Sleep -Milliseconds 50
                [Win32.MouseAPI]::mouse_event($MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

                Write-Host "  -> Double-click completed successfully." -ForegroundColor Green

            }
            catch {
                Write-Host "An error occurred while trying to double-click at ($X, $Y):" -ForegroundColor Red
                Write-Host $_.Exception.Message -ForegroundColor Yellow
            }
        }

        # --------------------------------------------------------------------------
        # --- MAIN SCRIPT EXECUTION ---
        # --------------------------------------------------------------------------

        # Define the list of coordinates to double-click.
        # You can add or remove coordinate pairs as needed.
        # Format: [PSCustomObject]@{ X = <X_coord>; Y = <Y_coord> }
        $coordinates = @(
            [PSCustomObject]@{ X = -700; Y = -100 },
            [PSCustomObject]@{ X = -700; Y = 555 },
            [PSCustomObject]@{ X = -700; Y = 1111 },
            [PSCustomObject]@{ X = 999; Y = 555 }
        )

        Write-Host "Starting automated double-click script." -ForegroundColor Cyan
        Write-Host "The script will process each coordinate with a delay between them." -ForegroundColor Cyan
        Write-Host "Press Ctrl+C at any time to stop the script." -ForegroundColor Gray
        Write-Host "--------------------------------------------------------"

        # Loop through each coordinate object in the list.
        foreach ($coord in $coordinates) {
            # Call the function to perform the double-click at the current coordinates.
            Set-dblClickAtXY -X $coord.X -Y $coord.Y

            # Pause between each double-click to prevent issues with focus or application responsiveness.
            # A delay of 1-2 seconds is recommended.
            Write-Host "Pausing for 2 seconds before moving to the next coordinate..." -ForegroundColor DarkGray
            Start-Sleep -Seconds 2
        }

        Write-Host "--------------------------------------------------------"
        Write-Host "Script completed." -ForegroundColor Cyan
    }
    catch {

    }
    finally {

    }
}

if ($config.launchAdditional) {
    Write-Host "Not yet, buddy."
}