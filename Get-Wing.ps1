function Set-WindowPosition {
    # It accepts a ProcessName parameter and returns a boolean to indicate success or failure.
    # Function to move and resize a window to specific X/Y coordinates.
    param(
        [int]$X,
        [int]$Y,
        [int]$Width,
        [int]$Height,
        [string]$ProcessName,
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

    # Find the process by name and window title. The process name is now dynamic.
    $process = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -like "*$WindowTitle*" }
    
    if ($process) {
        $handle = $process.MainWindowHandle

        # If the Maximize switch is used, use the ShowWindowAsync function
        if ($Maximize) {
            [User32]::ShowWindowAsync($handle, $SW_MAXIMIZE) | Out-Null
            Write-Host "Maximized window for '$($WindowTitle)' (Process: '$($ProcessName)')"
        }
        else {
            # Otherwise, use the standard SetWindowPos to move and resize
            [User32]::SetWindowPos($handle, [IntPtr]::Zero, $X, $Y, $Width, $Height, $SWP_NOZORDER) | Out-Null
            Write-Host "Moved and resized window for '$($WindowTitle)' (Process: '$($ProcessName)') to X=$X, Y=$Y, W=$Width, H=$Height"
        }
        
        return $true
    }
    else {
        # The window wasn't found, which is a normal state while we wait for it to load.
        # We return false to indicate that it has not been moved yet.
        Write-Warning "Process '$($ProcessName)' with title '$($WindowTitle)' not found. Waiting..."
        return $false
    }
} ### END Set-WindowPosition

# Dot-source the wing.conf.ps1 file to load its variables
# This variable can be a blank string if a selection of the script is run in the VS Codium terminal.
$script_root = if ($PSScriptRoot) {
    # If $PSScriptRoot has a value, use it.
    $PSScriptRoot
}
else {
    # Otherwise, fall back to the current working directory ($PWD).
    $PWD.Path
}
$conf_file = Join-Path -Path $script_root -ChildPath 'wing.conf'

# Use a single, simple condition to check if the file exists.
if (Test-Path -Path $conf_file -PathType Leaf) {

    # The try/catch block is for robust error handling.
    # It attempts to dot-source the file.
    try {
        # Dot-sourcing executes the script and loads its variables and functions
        # into the current scope.
        Write-Information "Loading configuration from '$conf_file'..."
        . $conf_file
        
        # A good practice is to provide positive feedback after a successful operation.
        Write-Information "Configuration loaded successfully."
    }
    catch {
        # If an error occurs during dot-sourcing (e.g., a syntax error in the conf file),
        # this block will catch it and provide a clear error message.
        Write-Error "An error occurred while loading the configuration file. Please check the file for syntax errors."
        Write-Error $_
    }
}
else {
    Write-Information "Configuration file '$conf_file' not found."
}

# Update the process names for the dynamic arrays
$client | ForEach-Object { $_.ProcessName = "EliteDangerous64" }
$edmc | ForEach-Object { $_.ProcessName = "EDMarketConnector" }

# Create a single list of all items to move for looping porpoises
$allWindowsToMove = $edeb + $edmc + $client

# Check that the necessary executables exist
$sbsTrue = Test-Path $sandboxieStart
$edmlTrue = Test-Path $edminLauncher

# Set a single variable to true if all paths exist
$all_apps_are_go = $sbsTrue -and $edmlTrue

if ($all_apps_are_go) {
    # Launch all four Elite Dangerous instances simultaneously.
    # The account names are now dynamically pulled from the $cmdrNames array.
    for ($i = 0; $i -lt $cmdrNames.Count; $i++) {
        Start-Process -FilePath $sandboxieStart -ArgumentList "/box:$($cmdrNames[$i]) `"$edminLauncher`" /frontier Account$($i+1) /edo /autorun /autoquit /skipInstallPrompt"
    }

    # --- WINDOW FINDING & WAITING ---

    $previousCount = -1
    do {
        $windowsFoundCount = 0
        foreach ($window in $allWindowsToMove) {
            $process = Get-Process -Name $window.ProcessName -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -like "*$($window.Name)*" }
            if ($process) {
                $windowsFoundCount++
            }
        }
    
        # Check if the current count is different from the previous count.
        if ($windowsFoundCount -ne $previousCount) {
            Clear-Host
            Write-Host "Polling for windows to load..."
            Write-Host "Found $windowsFoundCount of $($allWindowsToMove.Count)"
        
            # Update the previous count to the current count for the next loop.
            $previousCount = $windowsFoundCount
        }

        # Only sleep if we haven't found all the windows yet.
        if ($windowsFoundCount -lt $allWindowsToMove.Count) {
            Start-Sleep -Milliseconds 333
        }
    
    } until ($windowsFoundCount -eq $allWindowsToMove.Count)
    
    # --- WINDOW POSITIONING ---

    Write-Host "Positioning windows..."
    # This single loop will continue to check for all windows and move them
    # until all have been successfully positioned.
    $first_wait = $true
    do {
        # Loop through each window configuration
        foreach ($window in $allWindowsToMove) {
            # Only try to move windows that haven't been successfully moved yet
            if ($window.Moved -eq $false) {
                # We'll now wait until the window title contains the commander's name,
                # which indicates the game has loaded and is ready to be positioned.
                do {
                    $process = Get-Process -Name $window.ProcessName -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -like "*$($window.Name)*" }
                    if (-not $process) {
                        Write-Host "Waiting for process $($window.ProcessName) to have title with $($window.Name)..."
                        Start-Sleep -Milliseconds 500
                    }
                } while (-not $process)
                
                # Check if the current window is an Elite Dangerous client before waiting
                if ($window.ProcessName -eq "EliteDangerous64") {
                    Write-Host "Found new state for $($window.Name)."
                    if ($true -eq $first_wait) {
                        # Unfortunately the Get-Random -Maximum parameter is non-inclusive. So, let the games begin.
                        $seven = 7
                        $eleven = 12
                        Get-Random -Minimum $seven -Maximum $eleven -Verbose
                        Start-Sleep -Seconds $seven
                        $first_wait = $false
                    }
                }

                # If the function call is successful, update the 'Moved' property
                if (Set-WindowPosition -ProcessName $window.ProcessName -WindowTitle $window.Name -X $window.X -Y $window.Y -Width $window.Width -Height $window.Height -Maximize:$window.Maximize) {
                    $window.Moved = $true
                }
            }
        }
        # Wait a moment before checking again to prevent high CPU usage
        Start-Sleep -Milliseconds 500
        # The loop will exit when the count of windows that haven't been moved is 0
    } until ( ($allWindowsToMove | Where-Object { $_.Moved -eq $false }).Count -eq 0 )
}
else {
    Write-Error "Could not find one or more required executables. Check your paths."
}