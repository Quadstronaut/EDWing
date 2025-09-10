# --- FUNCTION TO MOVE AND RESIZE WINDOWS ---

# Function to move and resize a window to specific X/Y coordinates.
# It now accepts a ProcessName parameter and returns a boolean to indicate success or failure.
function Set-WindowPosition {
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
        } else {
            # Otherwise, use the standard SetWindowPos to move and resize
            [User32]::SetWindowPos($handle, [IntPtr]::Zero, $X, $Y, $Width, $Height, $SWP_NOZORDER) | Out-Null
            Write-Host "Moved and resized window for '$($WindowTitle)' (Process: '$($ProcessName)') to X=$X, Y=$Y, W=$Width, H=$Height"
        }
        
        return $true
    } else {
        # The window wasn't found, which is a normal state while we wait for it to load.
        # We return false to indicate that it has not been moved yet.
        Write-Warning "Process '$($ProcessName)' with title '$($WindowTitle)' not found. Waiting..."
        return $false
    }
}


# --- SCRIPT CONFIGURATION AND LAUNCH ---

# Define the paths for your executables
$sandboxieStart = 'C:\Users\Quadstronaut\scoop\apps\sandboxie-plus-np\current\Start.exe'
$edminlauncher = 'G:\SteamLibrary\steamapps\common\Elite Dangerous\MinEdLauncher.exe'

# Define the Elite Dangerous accounts to move and their target coordinates/dimensions
$boxes = @(
    @{ ProcessName = "EliteDangerous64"; Name = "CMDRBistronaut"; X = -1080; Y = -387; Width = 800; Height = 600; Moved = $false },
    @{ ProcessName = "EliteDangerous64"; Name = "CMDRTristronaut"; X = -1080; Y = 213; Width = 800; Height = 600; Moved = $false },
    @{ ProcessName = "EliteDangerous64"; Name = "CMDRQuadstronaut"; X = -1080; Y = 813; Width = 800; Height = 600; Moved = $false }
)

# Define the EDMC accounts to move and their target coordinates/dimensions
$edmcs = @(
    @{ ProcessName = "EDMarketConnector"; Name = "CMDRUnistronaut"; X = -280; Y = 1213; Width = 300; Height = 600; Moved = $false },
    @{ ProcessName = "EDMarketConnector"; Name = "CMDRBistronaut"; X = -280; Y = -387; Width = 300; Height = 600; Moved = $false },
    @{ ProcessName = "EDMarketConnector"; Name = "CMDRTristronaut"; X = -280; Y = 213; Width = 300; Height = 600; Moved = $false },
    @{ ProcessName = "EDMarketConnector"; Name = "CMDRQuadstronaut"; X = -280; Y = 813; Width = 300; Height = 600; Moved = $false }
)

# Define the Elite Dangerous Exploration Buddy account to move and its target action
$edeb = @(
    @{ ProcessName = "Elite Dangerous Exploration Buddy"; Name = "CMDRUnistronaut"; Maximize = $true; Moved = $false }
)

# Create a single list of all items to move for looping porpoises
$allWindowsToMove = $edeb + $edmcs + $boxes

# Check that the necessary executables exist
$sbsTrue = Test-Path $sandboxieStart
$edmlTrue = Test-Path $edminlauncher

if ($sbsTrue -and $edmlTrue) {
    # Launch all four Elite Dangerous instances simultaneously.
    # Note: CMDRUnistronaut is launched but will not be moved by this script
    # as it's not in the $boxes array. You have it in the $edmcs array, which is correct.
    Start-Process -FilePath $sandboxieStart -ArgumentList "/box:CMDRUnistronaut `"$edminlauncher`" /frontier Account1 /edo /autorun /autoquit /skipInstallPrompt"
    Start-Process -FilePath $sandboxieStart -ArgumentList "/box:CMDRBistronaut `"$edminlauncher`" /frontier Account2 /edo /autorun /autoquit /skipInstallPrompt"
    Start-Process -FilePath $sandboxieStart -ArgumentList "/box:CMDRTristronaut `"$edminlauncher`" /frontier Account3 /edo /autorun /autoquit /skipInstallPrompt"
    Start-Process -FilePath $sandboxieStart -ArgumentList "/box:CMDRQuadstronaut `"$edminlauncher`" /frontier Account4 /edo /autorun /autoquit /skipInstallPrompt"
    Start-Sleep -Seconds 30

    # --- WINDOW POSITIONING ---

    Write-Host "Waiting for windows to load and positioning them. This may take a moment..."
    
    # This single loop will continue to check for all windows and move them
    # until all have been successfully positioned.
    do {
        # Loop through each window configuration
        foreach ($window in $allWindowsToMove) {
            # Only try to move windows that haven't been successfully moved yet
            if ($window.Moved -eq $false) {
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
} else {
    Write-Error "Could not find one or more required executables. Check your paths."
}
