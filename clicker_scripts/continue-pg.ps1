# User config - big ugly list of actions
# Each object includes the coordinates and the type of click to perform (Single/Double).
$actions = @(
    [PSCustomObject]@{ X = -960; Y = -141; ClickType = "Double" },
    [PSCustomObject]@{ X = -800; Y = -42; ClickType = "Double" },
    [PSCustomObject]@{ X = -960; Y = 461; ClickType = "Double" },
    [PSCustomObject]@{ X = -800; Y = 555; ClickType = "Double" },
    [PSCustomObject]@{ X = -960; Y = 1061; ClickType = "Double" },
    [PSCustomObject]@{ X = -800; Y = 1111; ClickType = "Double" },
    [PSCustomObject]@{ X = 277; Y = 377; ClickType = "Double" },
    [PSCustomObject]@{ X = 666; Y = 666; ClickType = "Double" }
)

#- Functions -#
function Set-SingleClickAtPosition {
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

        # 2. Perform the click.
        Write-Host "  -> Performing single click."
        [Win32.MouseAPI]::mouse_event($MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        Start-Sleep -Milliseconds 150
        [Win32.MouseAPI]::mouse_event($MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

        Write-Host "  -> Single click completed successfully." -ForegroundColor Green

    }
    catch {
        Write-Host "An error occurred while trying to single-click at ($X, $Y):" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Yellow
    }
}
function Set-DoubleClickAtPosition {
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
        Write-Host "  -> Performing first click."
        [Win32.MouseAPI]::mouse_event($MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        Start-Sleep -Milliseconds 150
        [Win32.MouseAPI]::mouse_event($MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

        # 3. Pause for the double-click interval.
        Start-Sleep -Milliseconds 150

        # 4. Perform the second click.
        Write-Host "  -> Performing second click."
        [Win32.MouseAPI]::mouse_event($MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        Start-Sleep -Milliseconds 150
        [Win32.MouseAPI]::mouse_event($MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

        Write-Host "  -> Double-click completed successfully." -ForegroundColor Green

    }
    catch {
        Write-Host "An error occurred while trying to double-click at ($X, $Y):" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Yellow
    }
}

#- Prep -#
# Load the required .NET assembly which contains the Cursor class.
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

# Import the `mouse_event` function from user32.dll.
# This function is used to send mouse events directly to the operating system.
# There is a high likelihood this can trigger Windows Defender based on tests.
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

# --------------------------------------------------------------------------
# --- MAIN SCRIPT EXECUTION ---
# --------------------------------------------------------------------------
foreach ($action in $actions) {
    if ($action.ClickType -eq "Double") {
        # Call the function to perform the double-click at the current coordinates.
        Set-DoubleClickAtPosition -X $action.X -Y $action.Y
    }
    elseif ($action.ClickType -eq "Single") {
        # Call the function to perform the double-click at the current coordinates.
        Set-SingleClickAtPosition -X $action.X -Y $action.Y
    }
    else {
        Write-Host "Warning: Unknown click type $($action.ClickType) for coordinates $($action.X), $($action.Y). Skipping." -ForegroundColor Yellow
    }
    # Pause between each action to prevent issues with focus or application responsiveness.
    Start-Sleep -Seconds 1
}