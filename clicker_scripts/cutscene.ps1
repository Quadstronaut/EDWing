# Define the list of coordinates to double-click.
# You can add or remove coordinate pairs as needed.
# Format: [PSCustomObject]@{ X = <X_coord>; Y = <Y_coord> }
$coordinates = @(
    [PSCustomObject]@{ X = -700; Y = -100 },
    [PSCustomObject]@{ X = -700; Y = 555 },
    [PSCustomObject]@{ X = -700; Y = 1111 },
    [PSCustomObject]@{ X = 999; Y = 555 }
)

#- Functions -#
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
foreach ($coord in $coordinates) {
    # Call the function to perform the double-click at the current coordinates.
    Set-dblClickAtXY -X $coord.X -Y $coord.Y

    # Pause between each action to prevent issues with focus or application responsiveness.
    Start-Sleep -Seconds 1
}