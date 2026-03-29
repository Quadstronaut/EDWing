# MouseUtil.ps1 - Shared mouse automation helpers for clicker scripts
# Dot-source this file: . "$PSScriptRoot\MouseUtil.ps1"

Add-Type -AssemblyName System.Windows.Forms

if (-not ([System.Management.Automation.PSTypeName]'Win32.MouseAPI').Type) {
    Add-Type -MemberDefinition @'
[DllImport("user32.dll")]
public static extern void mouse_event(
    int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo
);
'@ -Namespace Win32 -Name MouseAPI
}

$script:MOUSEEVENTF_LEFTDOWN = 0x0002
$script:MOUSEEVENTF_LEFTUP   = 0x0004

function Invoke-SingleClick {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][int]$X,
        [Parameter(Mandatory)][int]$Y,
        [int]$HoldMs = 50
    )
    [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($X, $Y)
    [Win32.MouseAPI]::mouse_event($script:MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    Start-Sleep -Milliseconds $HoldMs
    [Win32.MouseAPI]::mouse_event($script:MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
}

function Invoke-DoubleClick {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][int]$X,
        [Parameter(Mandatory)][int]$Y,
        [int]$HoldMs = 50,
        [int]$GapMs  = 100
    )
    [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($X, $Y)
    [Win32.MouseAPI]::mouse_event($script:MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    Start-Sleep -Milliseconds $HoldMs
    [Win32.MouseAPI]::mouse_event($script:MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    Start-Sleep -Milliseconds $GapMs
    [Win32.MouseAPI]::mouse_event($script:MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    Start-Sleep -Milliseconds $HoldMs
    [Win32.MouseAPI]::mouse_event($script:MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
}

function Invoke-ClickAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][int]$X,
        [Parameter(Mandatory)][int]$Y,
        [ValidateSet("Single","Double")]
        [string]$ClickType = "Double",
        [int]$DelayAfterMs = 1000
    )
    Write-Verbose "Click ($ClickType) at X=$X Y=$Y"
    if ($ClickType -eq "Single") {
        Invoke-SingleClick -X $X -Y $Y
    }
    else {
        Invoke-DoubleClick -X $X -Y $Y
    }
    Start-Sleep -Milliseconds $DelayAfterMs
}
