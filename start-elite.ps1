[CmdletBinding()]
param (
    # Set this to $true to run EDDiscovery
    [Parameter(Mandatory = $false)]
    [bool]$EDDiscovery = $false,

    # Set this to $true to run EDEngineer
    [Parameter(Mandatory = $false)]
    [bool]$EDEngineer = $false,

    # Set this to $true to run EDHM_UI
    [Parameter(Mandatory = $false)]
    [bool]$EDHM_UI = $false,

    # Set this to $true to run EDMarketConnector
    [Parameter(Mandatory = $false)]
    [bool]$EDMarketConnector = $false,

    # Set this to $true to run Elite Dangerous
    [Parameter(Mandatory = $false)]
    [bool]$EliteDangerous = $false,

    # Set this to $true to run Elite Observatory
    [Parameter(Mandatory = $false)]
    [bool]$EliteObservatory = $false,

    # Set this to $true to run Elite Track
    [Parameter(Mandatory = $false)]
    [bool]$EliteTrack = $false,

    # Set this to $true to run VoiceAttack
    [Parameter(Mandatory = $false)]
    [bool]$VoiceAttack = $false,

    # Set this to $true to enter configuration mode
    [Parameter(Mandatory = $false)]
    [bool]$ConfigMode = $false,

    # Set this to $true to enter installer mode
    [Parameter(Mandatory = $false)]
    [bool]$InstallerMode = $false,

    # Set this to $true to enter installer mode
    [Parameter(Mandatory = $false)]
    [bool]$UninstallerMode = $false
)

# Check for named parameters using $PSBoundParameters automatic variable to examine the bound parameters
Switch -Regex ($PSBoundParameters.Keys) {
    ('config' -or 'configure' -or 'configuration') {
        $ConfigMode = $true
    }
    ('help' -or 'github') {
        Start-Process 'https://github.com/GitKageHub/EDQuick' -ErrorAction SilentlyContinue
        Exit 0
    }
    ('install' -or 'installer') {
        $InstallerMode = $true
    }
    ('uninstall' -or 'uninstaller') {
        $UnInstallerMode = $true
    }
    default {
        Write-Error "Unrecognized parameter: $_"
        Read-Host 'Press Enter to exit. Error status 1.'
        Exit 1
    }
}

### Sanity Checks ###

if (($ConfigMode -eq $false) -and ($InstallerMode -eq $false)) {
    # Don't operate normally under assumed admin role
    if (([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Output 'Normal execution running as administrator. Lets try launching again without Admin priviliges.'
        $arguments = "-File `"$PSCommandPath`""
        Start-Process powershell.exe -NoProfile -ExecutionPolicy Bypass -ArgumentList $arguments
        Exit
    }
}
elseif (($ConfigMode -eq $true) -or ($InstallerMode -eq $true)) {
    # Don't attempt config/installation without assumed admin role
    if (-not(([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))) {
        Write-Output 'Script is not running as administrator. Lets try launching again with Admin priviliges.'
        $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
        Start-Process powershell.exe -NoProfile -ExecutionPolicy Bypass -Verb RunAs -ArgumentList $arguments
        Exit
    }
}

### Functions ###

function Check-AllParametersAreFalse {
    # Check if all parameters are false
    param (
        [Parameter(Mandatory = $false)]
        [bool[]]$Params
    )

    # Check if all parameters are false
    $result = ($Params -notcontains $true)

    return $result
}

function Config-Initial () {

}

function InstallerMode() {
    
}

function Read-DataFile () {

}

function Read-UserConfirmation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$Message = 'Confirm with Y or Yes to continue, press Escape or deny with N or No to exit.'
    )
    do {
        $choice = Read-Host $Message
        if ($choice -in 'y', 'Y', 'yes', 'Yes', 'YES') {
            return $true
        }
        elseif ($choice -in 'n', 'N', 'no', 'No', 'NO', [char]27) {
            return $false
        }
        $Message = "Invalid input. Confirm with 'y/Y/yes' to continue, or deny with Escape or 'n/N/no' to exit."
    } while ($true)
}

function Start-SecondScreen ($appPath) {
    $process = Start-Process $appPath -PassThru
    # Get the handle of the main window of the process
    $windowHandle = $process.MainWindowHandle
    # Get the handle of the secondary monitor
    $secondaryMonitor = (Get-WmiObject -Namespace root\wmi -Class WmiMonitorBasicDisplayParams | Where-Object { $_.IsActive -eq $true }).InstanceName
    # Move the window to the secondary monitor
    $winAPI = Add-Type -Name WinAPI -MemberDefinition @'
[DllImport("user32.dll")]
public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
'@
    $winAPI::SetWindowPos($windowHandle, 0, $secondaryMonitor.ScreenWidth, 0, 0, 0, 0x0001)
}

### Ignition ###

if (($ConfigMode -eq $true) -or ($InstallerMode -eq $true)) {
    if ($ConfigMode -eq $true) {
        #TODO: Configure everything with PowerShellDataFile
    }
    if ($InstallerMode -eq $true) {
        #TODO: Install flagged
    }
}
elseif (($ConfigMode -eq $false) -and ($InstallerMode -eq $false)) {

    # Community Data
    if ($EDMarketConnector) { Start-SecondScreen($Path_EDMarketConnector) }
    if ($EDDiscovery) { Start-SecondScreen($Path_EDDiscovery) }
    if ($EDEngineer) { Start-SecondScreen($Path_EDEngineer) }
    if ($EliteObservatory) { Start-SecondScreen($Path_EliteObservatory) }

    # Local Software
    if ($EDHM_UI) { Start-SecondScreen($Path_EDHM_UI) }
    if ($EliteDangerous) { Start-Process -FilePath $Path_EliteDangerous }
    if ($VoiceAttack) { Start-SecondScreen($Path_VoiceAttack) }

    # Streaming
    if ($EliteTrack) { Start-SecondScreen($Path_EliteTrack) }
}
# See you space cowboy...