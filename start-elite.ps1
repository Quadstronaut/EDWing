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

    # Set this to $true to run Elite Odyssey Materials Helper
    [Parameter(Mandatory = $false)]
    [bool]$EliteOdysseyMaterialsHelper = $false,

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
    ('autoconfig' -or 'config' -or 'configure' -or 'configuration') {
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
    if (($EDDiscovery -eq $false) -and
    ($EDEngineer -eq $false) -and
    ($EDHM_UI -eq $false) -and
    ($EDMarketConnector -eq $false) -and
    ($EliteDangerous -eq $false) -and
    ($EliteObservatory -eq $false) -and
    ($EliteOdysseyMaterialsHelper -eq $false) -and
    ($EliteTrack -eq $false) -and
    ($VoiceAttack -eq $false) -and
    ($ConfigMode -eq $false) -and
    ($InstallerMode -eq $false) -and
    ($UninstallerMode -eq $false)) {
        Write-Error 'All parameters are set to $false'
        Exit ('ID10T')
    }
    # Don't operate normally under assumed admin role
    if (([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Output 'Normal execution running as administrator. Lets try launching again without Admin priviliges.'
        $arguments = "-File `"$PSCommandPath`""
        Start-Process powershell.exe -NoProfile -ExecutionPolicy Bypass -ArgumentList $arguments
        Exit 2
    }
}
elseif (($ConfigMode -eq $true) -or ($InstallerMode -eq $true)) {
    # Don't attempt config/installation without assumed admin role
    if (-not(([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))) {
        Write-Output 'Script is not running as administrator. Lets try launching again with Admin priviliges.'
        $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
        Start-Process powershell.exe -NoProfile -ExecutionPolicy Bypass -Verb RunAs -ArgumentList $arguments
        Exit 3
    }
}

### Functions ###

function Auto-Config () {
    Write-Host 'No configuration file detected, auto-configuring...'
    New-Config()
}

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

function InstallerMode() {
    
}

function New-Config () {
    # Define the default configuration
    $Config = @{
        EDDiscovery       = @{
            Path        = 'C:\Program Files\EDDiscovery\EDDiscovery.exe'
            IsInstalled = $false
        }
        EDEngineer        = @{
            Path        = "$HOME\AppData\Local\EDEngineer\EDEngineer.exe"
            IsInstalled = $false
        }
        EDHM_UI           = @{
            Path        = "$HOME\AppData\Local\EDHM_UI\EDHM_UI_mk2.exe"
            IsInstalled = $false
        }
        EDMarketConnector = @{
            Path        = 'C:\Program Files (x86)\EDMarketConnector\EDMarketConnector.exe'
            IsInstalled = $false
        }
        EliteDangerous    = @{
            Path        = 'steam://rungameid/359320'
            IsInstalled = $false
        }
        EliteObservatory  = @{
            Path        = 'C:\Program Files\Elite Observatory\ObservatoryCore.exe'
            IsInstalled = $false
        }
        EliteOdysseyMaterialsHelper  = @{
            Path        = 'C:\Program Files\Elite Observatory\ObservatoryCore.exe'
            IsInstalled = $false
        }
        EliteTrack        = @{
            Path        = "$HOME\AppData\Local\Programs\EliteTrack\EliteTrack.exe"
            IsInstalled = $false
        }
        VoiceAttack       = @{
            Path        = 'C:\Program Files\VoiceAttack\VoiceAttack.exe'
            IsInstalled = $false
        }
    }

    # Save the default configuration to a PSD1 file
    $Config | ConvertTo-Json | Out-File 'C:\Config\MyScript.psd1'
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

function Test-Config() {

}

function Test-SteamInstalled {
    $SteamRegistryKey = 'HKLM:\Software\Valve\Steam'
    $SteamExecutablePath = 'C:\Program Files (x86)\Steam\steam.exe'
    if (Test-Path $SteamRegistryKey -PathType Any -ErrorAction SilentlyContinue -ErrorVariable _) {
        return $true
    } elseif (Test-Path $SteamExecutablePath -PathType Leaf) {
        return $true
    } else {
        return $false
    }
}

### Logic ###

# Check if this is the first run
if (Test-Config()) { Auto-Config() }

# Enter special modes or launch
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
    
    # Local Software
    if ($EDEngineer) { Start-SecondScreen($Path_EDEngineer) }
    if ($EliteDangerous) { Start-Process -FilePath $Path_EliteDangerous }
    if ($EDHM_UI) { Start-SecondScreen($Path_EDHM_UI) }
    if ($EliteObservatory) { Start-SecondScreen($Path_EliteObservatory) }
    if ($EliteOdysseyMaterialsHelper) { Start-SecondScreen($Path_EliteOdysseyMaterialsHelper) }
    if ($VoiceAttack) { Start-SecondScreen($Path_VoiceAttack) }

    # Streaming
    if ($EliteTrack) { Start-SecondScreen($Path_EliteTrack) }
}
# See you space cowboy...