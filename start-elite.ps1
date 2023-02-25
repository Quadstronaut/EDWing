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
switch -Regex ($PSBoundParameters.Keys) {
    { $_ -in 'autoconfig', 'config', 'configure', 'configuration' } {
        $ConfigMode = $true
    }
    { $_ -in 'help', 'github' } {
        Start-Process 'https://github.com/GitKageHub/EDQuick' -ErrorAction SilentlyContinue
        Exit 0
    }
    { $_ -in 'install', 'installer', 'add' } {
        $InstallerMode = $true
    }
    { $_ -in 'uninstall', 'uninstaller', 'remove' } {
        $UnInstallerMode = $true
    }
    default {
        Write-Error "Unrecognized parameter: $_"
        Read-Host 'Press Enter to exit. Error status 1.'
        Exit 1
    }
}

### Functions ###

function AutoConfig () {
    Write-Host 'No configuration file detected, generating default config...'
    DefaultConfig
}

function DefaultConfig () {
    $Config = @{
        EDDiscovery                 = @{
            Path        = "$env:ProgramFiles\EDDiscovery\EDDiscovery.exe"
            IsInstalled = $false
        }
        EDEngineer                  = @{
            Path        = "$env:LocalAppData\EDEngineer\EDEngineer.exe"
            IsInstalled = $false
        }
        EDHM_UI                     = @{
            Path        = "$env:LocalAppData\Local\EDHM_UI\EDHM_UI_mk2.exe"
            IsInstalled = $false
        }
        EDMarketConnector           = @{
            Path        = "$env:ProgramFiles(x86)\EDMarketConnector\EDMarketConnector.exe"
            IsInstalled = $false
        }
        EliteDangerous              = @{
            Path        = 'steam://rungameid/359320'
            IsInstalled = $false
        }
        EliteObservatory            = @{
            Path        = "$env:ProgramFiles\Elite Observatory\ObservatoryCore.exe"
            IsInstalled = $false
        }
        EliteOdysseyMaterialsHelper = @{
            Path        = "$env:LocalAppData\Elite Dangerous Odyssey Materials Helper Launcher\Elite Dangerous Odyssey Materials Helper Launcher.exe"
            IsInstalled = $false
        }
        EliteTrack                  = @{
            Path        = "$env:LocalAppData\Programs\EliteTrack\EliteTrack.exe"
            IsInstalled = $false
        }
        VoiceAttack                 = @{
            Path        = "$env:ProgramFiles\VoiceAttack\VoiceAttack.exe"
            IsInstalled = $false
        }
    }

    # Save the default configuration to a PSD1 file
    if (-not (TestExistConfigDirectory)) {
        New-Item -ItemType Directory -Path $configDirectory -Force | Out-Null
    }
    $Config | ConvertTo-Json | Out-File -Encoding utf8 -FilePath "$env:LocalAppData\EDQuick\EDQuick.psd1"
}

function IsAdmin() {
    return (([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
}

function ReadDataFile () {
    $global:EDQConfig = Import-PowerShellDataFile -Path "$env:LocalAppData\EDQuick\EDQuick.psd1"
}

function SearchSoftware () {
    # This is a cpu intensive parallel search function to find all the software
}

function StartSecondScreen ($appPath) {
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

function TestExistConfig {
    $psd1Path = "$env:LocalAppData\EDQuick\EDQuick.psd1"
    if (Test-Path -Path $psd1Path -PathType Leaf) { return $true }
    else { return $false }
}

function TestExistConfigDirectory {
    $configDirectory = "$env:LocalAppData\EDQuick"
    if (Test-Path -Path $configDirectory -PathType Container) { return $true }
    else { return $false }
}

function TestProgramInstalled ($appPath) {

}

function TestSteamInstalled {
    $SteamRegistryKey = 'HKLM:\Software\Valve\Steam'
    $SteamExecutablePath = 'C:\Program Files (x86)\Steam\steam.exe'
    if (Test-Path $SteamRegistryKey -PathType Any -ErrorAction SilentlyContinue -ErrorVariable _) {
        return $true
    }
    elseif (Test-Path $SteamExecutablePath -PathType Leaf) {
        return $true
    }
    else {
        return $false
    }
}

### Logic ###

# Check if this is the first run
if (TestExistConfig) { DefaultConfig } else { ReadDataFile }

# Special Modes
if ($ConfigMode -eq $true) {
    #TODO: Configure everything with PowerShellDataFile
    # This will be an interactive mode
}
if ($InstallerMode -eq $true) {
    #TODO: Install flagged
    # This will be an uninteractive mode
}

# Ignition System
if (($ConfigMode -eq $false) -and ($InstallerMode -eq $false)) {
    # Community Data
    if ($EDMarketConnector) { StartSecondScreen($Global:EDQConfig.EDMarketConnector.Path) }
    if ($EDDiscovery) { StartSecondScreen($Path_EDDiscovery) }
    
    # Local Software
    if ($EDEngineer) { StartSecondScreen($Path_EDEngineer) }
    if ($EliteDangerous) { Start-Process -FilePath $Path_EliteDangerous }
    if ($EDHM_UI) { StartSecondScreen($Path_EDHM_UI) }
    if ($EliteObservatory) { StartSecondScreen($Path_EliteObservatory) }
    if ($EliteOdysseyMaterialsHelper) { StartSecondScreen($Path_EliteOdysseyMaterialsHelper) }
    if ($VoiceAttack) { StartSecondScreen($Path_VoiceAttack) }

    # Streaming
    if ($EliteTrack) { StartSecondScreen($Path_EliteTrack) }
}
# See you space cowboy...