[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [bool]$EDDiscovery = $false,
    [Parameter(Mandatory = $false)]
    [bool]$EDEngineer = $false,
    [Parameter(Mandatory = $false)]
    [bool]$EDHM_UI = $false,
    [Parameter(Mandatory = $false)]
    [bool]$EDMarketConnector = $false,
    [Parameter(Mandatory = $false)]
    [bool]$EliteDangerous = $false,
    [Parameter(Mandatory = $false)]
    [bool]$EliteObservatory = $false,
    [Parameter(Mandatory = $false)]
    [bool]$EliteTrack = $false,
    [Parameter(Mandatory = $false)]
    [bool]$VoiceAttack = $false,
    [Parameter(Mandatory = $false)]
    [bool]$InstallerMode = $false
)

# Check to see if -help was passed by using the $PSBoundParameters automatic variable to examine the bound parameters.
if ($PSBoundParameters.ContainsKey('help')) {
    Write-Host 'This script is designed to help you efficiently launch and kill all processes related to Elite Dangerous. It is highly dynamic and should work on most systems. For more information, please visit the repository at https://github.com/GitKageHub/EDQuick.'
    # Prompt the user to press Enter to continue
    Read-Host 'Press Enter to exit.'
}

if (([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Output 'Script is running as administrator. Running with elevated permissions causes issues with some of these programs and where they store your user data. Lets try launching again without Admin priviliges.'
    $arguments = "-File `"$PSCommandPath`""
    Start-Process powershell.exe -Verb RunAs -ArgumentList $arguments
    Exit
}

### Paths ###
$Path_EDHM_UI = "$HOME\AppData\Local\EDHM_UI\EDHM_UI_mk2.exe"
$Path_VoiceAttack = 'C:\Program Files\VoiceAttack\VoiceAttack.exe'
$Path_EDMarketConnector = 'C:\Program Files (x86)\EDMarketConnector\EDMarketConnector.exe'
$Path_EDDiscovery = 'C:\Program Files\EDDiscovery\EDDiscovery.exe'
$Path_EDEngineer = Get-ChildItem -Path "$HOME\AppData" -Filter 'EDEngineer.exe' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
$Path_EliteObservatory = 'C:\Program Files\Elite Observatory\ObservatoryCore.exe'
$Path_EliteTrack = "$HOME\AppData\Local\Programs\EliteTrack\EliteTrack.exe"
$Path_EliteDangerous = 'steam://rungameid/359320'

### Functions ###
function MSIinstall ($url) {
    if ($url.EndsWith('/latest')) {
        # Define the URL of the GitHub release page
        $url = 'https://github.com/BlueMystical/EDHM_UI/releases/latest'
        # Use the Invoke-WebRequest cmdlet to download the HTML of the release page
        $response = Invoke-WebRequest -Uri $url
        # Extract the download link for the MSI file from the HTML
        $downloadLink = ($response.Links | Where-Object { $_.href -like '*.msi' }).href
    }
    else { $downloadLink = $url }
    # Use the Path.GetFileName method to extract the filename from the URL string
    $fileName = [System.IO.Path]::GetFileName($downloadLink)
    # Define the path where the MSI file will be downloaded
    $msiFile = "$($env:USERPROFILE)\Downloads\$fileName"
    # Download the MSI file
    Write-Output "Downloading $url..."
    Invoke-WebRequest $url -OutFile $msiFile
    # Install the MSI file silently
    Write-Output "Installing $msiFile..."
    Start-Process msiexec.exe -ArgumentList "/i `"$msiFile`" /qn" -Wait
    Write-Output 'Done.'
}

function Start-SecondScreen ($appPath) {
    # Start the test.exe process
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

function Test-Install ($softwareArray) {
    $Path_EDHM_UI
}

function Test-Path-Prompt ($path) {
    Write-Output "$path is not installed. Install? (Y/N)"
}

### Test-Paths ###
if (-not(Test-Path -Path $Path_EDHM_UI -PathType Any -ErrorAction SilentlyContinue)) {
    $EDHM_UI = $false
    Test-Path-Prompt ('EDHM UI')
    $key = $host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown').Character
    if ($key -eq 'y' -or $key -eq 'Y') {
        # Define the URL to download the MSI file from
        $url = 'https://github.com/BlueMystical/EDHM_UI/releases/latest'
        MSIinstall ($url)
    }
}
if (-not(Test-Path -Path $Path_VoiceAttack -PathType Any -ErrorAction SilentlyContinue)) {
    $VoiceAttack = $false
    Write-Host "Voice Attack's website didn't facilitate programmatic installs right out of the gate. Best I can do for now is send you to the official Downloads page. Proceed?"
    $key = $host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown').Character
    if ($key -eq 'y' -or $key -eq 'Y') {
        # Define the URL to navigate to
        $url = 'https://voiceattack.com/Downloads.aspx'
        Start-Process $url -ErrorAction SilentlyContinue
    }
}
if (-not(Test-Path -Path $Path_EDMarketConnector -PathType Any -ErrorAction SilentlyContinue)) {
    $EDMarketConnector = $false
    Test-Path-Prompt ('ED Market Connector')
    $key = $host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown').Character
    if ($key -eq 'y' -or $key -eq 'Y') {
        # Define the URL to download the MSI file from
        $url = 'https://github.com/EDCD/EDMarketConnector/releases/latest'
        MSIinstall ($url)
    }
}
if (-not(Test-Path -Path $Path_EDDiscovery -PathType Any -ErrorAction SilentlyContinue)) {
    $EDDiscovery = $false
    Test-Path-Prompt ('ED Discovery')
    $key = $host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown').Character
    if ($key -eq 'y' -or $key -eq 'Y') { <#TODO: install EDDiscovery#> }
}
if (-not(Test-Path -Path $Path_EDEngineer -PathType Any -ErrorAction SilentlyContinue)) {
    $EDEngineer = $false
    Test-Path-Prompt ('ED Engineer')
    $key = $host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown').Character
    if ($key -eq 'y' -or $key -eq 'Y') { <#TODO: install EDEngineer#> }
}
if (-not(Test-Path -Path $Path_EliteObservatory -PathType Any -ErrorAction SilentlyContinue)) {
    $EliteObservatory = $false
    Test-Path-Prompt ('Elite Observatory')
    $key = $host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown').Character
    if ($key -eq 'y' -or $key -eq 'Y') { <#TODO: install EliteObservatory#> }
}
if (-not(Test-Path -Path $Path_EliteTrack -PathType Any -ErrorAction SilentlyContinue)) {
    $EliteTrack = $false
    if ($InstallerMode) {
        Test-Path-Prompt ('Elite Track')
        $key = $host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown').Character
        if ($key -eq 'y' -or $key -eq 'Y') { <#TODO: install EliteTrack#> }
    }
}
if (-not(Get-ItemProperty 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\Steam' -ErrorAction SilentlyContinue)) {
    if ($InstallerMode) {
        Test-Path-Prompt ('Steam')
        $key = $host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown').Character
        if ($key -eq 'y' -or $key -eq 'Y') {
            Write-Output 'Installing Steam...'
            # Check if Chocolatey is installed
            if (-not (Get-Command choco.exe -ErrorAction SilentlyContinue)) {
                Write-Host "Chocolatey is not installed, but requires admin priviliges in order to install and we don't have those."
                #TODO: add external chocolatey install call intiate an admin console
            }
            #choco install steam -y
        }
    }
}

### "Local" Software ###
if ($EDHM_UI) { Start-SecondScreen($Path_EDHM_UI) }
if ($VoiceAttack) { Start-SecondScreen($Path_VoiceAttack) }

### Community Data ###
if ($EDMarketConnector) { Start-SecondScreen($Path_EDMarketConnector) }
if ($EDDiscovery) { Start-SecondScreen($Path_EDDiscovery) }
if ($EDEngineer) { Start-SecondScreen($Path_EDEngineer) }
if ($EliteObservatory) { Start-SecondScreen($Path_EliteObservatory) }

### Streaming ###
if ($EliteTrack) { Start-SecondScreen($Path_EliteTrack) }

### Elite Dangerous ###
if ($EliteDangerous) { Start-Process -WindowStyle Maximized -ErrorAction SilentlyContinue -FilePath $Path_EliteDangerous }