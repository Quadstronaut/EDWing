<#### User Configuration ####
For each of the following programs, set them to true or false to determine if the script will attempt to launch them or not.
The script will automatically ignore lines for software that does not exist on your machine.
#>
$EDHM_UI = $true
$VoiceAttack = $false
$EDMarketConnector = $true
$EDDiscovery = $false
$EDEngineer = $false
$EliteObservatory = $false
$EliteTrack = $false
$EliteDangerous = $true
#### Do not edit beyond this line ####

### Sanity ###
if ($PSVersionTable.PSVersion.Major -lt 5) { Write-Error 'PowerShell version less than 5' -ErrorAction Stop }

### Paths ###
$Path_EDHM_UI = "$HOME\AppData\Local\EDHM_UI\EDHM_UI_mk2.exe"
$Path_VoiceAttack = 'C:\Program Files\VoiceAttack\VoiceAttack.exe'
$Path_EDMarketConnector = 'C:\Program Files (x86)\EDMarketConnector\EDMarketConnector.exe'
$Path_EDDiscovery = 'C:\Program Files\EDDiscovery\EDDiscovery.exe'
$Path_EDEngineer = Get-ChildItem -Path "$HOME\AppData" -Filter "EDEngineer.exe" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
$Path_EliteObservatory = 'C:\Program Files\Elite Observatory\ObservatoryCore.exe'
$Path_EliteTrack = "$HOME\AppData\Local\Programs\EliteTrack\EliteTrack.exe"
$Path_EliteDangerous = 'steam://rungameid/359320'
<#
$Path_EDHM_UI = Get-ChildItem -Path $env:AppData -Filter 'EDHM_UI_mk2.exe' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
$Path_VoiceAttack = Get-ChildItem -Path $env:ProgramFiles -Filter 'VoiceAttack.exe' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
$Path_EDMarketConnector = Get-ChildItem -Path "$env:ProgramFiles(x86)" -Filter 'EDMarketConnector.exe' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
$Path_EDDiscovery = Get-ChildItem -Path $env:ProgramFiles -Filter 'EDDiscovery.exe' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
$Path_EDEngineer = Get-ChildItem -Path $env:AppData -Filter 'EDEngineer.exe' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
$Path_EliteObservatory = Get-ChildItem -Path $env:ProgramFiles -Filter 'ObservatoryCore.exe' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
$Path_EliteTrack = Get-ChildItem -Path $env:AppData -Filter 'EliteTrack.exe' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
$Path_EliteDangerous = 'steam://rungameid/359320'
$Path_Steam = Get-ChildItem -Path $env:ProgramFiles(x86) -Filter 'Steam.exe' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
#>

### Functions ###
function Start-EliteApp ($app_path) {
    if ((Test-Path -Path $app_path -PathType Leaf) -or ($app_path.substring(0, 5) -eq 'steam'))
    { Start-Process -WindowStyle Maximized -ErrorAction SilentlyContinue -FilePath $app_path } 
}

### "Local" Software ###
if ($EDHM_UI) { if (Test-Path $Path_EDHM_UI) { Start-EliteApp($Path_EDHM_UI) } }
if ($VoiceAttack) { if (Test-Path $Path_VoiceAttack) { Start-EliteApp($Path_VoiceAttack) } }

### Community Data ###
if ($EDMarketConnector) { if (Test-Path $Path_EDMarketConnector) { Start-EliteApp($Path_EDMarketConnector) } }
if ($EDDiscovery) { if (Test-Path $Path_EDDiscovery) { Start-EliteApp($Path_EDDiscovery) } }
if ($EDEngineer) { if (Test-Path $Path_EDEngineer) { Start-EliteApp($Path_EDEngineer) } }
if ($EliteObservatory) { if (Test-Path $Path_EliteObservatory) { Start-EliteApp($Path_EliteObservatory) } }

### Streaming ###
if ($EliteTrack) { if (Test-Path $Path_EliteTrack) { Start-EliteApp($Path_EliteTrack) } }

### Elite Dangerous ###
#if ($EliteDangerous) { if (Test-Path $Path_Steam -PathType Leaf) { Start-EliteApp($Path_EliteDangerous) } }
if ($EliteDangerous) { Start-EliteApp($Path_EliteDangerous) }