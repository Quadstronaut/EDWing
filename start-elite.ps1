<#### User Configuration ####
For each of the following programs, set them to true or false to determine if the script will attempt to launch them or not.
The script will automatically ignore lines for software that does not exist on your machine.
#>
$EDHM_UI = $true
$VoiceAttack = $false
$EDMarketConnector = $true
$EDDiscovery = $true
$EDEngineer = $true
$EliteObservatory = $true
$EliteTrack = $true
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

### Functions ###
function Start-EliteApp ($app_path) {
    if ((Test-Path -Path $app_path -PathType Leaf) -or ($app_path.substring(0, 5) -eq 'steam'))
    { Start-Process -WindowStyle Maximized -ErrorAction SilentlyContinue -FilePath $app_path } 
}

### "Local" Software ###
if ($EDHM_UI) { Start-EliteApp($Path_EDHM_UI) }
if ($VoiceAttack) { Start-EliteApp($Path_VoiceAttack) }

### Community Data ###
if ($EDMarketConnector) { Start-EliteApp($Path_EDMarketConnector) }
if ($EDDiscovery) { Start-EliteApp($Path_EDDiscovery) }
if ($EDEngineer) { Start-EliteApp($Path_EDEngineer) }
if ($EliteObservatory) { Start-EliteApp($Path_EliteObservatory) }

### Streaming ###
if ($EliteTrack) { Start-EliteApp($Path_EliteTrack) }

### Elite Dangerous ###
if ($EliteDangerous) { Start-EliteApp($Path_EliteDangerous) }
