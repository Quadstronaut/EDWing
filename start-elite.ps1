#### User Configuration ####

<# For each of the following programs, set them to true or false to determine if the script will attempt to launch them or not.
The script will automatically ignore lines for software that does not exist on your machine. #>

$EDHM_UI = $true
$VoiceAttack = $true
$EDMarketConnector = $true
$EDDiscovery = $true
$EDEngineer = $true
$EliteObservatory = $true
$EliteTrack = $true
$EliteDangerous = $true

#### Do not edit beyond this point ####

### Paths ###

$Path_EDHM_UI = "$HOME\AppData\Local\EDHM_UI\EDHM_UI_mk2.exe"
$Path_VoiceAttack = 'C:\Program Files\VoiceAttack\VoiceAttack.exe'
$Path_EDMarketConnector = 'C:\Program Files (x86)\EDMarketConnector\EDMarketConnector.exe'
$Path_EDDiscovery = 'C:\Program Files\EDDiscovery\EDDiscovery.exe'
$Path_EDEngineer = "$HOME\AppData\Local\Apps\2.0\D96VD2JJ.7JH\585QYJEO.5TD\eden..tion_b9c6c2d0b4f2eae5_0001.0001_37a5eebcaa7d7023\EDEngineer.exe" #TODO: Cannot be hardcoded
$Path_EliteObservatory = 'C:\Program Files\Elite Observatory\ObservatoryCore.exe'
$Path_EliteTrack = "$HOME\AppData\Local\Programs\EliteTrack\EliteTrack.exe" 
$Path_EliteDangerous = 'steam://rungameid/359320'

### Functions ###

function Exec ($app) { if (Installed($app)) { Start-Process -WindowStyle Maximized -ErrorAction SilentlyContinue -FilePath $app } }

function Installed ($app) { Test-Path -Path $app -PathType Leaf && return $true || return $false }

### "Local" Software ###
if ($EDHM_UI) { Exec($Path_EDHM_UI) }
if ($VoiceAttack) { Exec($Path_VoiceAttack) }

### Community Data ###
if ($EDMarketConnector) { Exec($Path_EDMarketConnector) }
if ($EDDiscovery) { Exec($Path_EDDiscovery) }
if ($EDEngineer) { Exec($Path_EDEngineer) }
if ($EliteObservatory) { Exec($Path_EliteObservatory) }

### Streaming ###
if ($EliteTrack) { Exec($Path_EliteTrack) }

### Elite Dangerous ###
if ($EliteDangerous) { Exec($Path_EliteDangerous) }