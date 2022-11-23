#### User Configuration ####

<# For each of the following programs, set them to true or false to determine if the script will attempt to launch them or not.
# The script will automatically ignore lines for software that does not exist on your machine. #>

$EDHM_UI = $true
$VoiceAttack = $true
$EDMarketConnector = $true
$EDDiscovery = $true
$EDEngineer = $true
$EliteObservatory = $true
$EliteTrack = $false
$EliteDangerous = $true

#### Do Not Edit Beyond This Line ####

### Local ###

# EDHM_UI
if ($EDHM_UI) {
    if (!(Get-Process -Name EDHM* -ErrorAction SilentlyContinue)) { 
        Start-Process -WindowStyle Maximized -ErrorAction SilentlyContinue -FilePath '$env:USERPROFILE\AppData\Local\EDHM_UI\EDHM_UI_mk2.exe' 
    } 
}

# VoiceAttack
if ($VoiceAttack) {
    if (!(Get-Process -Name VoiceAttack -ErrorAction SilentlyContinue)) {
        Start-Process -WindowStyle Maximized -ErrorAction SilentlyContinue -FilePath 'C:\Program Files\VoiceAttack\VoiceAttack.exe' 
    } 
}

### Community Data ###

# ED Market Connector
if ($EDMarketConnector) {
    if (!(Get-Process -Name EDMarketConnector -ErrorAction SilentlyContinue)) {
        Start-Process -WindowStyle Maximized -ErrorAction SilentlyContinue -FilePath 'C:\Program Files (x86)\EDMarketConnector\EDMarketConnector.exe' 
    } 
}

# ED Discovery
if ($EDDiscovery) {
    if (!(Get-Process -Name EDDiscovery -ErrorAction SilentlyContinue)) {
        Start-Process -WindowStyle Maximized -ErrorAction SilentlyContinue -FilePath 'C:\Program Files\EDDiscovery\EDDiscovery.exe' 
    } 
}

# ED Engineer
if ($EDEngineer) {
    if (!(Get-Process -Name EDEngineer -ErrorAction SilentlyContinue)) {
        Start-Process -WindowStyle Maximized -ErrorAction SilentlyContinue -FilePath '$env:USERPROFILE\AppData\Local\Apps\2.0\D96VD2JJ.7JH\585QYJEO.5TD\eden..tion_b9c6c2d0b4f2eae5_0001.0001_37a5eebcaa7d7023\EDEngineer.exe' 
    }
}

# Elite Observatory
if ($EliteObservatory) {
    if (!(Get-Process -Name ObservatoryCore -ErrorAction SilentlyContinue)) {
        tart-Process -WindowStyle Maximized -ErrorAction SilentlyContinue -FilePath 'C:\Program Files\Elite Observatory\ObservatoryCore.exe' 
    }
}

### Streaming ###

# EliteTrack
if ($EliteTrack) {
    if (!(Get-Process -Name EliteTrack* -ErrorAction SilentlyContinue)) {
        Start-Process -WindowStyle Maximized -ErrorAction SilentlyContinue -FilePath '$env:USERPROFILE\AppData\Local\Programs\EliteTrack\EliteTrack.exe' 
    }
}

### Elite Dangerous ###

if ($EliteDangerous) {
    if (!(Get-Process -Name EliteDangerous64 -ErrorAction SilentlyContinue)) {
        Start-Process -WindowStyle Maximized -ErrorAction SilentlyContinue -FilePath 'steam://rungameid/359320' 
    }
}

Exit
