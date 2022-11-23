### Local ###
# EDHM_UI
if (!(Get-Process -Name EDHM*)) { Start-Process -WindowStyle Maximized -ErrorAction Continue -FilePath 'C:\Users\Quadstronaut\AppData\Local\EDHM_UI\EDHM_UI_mk2.exe' }
# VoiceAttack
# if (Get-Process -Name VoiceAttack) { $true } else { Start-Process -WindowStyle Maximized -ErrorAction Continue -FilePath "C:\Program Files\VoiceAttack\VoiceAttack.exe" }

### Community Data ###
# ED Market Connector
if (!(Get-Process -Name EDMarketConnector)) { Start-Process -WindowStyle Maximized -ErrorAction Continue -FilePath 'C:\Program Files (x86)\EDMarketConnector\EDMarketConnector.exe' }
# ED Discovery
if (!(Get-Process -Name EDDiscovery)) { Start-Process -WindowStyle Maximized -ErrorAction Continue -FilePath 'C:\Program Files\EDDiscovery\EDDiscovery.exe' }
# ED Engineer
if (!(Get-Process -Name EDEngineer)) { Start-Process -WindowStyle Maximized -ErrorAction Continue -FilePath 'C:\Users\Quadstronaut\AppData\Local\Apps\2.0\D96VD2JJ.7JH\585QYJEO.5TD\eden..tion_b9c6c2d0b4f2eae5_0001.0001_37a5eebcaa7d7023\EDEngineer.exe' }
# Elite Observatory
if (!(Get-Process -Name ObservatoryCore)) { Start-Process -WindowStyle Maximized -ErrorAction Continue -FilePath 'C:\Program Files\Elite Observatory\ObservatoryCore.exe' }

### Twitch ###
# EliteTrack
if (!(Get-Process -Name EDHM*)) { Start-Process -WindowStyle Maximized -ErrorAction Continue -FilePath 'C:\Users\Quadstronaut\AppData\Local\Programs\EliteTrack\EliteTrack.exe' }

### Elite Dangerous ###
if (!(Get-Process -Name EliteDangerous64 )) { Start-Process -WindowStyle Maximized -ErrorAction Continue -FilePath 'steam://rungameid/359320' }

Exit
