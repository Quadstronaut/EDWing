#Requires -Version 5.1
# Show-Cmdr.ps1 - bring one sandboxed CMDR's Elite window to the foreground (mode-2 primitive).
# Usage:  .\tools\Show-Cmdr.ps1 CMDRBistronaut
# Standalone (dot-sources WingLib itself) so it works from a normal prompt after Get-Wing exits.
param([Parameter(Mandatory)][string]$Box)
. "$PSScriptRoot\..\WingLib.ps1"
$w = Get-CmdrWindows -BoxNames @($Box) | Where-Object { $_.Box -eq $Box } | Select-Object -First 1
if (-not $w) { Write-Warning "No sandboxed window found for $Box"; exit 1 }
if (Set-CmdrForeground -Hwnd $w.Hwnd) { Write-Host "Focused $Box (hwnd 0x$('{0:X}' -f $w.Hwnd))" }
else { Write-Warning "Focus attempt for $Box did not take (try again, or check the window exists)" }
