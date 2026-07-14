#Requires -Version 5.1
# Get-CursorPos.ps1 - capture screen coordinates for per-CMDR menu click sets.
#
# Mode-2 arranges every client full-size on the PRIMARY monitor, so all click
# coordinates are primary-monitor pixels. To capture a CMDR's menu buttons:
#   1. Launch and stack the clients (.\Get-Wing.ps1), bring one CMDR forward
#      (Show-Cmdr -Box CMDRBistronaut) so its menu is visible on the primary monitor.
#   2. Run this tool. Hover each button (Continue, Private Group, the group row,
#      Launch), press Enter to print a ready-to-paste line; type q + Enter to finish.
#   3. Paste the lines into $config.CmdrClickSets['CMDRBistronaut'] in wing.conf.ps1.
#
# Because these are per-account (the private-group list is ordered alphabetically and
# differs per commander), capture a set per CMDR.

Add-Type -AssemblyName System.Windows.Forms
Write-Host "Hover a target, press Enter to capture. Type q + Enter to quit." -ForegroundColor Cyan
while ($true) {
    $k = Read-Host 'capture (Enter=grab, q=quit)'
    if ($k -eq 'q') { break }
    $p = [System.Windows.Forms.Cursor]::Position
    Write-Host ("        @{{ X = {0}; Y = {1}; ClickType = 'Double' }}," -f $p.X, $p.Y) -ForegroundColor Green
}
