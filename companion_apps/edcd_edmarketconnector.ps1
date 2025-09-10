### User Config ###
$filenamePattern = "EDMarketConnector_Installer_*.exe"
$preRelease = $false
$repo = "edcd/edmarketconnector"

# Ask GitHub for the target release and download to temp
if ($preRelease) {
    $releasesUri = "https://api.github.com/repos/$repo/releases"
    $downloadUri = ((Invoke-RestMethod -Method GET -Uri $releasesUri)[0].assets | Where-Object name -like $filenamePattern ).browser_download_url
}
else {
    $releasesUri = "https://api.github.com/repos/$repo/releases/latest"
    $downloadUri = ((Invoke-RestMethod -Method GET -Uri $releasesUri).assets | Where-Object name -like $filenamePattern ).browser_download_url
}

# Download the release - [0] is signed [1] is unsigned - signed is default
$pathInstaller = Join-Path -Path $([System.IO.Path]::GetTempPath()) -ChildPath $(Split-Path -Path $downloadUri[0] -Leaf)
Invoke-WebRequest -Uri $downloadUri[0] -Out $pathInstaller

# Silently install
if (Test-Path -Path $pathInstaller -PathType Leaf) {
    Start-Process -FilePath $pathInstaller -ArgumentList "/VERYSILENT" -Wait -NoNewWindow -Verbose
}

# Delete installer
Remove-Item -Path $pathInstaller -Force -ErrorAction SilentlyContinue