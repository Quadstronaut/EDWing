### User Config ###
$elitePath = 'G:\SteamLibrary\steamapps\common\Elite Dangerous'
$filenamePattern = "*win-x64.zip"
$innerDirectory = $true
$pathExtract = "G:\EliteApps\min-ed-launcher"
$preRelease = $false
$repo = "rfvgyhn/min-ed-launcher"

# Ask GitHub for the target release and download to temp
if ($preRelease) {
    $releasesUri = "https://api.github.com/repos/$repo/releases"
    $downloadUri = ((Invoke-RestMethod -Method GET -Uri $releasesUri)[0].assets | Where-Object name -like $filenamePattern ).browser_download_url
}
else {
    $releasesUri = "https://api.github.com/repos/$repo/releases/latest"
    $downloadUri = ((Invoke-RestMethod -Method GET -Uri $releasesUri).assets | Where-Object name -like $filenamePattern ).browser_download_url
}

# Extract the release
$pathZip = Join-Path -Path $([System.IO.Path]::GetTempPath()) -ChildPath $(Split-Path -Path $downloadUri -Leaf)
Invoke-WebRequest -Uri $downloadUri -Out $pathZip

# Clear extraction container
Remove-Item -Path $pathExtract -Recurse -Force -ErrorAction SilentlyContinue

# Decompression
if ($innerDirectory) {
    $tempExtract = Join-Path -Path $([System.IO.Path]::GetTempPath()) -ChildPath $((New-Guid).Guid)
    Expand-Archive -Path $pathZip -DestinationPath $tempExtract -Force
    Move-Item -Path "$tempExtract\*" -Destination $pathExtract -Force
    Remove-Item -Path $tempExtract -Force -Recurse -ErrorAction SilentlyContinue
}
else {
    Expand-Archive -Path $pathZip -DestinationPath $pathExtract -Force
}

# Clean up
Remove-Item $pathZip -Force

# Symlink the file to the $elitePath
Copy-Item -Path $(Join-Path -Path $elitePath -ChildPath "MinEdLauncher.exe") -Value (Join-Path -Path $pathExtract -ChildPath "MinEdLauncher.exe") -Force