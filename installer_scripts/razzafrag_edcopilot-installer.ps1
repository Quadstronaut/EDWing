### User Config ###
$filenamePattern = "*.msi"
$repo = "Razzafrag/EDCoPilot-Installer"

# Ask GitHub for the target release and download to temp
$releasesUri = "https://api.github.com/repos/$repo/releases/latest"
$downloadUri = ((Invoke-RestMethod -Method GET -Uri $releasesUri).assets | Where-Object name -like $filenamePattern ).browser_download_url

# Download latest release
$pathMSI = Join-Path -Path $([System.IO.Path]::GetTempPath()) -ChildPath $(Split-Path -Path $downloadUri -Leaf)
Invoke-WebRequest -Uri $downloadUri -Out $pathMSI

# Silently install
if (Test-Path -Path $pathMSI -PathType Leaf) {
    $msiArguments = @(
        "/i",
        "$pathMSI",
        "TARGETDIR=G:\EliteApps\"
    )
    Start-Process -FilePath "msiexec.exe" -ArgumentList $msiArguments -Wait -Verbose #TODO: This does not work.
}

# Clean up
Remove-Item $pathMSI -Force