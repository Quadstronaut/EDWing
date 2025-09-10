### User Config ###
$DestinationPath = 'G:\EliteApps\'

$websiteUrl = "https://www.panostrede.de/EDEB/"
$htmlContent = Invoke-WebRequest -Uri $websiteUrl -UseBasicParsing -UserAgent "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
$msiLink = $htmlContent.Links | Where-Object { $_.href -match '\.msi$' }

# Check if the link was found.
if (-not $msiLink) {
    throw "'.msi' not found."
}
else {
    # Link is relative, prepend base
    $downloadUrl = $msiLink.href
    if ($downloadUrl -notmatch "https?://") {
        $uri = New-Object System.Uri($websiteUrl)
        $downloadUrl = New-Object System.Uri($uri, $msiLink.href).AbsoluteUri
    }

    # Ensure the directory exists.
    if (-not (Test-Path $DestinationPath)) {
        New-Item -ItemType Directory -Force -Path $DestinationPath | Out-Null
    }

    # Extract the filename from the URL.
    $fileName = $msiLink.href
    $filePath = Join-Path -Path $DestinationPath -ChildPath $fileName
    Invoke-WebRequest -Uri $downloadUrl -OutFile $filePath
}