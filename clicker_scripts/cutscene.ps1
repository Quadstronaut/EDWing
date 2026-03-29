# cutscene.ps1 - Skip intro cutscene by double-clicking each client window
. "$PSScriptRoot\MouseUtil.ps1"

# Coordinates to double-click (one per client window center)
$coordinates = @(
    @{ X = -700; Y = -100 },
    @{ X = -700; Y = 555  },
    @{ X = -700; Y = 1111 },
    @{ X = 999;  Y = 555  }
)

foreach ($coord in $coordinates) {
    Invoke-DoubleClick -X $coord.X -Y $coord.Y
    Start-Sleep -Seconds 1
}
