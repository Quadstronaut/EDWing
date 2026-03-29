# continue-pg.ps1 - Automate Continue -> Private Group -> Launch for each client
. "$PSScriptRoot\MouseUtil.ps1"

# Each action: coordinates + click type + comment for maintainability
# Bi = CMDRBistronaut, Tri = CMDRTristronaut, Quad = CMDRQuadstronaut, Duv = CMDRDuvrazh
$actions = @(
    # --- Bistronaut ---
    @{ X = -960; Y = -141; ClickType = "Double" },  # Continue
    @{ X = -800; Y = -42;  ClickType = "Double" },  # PG
    @{ X = -731; Y = -123; ClickType = "Double" },  # Select Quad PG
    @{ X = -469; Y = -93;  ClickType = "Double" },  # Launch Session
    # --- Tristronaut ---
    @{ X = -960; Y = 461;  ClickType = "Double" },  # Continue
    @{ X = -800; Y = 555;  ClickType = "Double" },  # PG
    @{ X = -731; Y = 478;  ClickType = "Double" },  # Select Quad PG
    @{ X = -469; Y = 507;  ClickType = "Double" },  # Launch Session
    # --- Quadstronaut ---
    @{ X = -960; Y = 1061; ClickType = "Double" },  # Continue
    @{ X = -800; Y = 1111; ClickType = "Double" },  # PG
    @{ X = -471; Y = 1108; ClickType = "Double" },  # Launch Session
    # --- Duvrazh (primary monitor) ---
    @{ X = 277;  Y = 377;  ClickType = "Double" },  # Continue
    @{ X = 666;  Y = 666;  ClickType = "Double" },  # PG
    @{ X = 835;  Y = 525;  ClickType = "Double" },  # Select Quad PG
    @{ X = 1480; Y = 490;  ClickType = "Double" }   # Launch Session
)

foreach ($action in $actions) {
    Invoke-ClickAction -X $action.X -Y $action.Y -ClickType $action.ClickType
}
