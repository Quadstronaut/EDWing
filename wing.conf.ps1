# This file is used to store configuration data for the
# wing.ps1 script, such as CMDR names.
# To use it, place it in the same directory as the PowerShell script.

# Define an array of CMDR names to be used in the script.
# Add or remove names from this list to change the number of windows launched.
# The order here determines the order in which they are assigned to Elite Dangerous and EDMC instances.
$cmdrNames = @(
    "CMDRDuvrazh",
    "CMDRBistronaut",
    "CMDRTristronaut",
    "CMDRQuadstronaut"
)

# Alt Elite Dangerous commander names (excluding the first CMDR from the list)
$eliteDangerousCmdrs = $cmdrNames | Select-Object -Skip 1

# Define the paths for your executables
# These can be moved here to keep the main script cleaner
$sandboxieStart = 'C:\Users\Quadstronaut\scoop\apps\sandboxie-plus-np\current\Start.exe'
$edhm_uiLauncher = 'C:\Users\Quadstronaut\AppData\Local\EDHM-UI-V3\EDHM-UI-V3.exe'
$edminLauncher = 'G:\SteamLibrary\steamapps\common\Elite Dangerous\MinEdLauncher.exe'