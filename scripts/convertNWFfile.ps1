param (
    [string]$nwfPath,
    [string]$nwdPath
)

# Ensure Navisworks is installed
$navisworksPath = "C:\Program Files\Autodesk\Navisworks Manage 2024\Roamer.exe"

# Test pathing
if (-Not (Test-Path $navisworksPath)) {
    Write-Output "Navisworks is not installed!"
    exit 1
}
if (-Not (Test-Path $nwfPath)) {
    Write-Output "NWF file not found! Likely not a Navisworks Project."
    exit 1
}

# convert to nwd and save
Write-Output "Beginning conversion of NWF into NWD..."
$arguments = "-OpenFile `"$nwfPath`" -NoGui -SaveFile `"$nwdPath`" -Exit"
Start-Process -filepath $navisworksPath -Argumentlist $arguments

# Stall until the temporary file that does the conversion is created
Write-Output "Opening NWF..."
$tempFile = $nwdPath + "~"
$waiting = 60 # variable: the number of seconds to wait for the temp file to be created
while (-Not (Test-Path $tempFile) -and ($waiting -gt 0)) {
    Start-Sleep -Seconds 1
    $waiting--
}

# Stall until the temporary file no longer exists
if (Test-Path $tempFile) {
    Write-Output "Converting..."
    while (Test-Path $tempFile) {
        Start-Sleep -Seconds 1
    }
    Write-Output "Success: NWD generated"
    exit 0
} else {
    Write-Output "Conversion has begun. The NWF is large and takes awhile to open, so the NWD will be ready to open later."
    exit 1
}