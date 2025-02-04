param (
    [string]$nwfPath,
    [string]$nwdPath
)

# Ensure Navisworks is installed
$navisworksPath = "C:\Program Files\Autodesk\Navisworks Manage 2024\Roamer.exe"

# Test pathing
if (-Not (Test-Path $navisworksPath)) {
    Write-Host "Error during Conversion: Navisworks is not installed!"
    exit 1
}
if (-Not (Test-Path $nwfPath)) {
    Write-Host "Error during Conversion: NWF file not found!"
    exit 1
}

$arguments = "-OpenFile " + $nwfPath + " -NoGui -SaveFile " + $nwdPath + " -Exit"

# convert to nwd and save
Start-Process -filepath $navisworksPath -Argumentlist $arguments

# Wait until the temporary .nwd~ file disappears
$tempFile = $nwdPath + "~"
Write-Host "Waiting for temporary file to be removed: $tempFile"

# Stall until the temporary file no longer exists
while (Test-Path $tempFile) {
    Start-Sleep -Seconds 1
}

Write-Host "Success: NWD generated $nwdPath"