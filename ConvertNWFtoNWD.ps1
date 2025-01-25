param (
    [string]$nwfFile
    [string]$nwdFile
)

# Ensure Navisworks is installed
$navisworksPath = "C:\Program Files\Autodesk\Navisworks Manage 2024\Roamer.exe"

# Test pathing
if (-Not (Test-Path $navisworksPath)) {
    Write-Host "Error during Conversion: Navisworks is not installed!"
    exit 1
}
if (-Not (Test-Path $nwfFile)) {
    Write-Host "Error during Conversion: NWF file not found!"
    exit 1
}

$arguments = "-OpenFile " + $nwfFile + " -NoGui -SaveFile " + $nwdFile + " -Exit"

# convert to nwd and save
Start-Process -filepath $navisworksPath -Argumentlist $arguments

Write-Host "NWD generated: $nwdFile"