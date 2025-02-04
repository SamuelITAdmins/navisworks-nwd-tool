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

# convert to nwd and save
Write-Host "Beginning conversion of $nwfPath into $nwdPath"
$arguments = "-OpenFile " + $nwfPath + " -NoGui -SaveFile " + $nwdPath + " -Exit"
Start-Process -filepath $navisworksPath -Argumentlist $arguments

# Stall until the temporary file that does the conversion is created
$tempFile = $nwdPath + "~"
$waiting = 120 # Hyperparameter: the number of seconds to wait for the temp file to be created
while (-Not (Test-Path $tempFile)) {
    Start-Sleep -Seconds 1
    $waiting--
    if ($waiting -eq 0) {
        Write-Host "Error during Conversion: Conversion did not begin in time, aborting..."
        exit 1
    }
}

# Stall until the temporary file no longer exists
Write-Host "Waiting for temporary file to be removed: $tempFile"
while (Test-Path $tempFile) {
    Start-Sleep -Seconds 1
}

Write-Host "Success: NWD generated $nwdPath"