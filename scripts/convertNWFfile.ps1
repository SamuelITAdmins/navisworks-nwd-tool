param (
    [string]$navisworksPath,
    [string]$nwfPath,
    [string]$nwdPath
)

# Test pathing
if (-Not (Test-Path $navisworksPath)) {
    Write-Output "Navisworks is not installed!"
    exit 1
}
if (-Not (Test-Path $nwfPath)) {
    Write-Output "NWF file not found! This project is likely not a Navisworks Project."
    exit 1
}

# convert to nwd and save
Write-Output "Beginning conversion of NWF into NWD..."
$arguments = "-OpenFile `"$nwfPath`" -NoGui -SaveFile `"$nwdPath`" -Exit"
Start-Process -filepath $navisworksPath -Argumentlist $arguments

# Stall until the temporary file that does the conversion is created
Write-Output "Opening NWF..."
$tempPath = $nwdPath + "~"
$waiting = 180 # variable: the number of seconds to wait for the temp file to be created

# If this is the first time making the nwd, then change the temp file path
if (-Not (Test-Path $nwdPath)) {
    $tempPath = $nwdPath
    $firstConversion = $true
}

# wait for NWF to open and the temp file to be generated
while (-Not (Test-Path $tempPath) -and ($waiting -gt 0)) {
    Start-Sleep -Seconds 1
    $waiting--
    Write-Output "Opening NWF: $waiting"
}

# Stall until convertion is complete
if (Test-Path $tempPath) {
    Write-Output "Converting..."
    # if first nwd creation, then wait 15 seconds and complete script
    if ($firstConversion) {
        Start-Sleep -Seconds 15
        Write-Output "generated for the first time! You can open the NWD, but if it is temporarily corrupted then wait and reopen it in a few minutes."
        exit 0
    # otherwise, wait for temp file to finish conversion before completion
    } else {
        while (Test-Path $tempPath) {
            Start-Sleep -Seconds 1
        }
        Write-Output "generated."
        exit 0
    }
} else {
    Write-Output "began conversion. However, the NWF is large and takes awhile to open, so the NWD will be ready to open later."
    exit 0
}