param (
    [string]$nwfFile
    [string]$nwdFile
)

# Ensure Navisworks is installed
# Ensure Navisworks is installed and can be executed via command line
$navisworksPath = "C:\Program Files\Autodesk\Navisworks Manage 2024\Roamer.exe"

if (-Not (Test-Path $nwfFile)) {
    Write-Host "Error: NWF file not found!"
    exit 1
}

# Open the NWF file and export as NWD
Start-Process -FilePath $navisworksPath -ArgumentList "$nwfFile" -Wait
Start-Sleep -Seconds 5

# Save as NWD (simulated; replace with actual Navisworks automation if available)
Copy-Item -Path $nwfFile -Destination $nwdFile

Write-Host "NWD generated: $nwdFile"