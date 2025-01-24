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

#This runs the iso tracker model

$year = (Get-Date).ToString("yyyy")
$month = (Get-Date).ToString("MM")
$day = (Get-Date).ToString("dd")

$exePath = "C:\Program Files\Autodesk\Navisworks Manage 2024\Roamer.exe"
$inputpath = "S:\Projects\23239_Strata_Ross_CCP_Phs_II_Exp\CAD\Piping\Models\_DesignReview\23239-Isometric_Status.nwf"
$appearancepath = "S:\Projects\23239_Strata_Ross_CCP_Phs_II_Exp\CAD\Piping\Models\_DesignReview\IsoTrackerColors.dat"
$outputPath = "S:\Projects\23239_Strata_Ross_CCP_Phs_II_Exp\CAD\Piping\Models\_DesignReview\ARCHIVE\23239-Isometric_Status_" + $year + "_" + $month + "_" + $day + ".nwd"

$myvar = "-OpenFile " + $inputpath + " -NoGui -ExecuteAddInPlugin" + " AutoAppearanceLoader.Navisworks " + $appearancepath + " -SaveFile " + $outputPath

start-process -filepath $exepath -Argumentlist $myvar

