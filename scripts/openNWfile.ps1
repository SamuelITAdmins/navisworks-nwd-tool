param (
    [string]$navisworksAppPath,
    [string]$sourcePath,
    [string]$destinationPath
)

# DestinationFile = "C:\Navisworks\#####-OverallModel.nwf"
# SourceFile = "\\stor-dn-01\projects\Projects\24317_Electra_CO_EPCM\CAD\Piping\Models\_DesignReview\24317-OverallModel.nwf"

$navisworksPath = "C:\Navisworks\"

# Test pathing
if (-Not (Test-Path $navisworksAppPath)) {
    Write-Host "Error during Opening: Navisworks is not installed!"
    exit 1
}
if (-Not (Test-Path $sourcePath)) {
    Write-Host "Error during Opening: Source File not found! $sourcePath"
    exit 1
}
# Create the navisworks folder if it does not already exist
If (-Not (Test-Path $navisworksPath)) {
    $navisworksFolder = New-Item -Path $navisworksPath -ItemType Directory
}

if (Test-Path $destinationPath) {
  $destinationFile = Get-Item $destinationPath
  if ($destinationFile.IsReadOnly) {
    # The file exists and is not read-only.  Safe to replace the file.
    Copy-Item -Path $sourcePath -Destination $destinationPath -Force
  } else {
    # The file exists and is read-only.
    Set-ItemProperty -Path $destinationPath -Name IsReadOnly -Value $false
    Copy-Item -Path $sourcePath -Destination $destinationPath -Force
    Set-ItemProperty -Path $destinationPath -Name IsReadOnly -Value $true
  }
} else {
  # The file does not exist in the destination folder. Safe to copy file to this folder.
  Copy-Item -Path $sourcePath -Destination $destinationPath
}

# open the copied navisworks file
$arguments = "-OpenFile " + $destinationPath + " -ShowGui"
Start-Process -filepath $navisworksAppPath -Argumentlist $arguments

Write-Host "Success: Opened $sourcePath locally as $destinationPath"
exit 0