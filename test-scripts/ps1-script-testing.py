import os
import subprocess

# Define the paths
# - These paths used for convertNWFfile.ps1
sourceFile = "S:\\Projects\\24317_Electra_CO_EPCM\\CAD\\Piping\\Models\\_DesignReview\\24317-OverallModel.nwf"
destFile = "S:\\Projects\\24317_Electra_CO_EPCM\\CAD\\Piping\\Models\\_DesignReview\\24317-SCRIPT_TEST.nwd"
# - These paths used for openNWfile.ps1
# sourceFile = "S:\\Projects\\24198_Chemtrade_ACP_FEL_3\\CAD\\Piping\\Models\\_DesignReview\\24198-DO_NOT_OPEN.nwd"
# sourceFile = "S:\\Projects\\24317_Electra_CO_EPCM\\CAD\\Piping\\Models\\_DesignReview\\24317-OverallModel.nwf"
# destFile = "C:\\Navisworks\\24317-OverallModel.nwf" # switch between nwf/nwd

# Define the PowerShell script path
powershell_script = os.path.abspath("scripts/convertNWFfile.ps1")
# powershell_script = os.path.abspath("scripts/openNWfile.ps1")

# Construct the PowerShell command
command = ["powershell", "-ExecutionPolicy", "Bypass", "-File", powershell_script, sourceFile, destFile]

# Run the PowerShell script
result = subprocess.run(command, capture_output=True, text=True)

# Print the output and errors (if any)
print("Output:", result.stdout)
print("Errors:", result.stderr)
