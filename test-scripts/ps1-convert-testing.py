import os
import subprocess

# Get the Roamer.exe path (requires the getNWroamerfile.ps1 script to be working properly)
get_roamer_script = os.path.abspath("scripts/getNWroamerfile.ps1")
command = ["powershell", "-ExecutionPolicy", "Bypass", "-File", get_roamer_script]
result = subprocess.run(command, capture_output=True, text=True)
roamer_path = result.stdout.strip().split('\n')[-1]

# Define the paths
# - These paths are used for convertNWFfile.ps1
sourceFile = "S:\\Projects\\24206_BHM_Dust_Collection\\CAD\\Piping\\Models\\_DesignReview\\24206-OverallModel.nwf"
destFile = "S:\\Projects\\24206_BHM_Dust_Collection\\CAD\\Piping\\Models\\_DesignReview\\24206-SCRIPT_TEST.nwd"

# Define the PowerShell script path
powershell_script = os.path.abspath("scripts/convertNWFfile.ps1")

# Construct the PowerShell command
command = ["powershell", "-ExecutionPolicy", "Bypass", "-File", powershell_script, roamer_path, sourceFile, destFile]

# Run the PowerShell script
result = subprocess.run(command, capture_output=True, text=True)

# Print the output and errors (if any)
print("Output:", result.stdout)
print("Errors:", result.stderr)
