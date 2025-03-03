import os
import subprocess

# Define the PowerShell script path
powershell_script = os.path.abspath("scripts/getNWroamerfile.ps1")

# Construct the PowerShell command
command = ["powershell", "-ExecutionPolicy", "Bypass", "-File", powershell_script]

# Run the PowerShell script
result = subprocess.run(command, capture_output=True, text=True)

# Print the output and errors (if any)
print("Output:", result.stdout)
print("Errors:", result.stderr)
