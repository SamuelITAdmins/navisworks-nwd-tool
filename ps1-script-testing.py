import subprocess

# Define the paths
path1 = "C:\\Path\\To\\First\\Folder"
path2 = "C:\\Path\\To\\Second\\Folder"

# Define the PowerShell script path
powershell_script = "C:\\Path\\To\\script.ps1"

# Construct the PowerShell command
command = ["powershell", "-ExecutionPolicy", "Bypass", "-File", powershell_script, path1, path2]

# Run the PowerShell script
result = subprocess.run(command, capture_output=True, text=True)

# Print the output and errors (if any)
print("Output:", result.stdout)
print("Errors:", result.stderr)
