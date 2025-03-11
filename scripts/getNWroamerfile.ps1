# Function to find Roamer.exe inside Autodesk's installation directories
$installDirs = @(
    "$env:ProgramFiles\Autodesk",
    "$env:ProgramW6432\Autodesk"
)

# Establish a priority for the Navisworks type
$foundPaths = @{
    "Manage"   = $null
    "Simulate" = $null
    "Freedom"  = $null
}

foreach ($dir in $installDirs) {
    if (Test-Path $dir) {
        # Search for directories containing 'Navisworks' in the name and sort by last modified date (newest first)
        $navisworksDirs = Get-ChildItem -Path $dir -Directory -Filter "Navisworks*" -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending

        foreach ($nwDir in $navisworksDirs) {
            $roamerPath = Join-Path -Path $nwDir.FullName -ChildPath "Roamer.exe"
            if (Test-Path $roamerPath) {
                # Store keys separately to prevent modification issues
                $priorities = @("Manage", "Simulate", "Freedom")
                foreach ($priority in $priorities) {
                    if ($nwDir.Name -match $priority -and -not $foundPaths[$priority]) {
                        $foundPaths[$priority] = $roamerPath
                    }
                }
            }
        }
    }
}

# Return the highest-priority version found
foreach ($priority in $foundPaths.Keys) {
    if ($foundPaths[$priority]) {
        Write-Host "Success: Found Navisworks Application as" $foundPaths[$priority]
        Write-Output $foundPaths[$priority]
        exit 0
    }
}

# If Roamer.exe not found:
Write-Host "Navisworks Application Not Found. Install Navisworks in C:\ProgramFiles\Autodesk."
exit 1