# Navisworks Utility App

This application is designed to help users generate NWD files and open NWD/NWF files from the SE project folder using a graphical interface.

## Table of Contents
- [Usage (For Users)](#usage-for-users)
- [Developer Guide](#developer-guide)
- [Building the Executable](#building-the-executable)

## Usage (For Users)

### Running the Application
1. **Open the Executable:** Run the `navisworks-utility-app.exe` file.
2. **Select a Project:** Use the dropdown menu to choose the desired project.
3. **Available Actions:**
   - **Generate NWD:** Converts the NWF file to an NWD file.
   - **Open NWD:** Opens the NWD file locally.
   - **Open NWF:** Opens the NWF file locally.
   - **Refresh List:** Reloads the project list in case of updates.

### Expected File Locations
- NWF files should be in:  
  `S:\Projects\<Project_Name>\CAD\Piping\Models\_DesignReview\`
- Generated NWD files will be saved in the same directory with the DO_NOT_OPEN tag.
- Opened NWF/NWD files will be opened locally in:
  `C:\Navisworks\`

## Developer Guide

### Project Structure
navisworks-utility-app/ 
│-- navisworks-utility-app.py # Main application script 
│-- scripts/ 
│  ├── convertNWFfile.ps1 # PowerShell script for NWD generation 
│  ├── openNWfile.ps1 # PowerShell script for opening Navisworks files 
│-- test-scripts/ 
│  ├── ps1-script-testing.py # python test script for testing powershell scripts
│  ├── vbs-script-testing.py # python test script for testing vbs scripts
│-- se.ico # Application icon 
│-- README.md # Documentation

### Key Components
- `NWGUI` Class: Manages the user interface and interactions.
- `generate_nwd()`: Calls the PowerShell script to convert NWF to NWD.
- `open_nwd() / open_nwf()`: Opens the respective files using PowerShell.
- `load_projects()`: Retrieves and sorts project names from the directory.

### Dependencies
- Python 3.x
- `pyinstaller`
- `tkinter`
- `subprocess`
- `os`
- `sys`
- `re`

### Setting Up for Development
1. **Clone the repository:**
   - `git clone <repo_url>`
   - `cd navisworks-utility-app`
2. **Install dependencies if needed:**
   - `pip install -r requirements.txt`
3. **Run the script:**
   - `python navisworks-utility-app.py`

## Building the Executable
To package the script into the executable file, use this PyInstaller command:
`pyinstaller --onefile --noconsole --icon=se.ico --add-data "scripts;scripts" navisworks-utility-app.py`

This will generate dist/navisworks-utility-app.exe, which can be distributed to users.
This will need to be run whenever a critical change is made to the codebase for production.
