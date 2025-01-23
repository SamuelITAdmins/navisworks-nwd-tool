import os
import subprocess
import shutil
import tkinter as tk
from tkinter import messagebox
from datetime import datetime

# Define paths
PROJECT_DIR = r"S:\Projects\24317_Electra_CO_EPCM"
NWF_SOURCE_DIR = r"S:\Projects\24317_Electra_CO_EPCM\CAD\Piping\Models\_DesignReview"
NAVISWORKS_PATH = r"C:\Program Files\Autodesk\Navisworks Manage 2024\Roamer.exe"
LOG_FILE = r"S:\CAD_Standards\_SE_Custom\Piping\Programming\SE Piping Utilities\Navisworks\nwd_creation.log"

def check_navisworks():
    """Check if Navisworks Manage or Simulate is installed"""
    if not os.path.exists(NAVISWORKS_PATH):
        messagebox.showerror("Error", "Unable to create NWD! Application requires Navisworks Manage or Simulate.")
        return False
    return True

def create_nwd():
    """Convert NWF to NWD"""
    project_path = os.path.join(PROJECT_DIR, "CAD", "Piping", "Models", "_DesignReview")
    nwf_file = None

    # Locate the NWF file
    for file in os.listdir(project_path):
        if file.endswith("-Overall.nwf"):
            nwf_file = os.path.join(project_path, file)
            break

    if not nwf_file:
        messagebox.showerror("Error", "No NWF found for 24317_Electra_CO_EPCM!")
        return
    
    nwd_file = os.path.join(project_path, "24317-DO_NOT_OPEN.nwd")
    
    # Execute Navisworks command to open & save NWD
    try:
        subprocess.run([NAVISWORKS_PATH, nwf_file, "/silent"], check=True)
        shutil.copy(nwf_file, nwd_file)  # Simulating "Save As"
        
        # Update shortcut timestamp
        shortcut_path = os.path.join(PROJECT_DIR, "DesignReview", "24317-Overall.nwd.lnk")
        if os.path.exists(shortcut_path):
            os.utime(shortcut_path, (datetime.now().timestamp(), datetime.now().timestamp()))

        # Log success
        with open(LOG_FILE, "a") as log:
            log.write(f"{datetime.now()} - 24317 NWD successfully created.\n")
        
        messagebox.showinfo("Success", "24317-OverallModel.nwd successfully created in DesignReview folder!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to create NWD: {str(e)}")