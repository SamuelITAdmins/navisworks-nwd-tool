import os
import subprocess
import shutil
import tkinter as tk
from tkinter import messagebox
from datetime import datetime

# Define paths
NAVISWORKS_PATH = r"C:\Program Files\Autodesk\Navisworks Manage 2024\Roamer.exe"
LOG_FILE = r"S:\CAD_Standards\_SE_Custom\Piping\Programming\SE Piping Utilities\Navisworks\nwd_creation.log"
PROJECTS_DIR = r"S:\Projects"

def check_navisworks():
    """Check if Navisworks Manage or Simulate is installed"""
    if not os.path.exists(NAVISWORKS_PATH):
        messagebox.showerror("Error", "Unable to create NWD! Application requires Navisworks Manage or Simulate.")
        return False
    return True

def get_projects():
    """Fetch available projects"""
    return [f for f in os.listdir(PROJECTS_DIR) if os.path.isdir(os.path.join(PROJECTS_DIR, f))]

def create_nwd(project_name):
    """Convert NWF to NWD"""
    project_path = os.path.join(PROJECTS_DIR, project_name, "CAD", "Piping", "Models", "_DesignReview")
    nwf_file = None

    # Locate the NWF file
    for file in os.listdir(project_path):
        if file.endswith("-Overall.nwf"):
            nwf_file = os.path.join(project_path, file)
            break

    if not nwf_file:
        messagebox.showerror("Error", f"No NWF found for {project_name}!")
        return
    
    nwd_file = os.path.join(project_path, f"{project_name}-DO_NOT_OPEN.nwd")
    
    # Execute Navisworks command to open & save NWD
    try:
        subprocess.run([NAVISWORKS_PATH, nwf_file, "/silent"], check=True)
        shutil.copy(nwf_file, nwd_file)  # Simulating "Save As"
        
        # Update shortcut timestamp
        shortcut_path = os.path.join(PROJECTS_DIR, project_name, "DesignReview", f"{project_name}-Overall.nwd.lnk")
        if os.path.exists(shortcut_path):
            os.utime(shortcut_path, (datetime.now().timestamp(), datetime.now().timestamp()))

        # Log success
        with open(LOG_FILE, "a") as log:
            log.write(f"{datetime.now()} - {project_name} NWD successfully created.\n")
        
        messagebox.showinfo("Success", f"{project_name}-OverallModel.nwd successfully created in DesignReview folder!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to create NWD: {str(e)}")

def on_submit():
    """Handler for the submit button"""
    selected_project = project_listbox.get(tk.ACTIVE)
    if not selected_project:
        messagebox.showwarning("Warning", "Please select a project.")
        return

    if check_navisworks():
        create_nwd(selected_project)

# GUI Setup
root = tk.Tk()
root.title("Navisworks NWD Creator")

tk.Label(root, text="Select Project:").pack()

# Listbox for project selection
project_listbox = tk.Listbox(root, height=10)
for project in get_projects():
    project_listbox.insert(tk.END, project)
project_listbox.pack()

# Buttons
submit_button = tk.Button(root, text="Create NWD", command=on_submit)
submit_button.pack()

exit_button = tk.Button(root, text="Exit", command=root.quit)
exit_button.pack()

root.mainloop()
