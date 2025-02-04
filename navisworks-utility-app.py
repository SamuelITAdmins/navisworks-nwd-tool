import os
import re
import sys
import tkinter as tk
from tkinter import ttk, messagebox
import subprocess

# nwf file example: "\\stor-dn-01\projects\Projects\24317_Electra_CO_EPCM\CAD\Piping\Models\_DesignReview\24317-OverallModel.nwf"
# nwd file example: "\\stor-dn-01\projects\Projects\24317_Electra_CO_EPCM\CAD\Piping\Models\_DesignReview\24317-DO_NOT_OPEN.nwd"

# Paths (Modify as needed)
PROJECTS_DIR = r"S:\Projects"  # Directory containing all projects
NAVISWORKS_TEMP_DIR = r"C:\Navisworks" # Directory containing copied Navisworks files
NWF_EXT = ".nwf"
NWD_EXT = ".nwd"

def get_resource_path(relative_path):
    '''Get absolute path to a resource, works for development and PyInstaller builds'''
    if getattr(sys, '_MEIPASS', False): # If running from .exe file
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.abspath(relative_path) # When running locally (python cmd)

CONVERT_PS_SCRIPT = get_resource_path("scripts/convertNWFfile.ps1")
OPEN_PS_SCRIPT = get_resource_path("scripts/openNWfile.ps1")

class NWGUI:
    def __init__(self, root):
        self.root = root
        root.iconbitmap(get_resource_path("se.ico"))
        self.root.title("Navisworks Project Manager")

        # Dropdown Label
        ttk.Label(root, text="Select Project:").grid(row=0, column=0, padx=5, pady=5)

        # Dropdown for Projects
        self.project_num = ''
        self.project_name = tk.StringVar()
        self.project_dropdown = ttk.Combobox(root, textvariable=self.project_name, width=40, state="readonly")
        self.project_dropdown.grid(row=1, column=0, columnspan=3, padx=5, pady=5)
        self.load_projects()

        # Buttons
        ttk.Button(root, text="Generate NWD", command=self.generate_nwd).grid(row=2, column=0, padx=5, pady=5)
        ttk.Button(root, text="Open NWD", command=self.open_nwd).grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(root, text="Open NWF", command=self.open_nwf).grid(row=2, column=2, padx=5, pady=5)
        ttk.Button(root, text="Refresh List", command=self.load_projects).grid(row=3, column=0, padx=5, pady=5)
    
    def extract_project_num(self, project_name):
        """Extract the project number from the full project name, discard if the project name does not start with numbers"""
        prefix = re.split(r'[_ ]', project_name, 1)[0]
        return prefix if re.match(r'^\d[\d-]*$', prefix) else None

    def load_projects(self):
        """Load project numbers from the directory."""
        if not os.path.exists(PROJECTS_DIR):
            messagebox.showerror("Error", "Project Directory not found!")
            return

        valid_projects = []
        for d in os.listdir(PROJECTS_DIR):
            if os.path.isdir(os.path.join(PROJECTS_DIR, d)):
                prefix = self.extract_project_num(d)
                if prefix:
                    valid_projects.append((d, prefix))

        # sort project names based off of the project number (desc)
        valid_projects.sort(key=lambda x: x[1], reverse=True)

        if valid_projects:
            sorted_project_names = [p[0] for p in valid_projects]
            self.project_dropdown["values"] = sorted_project_names
            self.project_name.set(sorted_project_names[0])
            self.project_num = (valid_projects[0][1])

    def get_selected_project(self):
        """Returns the selected project path or None if none selected."""
        project = self.project_name.get()
        if not project:
            messagebox.showerror("Error", "No project selected!")
            return None
        
        self.project_num = self.extract_project_num(project)
        if not self.project_num:
            messagebox.showerror("Error", "Invalid project number structure!")

        return os.path.join(PROJECTS_DIR, project)

    def generate_nwd(self):
        """Generate an NWD file from the NWF file using PowerShell."""
        project_path = self.get_selected_project()
        if not project_path:
            return

        nwf_file = os.path.join(project_path, "CAD", "Piping", "Models", "_DesignReview", f"{self.project_num}-OverallModel{NWF_EXT}")
        nwd_file = os.path.join(project_path, "CAD", "Piping", "Models", "_DesignReview", f"{self.project_num}-DO_NOT_OPEN{NWD_EXT}")

        if not os.path.exists(nwf_file):
            messagebox.showerror("Error", f"NWF file not found for project: {project_path}")
            return

        try:
            command = ["powershell", "-ExecutionPolicy", "Bypass", "-File", CONVERT_PS_SCRIPT, nwf_file, nwd_file]
            subprocess.run(command, check=True, shell=True)
            messagebox.showinfo("Success", f"Generating NWD: {os.path.basename(nwd_file)}...")
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Error", f"Failed to generate NWD: {e}")

    def open_file(self, source_path, dest_path):
        """Open a file using PowerShell."""
        if not os.path.exists(source_path):
            messagebox.showerror("Error", f"File not found when opening: {source_path}")
            return

        try:
            command = ["powershell", "-ExecutionPolicy", "Bypass", "-File", OPEN_PS_SCRIPT, source_path, dest_path]
            subprocess.run(command, check=True, shell=True)
            messagebox.showinfo("Success", f"Opened {os.path.basename(source_path)} locally as {dest_path}")
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Error", f"Failed to open {source_path}: {e}")

    def open_nwd(self):
        """Open the NWD file."""
        project_path = self.get_selected_project()
        if not project_path:
            return

        nwd_file = os.path.join(project_path, "CAD", "Piping", "Models", "_DesignReview", f"{self.project_num}-DO_NOT_OPEN{NWD_EXT}")
        dest_file = os.path.join(NAVISWORKS_TEMP_DIR, f"{self.project_num}-OverallModel{NWD_EXT}")
        self.open_file(nwd_file, dest_file)

    def open_nwf(self):
        """Open the NWF file."""
        project_path = self.get_selected_project()
        if not project_path:
            return

        nwf_file = os.path.join(project_path, "CAD", "Piping", "Models", "_DesignReview", f"{self.project_num}-OverallModel{NWF_EXT}")
        dest_file = os.path.join(NAVISWORKS_TEMP_DIR, f"{self.project_num}-OverallModel{NWF_EXT}")
        self.open_file(nwf_file, dest_file)

if __name__ == "__main__":
    root = tk.Tk()
    app = NWGUI(root)
    root.mainloop()