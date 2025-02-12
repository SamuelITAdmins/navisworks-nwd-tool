import os
import re
import sys
import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import threading

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
        ttk.Label(root, text="Select Project:").grid(row=0, column=0, padx=0, pady=2)

        # Dropdown for Projects
        self.project_num = ''
        self.project_name = tk.StringVar()
        self.project_dropdown = ttk.Combobox(root, textvariable=self.project_name, width=45, state="readonly")
        self.project_dropdown.grid(row=1, column=0, columnspan=3, padx=5, pady=0)

        # Buttons
        ttk.Button(root, text="Generate NWD", command=self.generate_nwd).grid(row=2, column=0, padx=10, pady=15)
        ttk.Button(root, text="Open NWD", command=self.open_nwd).grid(row=2, column=1, padx=10, pady=15)
        ttk.Button(root, text="Open NWF", command=self.open_nwf).grid(row=2, column=2, padx=10, pady=15)
        ttk.Button(root, text="Refresh List", command=self.load_projects).grid(row=3, column=0, padx=10, pady=5)

        # Loading message (hidden initially)
        self.loading_label = ttk.Label(root, text="", font=('Segoe UI', 12))
        self.loading_label.grid(row=4, column=0, columnspan=3, padx=5, pady=5)
        self.load_projects()
    
    def extract_project_num(self, project_name):
        """Extract the project number from the full project name, discard if the project name does not start with numbers"""
        prefix = re.split(r'[_ ]', project_name, 1)[0]
        return prefix if re.match(r'^\d[\d-]*$', prefix) else None

    def load_projects(self):
        """Load project numbers from the directory with a wait message."""
        self.disable_gui()
        self.loading_label.config(text="Loading projects, please wait...")  # Show loading message
        self.root.update_idletasks()  # Force update GUI

        def background_task():
            if not os.path.exists(PROJECTS_DIR):
                messagebox.showerror("Error", "Project Directory not found!")
                return

            valid_projects = []
            for d in os.listdir(PROJECTS_DIR):
                if os.path.isdir(os.path.join(PROJECTS_DIR, d)):
                    prefix = self.extract_project_num(d)
                    if prefix:
                        valid_projects.append((d, prefix))

            # Sort project names based on project number (desc)
            valid_projects.sort(key=lambda x: x[1], reverse=True)

            if valid_projects:
                sorted_project_names = [p[0] for p in valid_projects]
                self.project_dropdown["values"] = sorted_project_names
                self.project_name.set(sorted_project_names[0])
                self.project_num = (valid_projects[0][1])

            self.loading_label.config(text="")  # Hide loading message after projects are loaded
            self.enable_gui()

        # Run in background thread
        threading.Thread(target=background_task, daemon=True).start()

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
            # Disable the entire GUI until powershell script executes
            self.disable_gui()
            self.loading_label.config(text="Generating NWD, please wait...")
            self.root.update_idletasks()

            # Run the powershell conversion script
            command = ["powershell", "-ExecutionPolicy", "Bypass", "-File", CONVERT_PS_SCRIPT, nwf_file, nwd_file]
            result = subprocess.run(command, capture_output=True, text=True, shell=True)

            # Re-enable the entire GUI once the script comes to a result
            self.loading_label.config(text="")
            self.enable_gui()

            if result.returncode != 0:
                # If the script errored...
                err_msg = result.stdout.strip().split('\n')[-1]
                messagebox.showerror("Error", f"NWD Conversion Error: {err_msg}")
                print(result.stdout.strip())
            else:
                # If conversion was successful
                messagebox.showinfo("Success", f"Generated NWD for {self.project_num}")
                print(result.stdout.strip())
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred, report the bug and close the program")
            return

    def open_file(self, source_path, dest_path):
        """Open a file using PowerShell."""
        if not os.path.exists(source_path):
            messagebox.showerror("Error", f"File not found when opening: {source_path}")
            return

        try:
            command = ["powershell", "-ExecutionPolicy", "Bypass", "-File", OPEN_PS_SCRIPT, source_path, dest_path]
            subprocess.run(command, check=True, shell=True)
            # messagebox.showinfo("Success", f"Opened {os.path.basename(source_path)} locally as {dest_path}")
        except subprocess.CalledProcessError as e:
            err_msg = e.stdout.strip().split('\n')[-1]
            messagebox.showerror("Error", f"Failed to open {source_path} locally as {dest_path}: {err_msg}")

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
    
    def disable_gui(self):
        """Disable all widgets in the root window."""
        for widget in self.root.winfo_children():
            widget.config(state="disabled")

    def enable_gui(self):
        """Enable all widgets in the root window."""
        for widget in self.root.winfo_children():
            widget.config(state="normal")

if __name__ == "__main__":
    root = tk.Tk()
    app = NWGUI(root)
    root.mainloop()