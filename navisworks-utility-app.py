import os
import re
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import subprocess

# nwf file = "\\stor-dn-01\projects\Projects\24317_Electra_CO_EPCM\CAD\Piping\Models\_DesignReview\24317-OverallModel.nwf"
# nwd file = "\\stor-dn-01\projects\Projects\24317_Electra_CO_EPCM\CAD\Piping\Models\_DesignReview\24317-DO_NOT_OPEN.nwd"

# Paths (Modify as needed)
PROJECTS_DIR = r"S:\Projects"  # Directory containing all projects
NAVISWORKS_TEMP_DIR = r"C:\Navisworks" # Directory containing copied Navisworks files
NWF_EXT = ".nwf"
NWD_EXT = ".nwd"
CONVERT_PS_SCRIPT = "C:\\Users\\sapozder\\Sandbox\\navisworks-nwd-tool\\ConvertNWFtoNWD.ps1"
OPEN_PS_SCRIPT = "C:\\Users\\sapozder\\Sandbox\\navisworks-nwd-tool\\OpenNWFile.ps1"

class NWGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Navisworks Project Manager")

        # Dropdown Label
        ttk.Label(root, text="Select Project:").grid(row=0, column=0, padx=5, pady=5)

        # Dropdown for Projects
        self.project_num = ''
        self.project_name = tk.StringVar()
        self.project_dropdown = ttk.Combobox(root, textvariable=self.project_name, state="readonly")
        self.project_dropdown.grid(row=0, column=1, padx=5, pady=5)
        self.load_projects()

        # Buttons
        ttk.Button(root, text="Generate NWD", command=self.generate_nwd).grid(row=1, column=0, padx=5, pady=5)
        ttk.Button(root, text="Open NWD", command=self.open_nwd).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(root, text="Open NWF", command=self.open_nwf).grid(row=2, column=0, padx=5, pady=5)
        # ttk.Button(root, text="Add Project", command=self.create_project).grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(root, text="Refresh", command=self.load_projects).grid(row=3, column=0, padx=5, pady=5)
    
    def extract_project_num(project_name):
        """Extract the project number from the full project name, discard if the project name does not start with numbers"""
        prefix = re.split(r'[_ ]', project_name, 1)[0]
        return prefix if re.match(r'^\d[\d-]*$', prefix) else None

    def load_projects(self):
        """Load project numbers from the directory."""
        if not os.path.exists(PROJECTS_DIR):
            messagebox.showerror("Error", "App: Project Directory not found.")
            return

        valid_projects = [
            (d, prefix) for d in os.listdir(PROJECTS_DIR) 
            if os.path.isdir(os.path.join(PROJECTS_DIR, d)) and (prefix := self.extract_project_num(d))
        ]
        # sort project names based off of the project number
        valid_projects.sort(key=lambda x: x[1])

        if valid_projects:
            sorted_project_names = [p[0] for p in valid_projects]
            self.project_dropdown["values"] = sorted_project_names
            self.project_name.set(sorted_project_names[0])
            self.project_num = ([p[1] for p in valid_projects][0])

    def get_selected_project(self):
        """Returns the selected project path or None if none selected."""
        project = self.project_name.get()
        if not project:
            messagebox.showerror("Error", "App: No project selected.")
            return None
        self.project_num = (self.extract_project_num(project))
        return os.path.join(PROJECTS_DIR, project)

    def generate_nwd(self):
        """Generate an NWD file from the NWF file using PowerShell."""
        project_path = self.get_selected_project()
        if not project_path:
            return

        nwf_file = f"{project_path}\CAD\Piping\Models\_DesignReview\{self.project_num}-OverallModel{NWF_EXT}"
        nwd_file = f"{project_path}\CAD\Piping\Models\_DesignReview\{self.project_num}-DO_NOT_OPEN{NWD_EXT}"

        if not os.path.exists(nwf_file):
            messagebox.showerror("Error", f"App: NWF file not found: {nwf_file}")
            return

        try:
            command = ["powershell", "-ExecutionPolicy", "Bypass", "-File", CONVERT_PS_SCRIPT, nwf_file, nwd_file]
            subprocess.run(command, check=True)
            messagebox.showinfo("Success", f"App generated NWD: {nwd_file}")
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Error", f"App failed to generate NWD: {e}")

    def open_file(self, source_path, dest_path):
        """Open a file using PowerShell."""
        if not os.path.exists(source_path):
            messagebox.showerror("Error", f"App: File not found: {source_path}")
            return

        try:
            command = ["powershell", "-ExecutionPolicy", "Bypass", "-File", OPEN_PS_SCRIPT, source_path, dest_path]
            subprocess.run(command, check=True)
            messagebox.showinfo("Success", f"App copied from {source_path} to {dest_path}")
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Error", f"App failed to open {source_path}: {e}")

    def open_nwd(self):
        """Open the NWD file."""
        project_path = self.get_selected_project()
        if not project_path:
            return

        nwd_file = f"{project_path}\CAD\Piping\Models\_DesignReview\{self.project_num}-DO_NOT_OPEN{NWD_EXT}"
        dest_file = f"C:\\Navisworks\\{self.project_num}-OverallModel{NWD_EXT}"
        self.open_file(nwd_file, dest_file)

    def open_nwf(self):
        """Open the NWF file."""
        project_path = self.get_selected_project()
        if not project_path:
            return

        nwf_file = f"{project_path}\CAD\Piping\Models\_DesignReview\{self.project_num}-OverallModel{NWF_EXT}"
        dest_file = f"C:\\Navisworks\\{self.project_num}-OverallModel{NWF_EXT}"
        self.open_file(nwf_file, dest_file)

    '''
    def create_project(self):
        """Create a new project folder."""
        new_project = simpledialog.askstring("New Project", "Enter project number:")
        if not new_project:
            return

        new_project_path = os.path.join(PROJECTS_DIR, new_project)

        if os.path.exists(new_project_path):
            messagebox.showerror("Error", "Project already exists.")
            return

        os.makedirs(new_project_path)

        # Creating an empty NWF file
        open(os.path.join(new_project_path, f"{new_project}{NWF_EXT}"), "w").close()

        messagebox.showinfo("Success", f"Project created: {new_project}")
        self.load_projects()
    '''

if __name__ == "__main__":
    root = tk.Tk()
    app = NWGUI(root)
    root.mainloop()
