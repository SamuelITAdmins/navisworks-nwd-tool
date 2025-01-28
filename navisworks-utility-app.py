import os
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import subprocess

# nwf file = "\\stor-dn-01\projects\Projects\24317_Electra_CO_EPCM\CAD\Piping\Models\_DesignReview\24317-OverallModel.nwf"
# nwd file = "\\stor-dn-01\projects\Projects\24317_Electra_CO_EPCM\CAD\Piping\Models\_DesignReview\24317-DO_NOT_OPEN.nwd"

# Paths (Modify as needed)
PROJECTS_DIR = r"\\stor-dn-01\projects\Projects\24317_Electra_CO_EPCM"  # Directory containing project
NWF_EXT = ".nwf"
NWD_EXT = ".nwd"
POWERSHELL_SCRIPT = "ConvertNWFtoNWD.ps1"
VBSCRIPT = "OpenFile.vbs"

class NWGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Navisworks Project Manager")

        # Dropdown Label
        ttk.Label(root, text="Select Project:").grid(row=0, column=0, padx=5, pady=5)

        # Dropdown for Projects
        self.project_var = tk.StringVar()
        self.project_dropdown = ttk.Combobox(root, textvariable=self.project_var, state="readonly")
        self.project_dropdown.grid(row=0, column=1, padx=5, pady=5)
        self.load_projects()

        # Buttons
        ttk.Button(root, text="Generate NWD", command=self.generate_nwd).grid(row=1, column=0, padx=5, pady=5)
        ttk.Button(root, text="Open NWD", command=self.open_nwd).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(root, text="Open NWF", command=self.open_nwf).grid(row=2, column=0, padx=5, pady=5)
        ttk.Button(root, text="Add Project", command=self.create_project).grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(root, text="Refresh", command=self.load_projects).grid(row=3, column=0, padx=5, pady=5)

    def load_projects(self):
        """Load project numbers from the directory."""
        if not os.path.exists(PROJECTS_DIR):
            os.makedirs(PROJECTS_DIR)

        projects = [d for d in os.listdir(PROJECTS_DIR) if os.path.isdir(os.path.join(PROJECTS_DIR, d))]
        self.project_dropdown["values"] = projects
        if projects:
            self.project_var.set(projects[0])

    def get_selected_project(self):
        """Returns the selected project path or None if none selected."""
        project = self.project_var.get()
        if not project:
            messagebox.showerror("Error", "App: No project selected.")
            return None
        return os.path.join(PROJECTS_DIR, project)

    def generate_nwd(self):
        """Generate an NWD file from the NWF file using PowerShell."""
        project_path = self.get_selected_project()
        if not project_path:
            return

        nwf_file = os.path.join(project_path, f"{os.path.basename(project_path)} + \CAD\Piping\Models\_DesignReview\24317-OverallModel + {NWF_EXT}")
        nwd_file = os.path.join(project_path, f"{os.path.basename(project_path)} + \CAD\Piping\Models\_DesignReview\24317-DO_NOT_OPEN + {NWD_EXT}")

        if not os.path.exists(nwf_file):
            messagebox.showerror("Error", f"App: NWF file not found: {nwf_file}")
            return

        try:
            subprocess.run(["powershell", "-File", POWERSHELL_SCRIPT, nwf_file, nwd_file], check=True)
            messagebox.showinfo("Success", f"App generated NWD: {nwd_file}")
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Error", f"App failed to generate NWD: {e}")

    def open_file(self, file_path, dest_path):
        """Open a file using VBScript."""
        if not os.path.exists(file_path):
            messagebox.showerror("Error", f"App: File not found: {file_path}")
            return

        try:
            subprocess.run(["wscript", VBSCRIPT, file_path, dest_path], check=True)
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Error", f"App: Failed to open file: {e}")

    def open_nwd(self):
        """Open the NWD file."""
        project_path = self.get_selected_project()
        if not project_path:
            return

        nwd_file = os.path.join(project_path, f"{os.path.basename(project_path)}{NWD_EXT}")
        self.open_file(nwd_file)

    def open_nwf(self):
        """Open the NWF file."""
        project_path = self.get_selected_project()
        if not project_path:
            return

        nwf_file = os.path.join(project_path, f"{os.path.basename(project_path)}{NWF_EXT}")
        self.open_file(nwf_file)

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

if __name__ == "__main__":
    root = tk.Tk()
    app = NWGUI(root)
    root.mainloop()
