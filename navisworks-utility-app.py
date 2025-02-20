import re
import sys
import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import threading
import queue
from pathlib import Path
import time

# nwf file example: "\\stor-dn-01\projects\Projects\24317_Electra_CO_EPCM\CAD\Piping\Models\_DesignReview\24317-OverallModel.nwf"
# nwd file example: "\\stor-dn-01\projects\Projects\24317_Electra_CO_EPCM\CAD\Piping\Models\_DesignReview\24317-DO_NOT_OPEN.nwd"

# Paths (Modify as needed)
PROJECTS_DIR = Path(r"S:\Projects")  # Directory containing all projects
NAVISWORKS_TEMP_DIR = Path(r"C:\Navisworks") # Directory containing copied Navisworks files
NW_FILES_PATH = Path("CAD") / "Piping" / "Models" / "_DesignReview"
NWF_EXT = ".nwf"
NWD_EXT = ".nwd"

def get_resource_path(relative_path):
    '''Get absolute path to a resource, works for development and PyInstaller builds'''
    if getattr(sys, '_MEIPASS', False): # If running from .exe file
        return Path(sys._MEIPASS) / relative_path
    return Path(relative_path).resolve() # When running locally (python cmd)

CONVERT_PS_SCRIPT = get_resource_path("scripts/convertNWFfile.ps1")
OPEN_PS_SCRIPT = get_resource_path("scripts/openNWfile.ps1")

# TODO: Implement progress bar maps if additional operations are added
# TODO: Maintain the threading operation with a conversion map during temp file conversion
# Progress Bar Maps for Operations
NWD_CONVERSION_MAP = {
    "Beginning conversion": 10,
    "Opening NWF": 25,
    "Converting": 50,
    "Success": 100,
}

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
        self.generate_button = ttk.Button(root, text="Generate NWD", command=lambda: threading.Thread(target=self.generate_nwd, daemon=True).start())
        self.generate_button.grid(row=2, column=0, padx=10, pady=15)
        ttk.Button(root, text="Open NWD", command=self.open_nwd).grid(row=2, column=1, padx=10, pady=15)
        ttk.Button(root, text="Open NWF", command=self.open_nwf).grid(row=2, column=2, padx=10, pady=15)
        ttk.Button(root, text="Refresh List", command=self.load_projects).grid(row=3, column=0, padx=10, pady=5)

        # Loading message (hidden initially)
        self.loading_label = ttk.Label(root, text="", font=('Segoe UI', 12))
        self.loading_label.grid(row=4, column=0, columnspan=3, padx=5, pady=5)
        self.load_projects()

        # NWD conversion progress bar (hidden initially)
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(root, length=300, mode="determinate", variable=self.progress_var)
        self.progress_bar.grid(row=5, column=0, columnspan=3, padx=5, pady=5)
        self.progress_bar.grid_remove() # hide the progrss bar

        # Queue for powershell script communication
        self.script_queue = queue.Queue()

    def load_projects(self):
        """Load project numbers from the directory with a wait message."""
        self.disable_gui()
        self.loading_label.config(text="Loading projects, please wait...")  # Show loading message

        def fetch_projects():
            if not PROJECTS_DIR.exists():
                messagebox.showerror("Error", "Project Directory not found!")
                return

            valid_projects = [(p.name, self.extract_project_num(p.name)) for p in PROJECTS_DIR.iterdir() if p.is_dir() and self.extract_project_num(p.name)]
            valid_projects.sort(key=lambda x: x[1], reverse=True)

            if valid_projects:
                sorted_project_names = [p[0] for p in valid_projects]
                self.project_dropdown["values"] = sorted_project_names
                self.project_name.set(sorted_project_names[0])
                self.project_num = (valid_projects[0][1])

            self.loading_label.config(text="")  # Hide loading message after projects are loaded
            self.enable_gui()

        # Run in background thread to not make the GUI unresponsive
        threading.Thread(target=fetch_projects, daemon=True).start()

    def get_selected_project(self):
        """Returns the selected project path or None if none selected."""
        project = self.project_name.get()
        if not project:
            messagebox.showerror("Error", "No project selected!")
            return None
        
        self.project_num = self.extract_project_num(project)
        if not self.project_num:
            messagebox.showerror("Error", "Invalid project number structure!")

        return PROJECTS_DIR / project
    
    def extract_project_num(self, project_name):
        """Extract the project number from the full project name, discard if the project name does not start with numbers"""
        prefix = re.split(r'[_ ]', project_name, 1)[0]
        return prefix if re.match(r'^\d[\d-]*$', prefix) else None

    def generate_nwd(self):
        """Generate an NWD file from the NWF file using PowerShell."""
        project_path = self.get_selected_project()
        if not project_path:
            return

        nwf_file = project_path / NW_FILES_PATH / f"{self.project_num}-OverallModel{NWF_EXT}"
        nwd_file = project_path / NW_FILES_PATH / f"{self.project_num}-DO_NOT_OPEN{NWD_EXT}"
        
        if not CONVERT_PS_SCRIPT.exists():
            messagebox.showerror("Error", "Conversion PowerShell script not found!")
            return
        
        # Disable the entire GUI until powershell script executes
        self.disable_gui()
        self.loading_label.config(text="Generating NWD, please wait...")
        self.progress_var.set(0)
        self.progress_bar.grid()

        # Run the powershell conversion script
        command = ["powershell", "-ExecutionPolicy", "Bypass", "-File", str(CONVERT_PS_SCRIPT), str(nwf_file), str(nwd_file)]
        process = subprocess.Popen(
            command, 
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True, 
            # shell=True
        )
        
        def process_output():
            """Reads PowerShell script output and updates UI accordingly."""
            error_message = None
            try:
                for line in iter(process.stdout.readline, ''):
                    message = line.strip()
                    print(message) # message stream for debugging
                    if message:
                        self.script_queue.put(message)
                        self.root.after(0, lambda: self.loading_label.config(text=message))
                process.wait()

                if process.returncode != 0:
                    error_message = message or "Unknown error occured."
                process.stdout.close()
            except Exception as e:
                error_message = f"Unexpected error: {str(e)}"
            
            # Schedule GUI updates back on the main thread
            self.root.after(0, finalize_output, error_message)
        
        def finalize_output(error_message):
            """Handles GUI updates after script execution."""
            self.enable_gui()
            self.loading_label.config(text="")
            self.progress_bar.grid_remove()

            if error_message:
                messagebox.showerror("Error", f"NWD Conversion Error: {error_message}")
            else:
                messagebox.showinfo("Success", f"Generated NWD for {self.project_num}")
        
        # Start output processing in a separate thread
        threading.Thread(target=process_output, daemon=True).start()
        self.root.after(100, self.update_progress)

    def update_progress(self):
        """Reads messages from PowerShell and updates the progress bar accordingly"""
        while not self.script_queue.empty():
            message = self.script_queue.get_nowait()

            for key in NWD_CONVERSION_MAP:
                if key in message:
                    progress = NWD_CONVERSION_MAP.get(key, self.progress_var.get())
                    self.progress_var.set(progress)
                    break # stop checking when a key match is found
            
            # trigger track_file_size during NWD conversion
            if "Converting" in message:
                threading.Thread(target=self.track_file_size, daemon=True).start()

        if self.progress_var.get() < 100:
            self.root.after(100, self.update_progress)  # Recursion until completed
    
    def track_file_size(self):
        """Monitors the NWD~ (temp file) size and updates progress dynammically"""
        project_path = self.get_selected_project()
        if not project_path:
            return
        
        final_file = project_path / NW_FILES_PATH / f"{self.project_num}-DO_NOT_OPEN{NWD_EXT}"
        temp_file = project_path / NW_FILES_PATH / f"{self.project_num}-DO_NOT_OPEN{NWD_EXT}~"

        if not temp_file.exists() or not final_file.exists():
            return

        final_size = final_file.stat().st_size
        while temp_file.exists():
            temp_size = temp_file.stat().st_size
            if temp_size >= final_size:
                break

            # scale dynamically from 50 -> 100 based on temp file size
            new_progress = 50 + (temp_size / final_size) * 50
            self.progress_var.set(max(self.progress_var.get(), new_progress))  # Avoid rollback
            time.sleep(0.5)
        
        self.progress_var.set(100)

    def open_file(self, source_path, dest_path):
        """Open a file using PowerShell."""
        if not source_path.exists():
            messagebox.showerror("Error", f"File not found when opening: {source_path}")
            return

        try:
            if not OPEN_PS_SCRIPT.exists():
                messagebox.showerror("Error", "Open PowerShell script not found!")
                return

            command = ["powershell", "-ExecutionPolicy", "Bypass", "-File", str(OPEN_PS_SCRIPT), str(source_path), str(dest_path)]
            subprocess.run(command, check=True, shell=True)
            # messagebox.showinfo("Success", f"Opened {source_path.name} locally as {dest_path}")
        except subprocess.CalledProcessError as e:
            err_msg = e.stdout.strip().split('\n')[-1]
            messagebox.showerror("Error", f"Failed to open {source_path} locally as {dest_path}: {err_msg}")

    def open_nwd(self):
        """Open the NWD file."""
        project_path = self.get_selected_project()
        if not project_path:
            return

        nwd_file = project_path / NW_FILES_PATH / f"{self.project_num}-DO_NOT_OPEN{NWD_EXT}"
        dest_file = NAVISWORKS_TEMP_DIR / f"{self.project_num}-OverallModel{NWD_EXT}"
        self.open_file(nwd_file, dest_file)

    def open_nwf(self):
        """Open the NWF file."""
        project_path = self.get_selected_project()
        if not project_path:
            return

        nwf_file = project_path / NW_FILES_PATH / f"{self.project_num}-OverallModel{NWF_EXT}"
        dest_file = NAVISWORKS_TEMP_DIR / f"{self.project_num}-OverallModel{NWF_EXT}"
        self.open_file(nwf_file, dest_file)
    
    def disable_gui(self):
        """Disable all widgets in the root window."""
        for widget in self.root.winfo_children():
            try:
                widget.config(state="disabled")
            except tk.TclError:
                pass

    def enable_gui(self):
        """Enable all widgets in the root window."""
        for widget in self.root.winfo_children():
            try:
                widget.config(state="normal")
            except tk.TclError:
                pass

if __name__ == "__main__":
    root = tk.Tk()
    app = NWGUI(root)
    root.mainloop()