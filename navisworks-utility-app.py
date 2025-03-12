import re
import os
import sys
import shutil
import json
import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import threading
import queue
from pathlib import Path
import time

# TODO: Change NW_FILES_PATH from a hardcoded path to a directory search algorithm that finds the NW files so that 
#       it translates to all project file structures
# TODO: Add a timer when the progress bar says "Opening NWF..." that displays how many minute(s) have elapsed
# TODO: Filter project list down and/or archive unused projects in the S drive
# TODO: Change the threading type during project list retrieval to optimize fetching

# nwf file example: "\\stor-dn-01\projects\Projects\24317_Electra_CO_EPCM\CAD\Piping\Models\_DesignReview\24317-OverallModel.nwf"
# nwd file example: "\\stor-dn-01\projects\Projects\24317_Electra_CO_EPCM\CAD\Piping\Models\_DesignReview\24317-DO_NOT_OPEN.nwd"

# Paths (Modify as needed)
PROJECTS_DIR = Path(r"S:\Projects")  # Directory containing all projects
NAVISWORKS_TEMP_DIR = Path(r"C:\Navisworks") # Directory containing copied Navisworks files
CACHE_FILE = Path(Path.home()) / "AppData" / "Local" / "Temp" / "projectsListCache.json"
CACHE_EXPIRATION = 7 * 24 * 60 * 60 # Cache expires in a week (in seconds)
NW_FILES_PATH = Path("CAD") / "Piping" / "Models" / "_DesignReview"
NWF_EXT = ".nwf"
NWD_EXT = ".nwd"

def get_resource_path(relative_path):
    '''Get absolute path to a resource, works for development and PyInstaller builds'''
    if getattr(sys, '_MEIPASS', False): # If running from .exe file
        return Path(sys._MEIPASS) / relative_path
    return Path(relative_path).resolve() # When running locally (python cmd)

GET_NW_PATH_SCRIPT = get_resource_path("scripts/getNWroamerfile.ps1")
CONVERT_PS_SCRIPT = get_resource_path("scripts/convertNWFfile.ps1")
OPEN_PS_SCRIPT = get_resource_path("scripts/openNWfile.ps1")

def cleanup_temp_folder():
    """Deletes the PyInstaller _MEIxxxxx folder on exit."""
    if getattr(sys, 'frozen', False):  # Only in PyInstaller executable
        temp_dir = sys._MEIPASS  # PyInstaller extracts files here
        try:
            shutil.rmtree(temp_dir)  # Remove the temp directory
        except PermissionError:
            print(f"Warning: Could not remove temp dir {temp_dir}")

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

        # Initialize Navisworks Roamer and check Permissions
        self.roamer_path = self.get_NW_path()
        self.editor = self.check_NW_permission(self.roamer_path)

        # Dropdown Label
        ttk.Label(root, text="Select Project:").grid(row=0, column=0, padx=0, pady=2)

        # Dropdown for Projects
        self.project_num = ''
        self.project_name = tk.StringVar()
        self.project_dropdown = ttk.Combobox(root, textvariable=self.project_name, width=45, state="readonly")
        self.project_dropdown.grid(row=1, column=0, columnspan=3, padx=5, pady=0)

        # Buttons
        self.generate_button = ttk.Button(root, text="Generate NWD", command=self.generate_nwd)
        self.open_nwf_button = ttk.Button(root, text="Open NWF", command=self.open_nwf)
        self.open_nwd_button = ttk.Button(root, text="Open NWD", command=self.open_nwd)
        self.refresh_button = ttk.Button(root, text="Refresh List", command=self.reload_projects)

        # Grid Placement (based on editor permissions)
        self.open_nwd_button.grid(row=2, column=0, padx=10, pady=15)
        if self.editor:
            self.generate_button.grid(row=2, column=1, padx=10, pady=15)
            self.open_nwf_button.grid(row=2, column=2, padx=10, pady=15)
        else:
            self.generate_button.grid_remove()
            self.open_nwf_button.grid_remove()
        self.refresh_button.grid(row=3, column=0, padx=10, pady=5)

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

    def get_NW_path(self):
        """Gets the Roamer path for any Navisworks version (Manage, Simulate, Freedom) from the C drive"""
        try:
            if not GET_NW_PATH_SCRIPT.exists():
                self.show_error("Get NW Roamer script not found!")
                return

            command = ["powershell", "-ExecutionPolicy", "Bypass", "-File", str(GET_NW_PATH_SCRIPT)]
            result = subprocess.run(command, capture_output=True, check=True)
            
            if result.returncode == 0:
                nw_path = Path(result.stdout.decode().strip().split('\n')[-1]) # gets the last return message
                if nw_path.exists():
                    return nw_path
                else:
                    self.show_error("Failed to retrieve Navisworks Roamer path.")
                    self.disable_gui()
        except subprocess.CalledProcessError as e:
            err_msg = e.stdout.strip().split('\n')[-1]
            self.show_error(err_msg)
            self.disable_gui()

    def check_NW_permission(self, roamer_path):
        """If the user has Navisworks Freedom, then return false for editing permissions, true otherwise."""
        return not "Freedom" in str(roamer_path)
    
    def reload_projects(self):
        """Delete the cache and load_projects upon refreshing the list."""
        if CACHE_FILE.exists():
            CACHE_FILE.unlink()

        self.load_projects()

    def load_projects(self):
        """Load project numbers from the directory with a wait message."""
        self.disable_gui()
        self.loading_label.config(text="Loading projects, please wait...")  # Show loading message

        def fetch_projects():
            if not PROJECTS_DIR.exists():
                self.show_error("Project Directory not found!")
                return
            
            # load projects from cache if it is still valid
            cached_projects = self.load_projects_from_cache()
            if cached_projects:
                self.root.after(0, lambda: self.update_projects(cached_projects))
                return

            # if cache is invalid, rebuild project list from project directory
            valid_projects = []
            with os.scandir(PROJECTS_DIR) as projects:
                for p in projects:
                    if p.is_dir():
                        project_num = self.extract_project_num(p.name)
                        if project_num:
                            valid_projects.append((p.name, project_num))

            valid_projects.sort(key=lambda x: x[1], reverse=True)

            # save to cache
            self.save_projects_to_cache(valid_projects)

            # update UI safely
            self.root.after(0, lambda: self.update_projects(valid_projects))

        # Run in background thread to prevent UI freezing
        threading.Thread(target=fetch_projects, daemon=True).start()

    def update_projects(self, valid_projects):
        """Safely update the UI with the loaded project list."""
        if valid_projects:
            sorted_project_names = [p[0] for p in valid_projects]
            self.project_dropdown["values"] = sorted_project_names
            self.project_name.set(sorted_project_names[0])
            self.project_num = valid_projects[0][1]

        # return GUI to normal
        self.loading_label.config(text="")
        self.enable_gui()

    def get_selected_project(self):
        """Returns the selected project path or None if none selected."""
        project = self.project_name.get()
        if not project:
            self.root.after(0, lambda: messagebox.showerror("Error", "No project selected!"))
            return None
        
        self.project_num = self.extract_project_num(project)
        if not self.project_num:
            self.show_error("Invalid project number structure!")

        return PROJECTS_DIR / project
    
    def extract_project_num(self, project_name):
        """Extract the project number from the full project name, discard if the project name does not start with numbers."""
        prefix = re.split(r'[_ ]', project_name, 1)[0]
        return prefix if re.match(r'^\d[\d-]*$', prefix) else None

    def load_projects_from_cache(self):
        """Load projects from cache if it's still valid."""
        if CACHE_FILE.exists():
            cache_age = time.time() - CACHE_FILE.stat().st_mtime
            if cache_age < CACHE_EXPIRATION:
                try:
                    with open(CACHE_FILE, "r") as f:
                        return json.load(f)
                except json.JSONDecodeError:
                    self.show_error("Cache file is corrupted. Regenerating project list.")
        return None
    
    def save_projects_to_cache(self, valid_projects):
        """Save project list to cache as JSON."""
        try:
            with open(CACHE_FILE, "w") as f:
                json.dump(valid_projects, f)
        except Exception as e:
            self.show_error(f"Failed to save cache: {e}")

    def generate_nwd(self):
        """Generate the selected project NWD file from the NWF file using PowerShell."""
        project_path = self.get_selected_project()
        if not project_path:
            return

        nwf_file = project_path / NW_FILES_PATH / f"{self.project_num}-OverallModel{NWF_EXT}"
        nwd_file = project_path / NW_FILES_PATH / f"{self.project_num}-DO_NOT_OPEN{NWD_EXT}"
        
        if not CONVERT_PS_SCRIPT.exists():
            self.show_error("Conversion PowerShell script not found!")
            return
        
        # Disable the entire GUI until powershell script executes
        self.loading_label.config(text="Generating NWD, please wait...")
        self.progress_var.set(0)
        self.progress_bar.grid()
        self.disable_gui()

        # Run the powershell conversion script
        command = ["powershell", "-ExecutionPolicy", "Bypass", "-File", str(CONVERT_PS_SCRIPT), str(self.roamer_path), str(nwf_file), str(nwd_file)]
        process = subprocess.Popen(
            command, 
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        
        def process_output():
            """Reads PowerShell script output and updates UI accordingly."""
            error_message = None
            try:
                for line in iter(process.stdout.readline, ''):
                    message = line.strip()
                    # print(message) # message stream for debugging
                    if message:
                        self.script_queue.put(message)
                        self.root.after(0, lambda: self.loading_label.config(text=message))
                        self.root.after(0, lambda msg=message: self.root.after(0, self.update_progress(message)))
                process.wait()

                # if returncode = 1 then error, otherwise returncode = 0 and successful:
                if process.returncode != 0:
                    error_message = message or "Unknown error occured."
                else:
                    final_message = message or f"Generated NWD for {self.project_num}"
                process.stdout.close()
            except Exception as e:
                error_message = f"Unexpected error: {str(e)}"
            
            # Schedule GUI updates back on the main thread
            self.root.after(0, finalize_output, final_message, error_message)
        
        def finalize_output(final_message, error_message):
            """Handles GUI updates after script execution."""
            self.enable_gui()
            self.loading_label.config(text="")
            self.progress_var.set(0)
            self.progress_bar.grid_remove()

            if error_message:
                self.root.after(0, lambda: messagebox.showerror("Error",  f"NWD Conversion Error: {error_message}"))
            else:
                self.root.after(0, lambda: messagebox.showinfo("Success", f"NWD for {self.project_num} {final_message}"))
        
        # Start output processing in a separate thread
        threading.Thread(target=process_output, daemon=True).start()

    def update_progress(self, message):
        """Updates the progress bar according to the inputted message"""
        for key in NWD_CONVERSION_MAP:
            if key in message:
                progress = NWD_CONVERSION_MAP.get(key, self.progress_var.get())
                self.progress_var.set(progress)
                break # stop checking when a key match is found
        
        # trigger track_file_size during NWD conversion
        if "Converting" in message:
            threading.Thread(target=self.track_file_size, daemon=True).start()
    
    def track_file_size(self):
        """Monitors the NWD~ (temp file) size and updates progress dynammically"""
        project_path = self.get_selected_project()
        if not project_path:
            return
        
        final_file = project_path / NW_FILES_PATH / f"{self.project_num}-DO_NOT_OPEN{NWD_EXT}"
        temp_file = project_path / NW_FILES_PATH / f"{self.project_num}-DO_NOT_OPEN{NWD_EXT}~"

        # if this is the first time the nwd is generated, then the temp file will not exist, thus exit
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
            self.root.after(0, lambda: messagebox.showerror("Error", f"File to open not found: {source_path}\nEither generate the file, or the project ({self.project_name.get()}) may not be a Navisworks project."))
            return

        try:
            if not OPEN_PS_SCRIPT.exists():
                self.show_error("Open PowerShell script not found!")
                return

            command = ["powershell", "-ExecutionPolicy", "Bypass", "-File", str(OPEN_PS_SCRIPT), str(self.roamer_path), str(source_path), str(dest_path)]
            subprocess.run(command, check=True)
        except subprocess.CalledProcessError as e:
            err_msg = e.stdout.strip().split('\n')[-1]
            self.show_error(f"Failed to open {source_path} locally as {dest_path}: {err_msg}")

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

    def show_error(self, message):
        """Safely display an error message and prompt the user to alert the IT team."""
        self.root.after(0, lambda: messagebox.showerror("Error", message + " Report this error to IT."))

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = NWGUI(root)
        root.mainloop()
    finally:
        cleanup_temp_folder()  # Clean up _MEIxxxxx folder