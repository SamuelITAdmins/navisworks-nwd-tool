import os
import subprocess

def run_vbs(script_path):
    """Executes a VBScript using Windows cscript"""
    if not os.path.exists(script_path):
        print(f"Error: {script_path} not found!")
        return

    try:
        subprocess.run(["cscript", script_path], check=True)
        print(f"Successfully executed: {script_path}")
    except subprocess.CalledProcessError as e:
        print(f"Execution failed: {e}")

# Example Usage for Project 24317
project_id = "24317"
base_vbs_path = "S:\\CAD_Standards\\_SE_Custom\\Piping\\Programming\\SE Piping Utilities\\Navisworks"

# Run the NWF loader script
run_vbs(f"{base_vbs_path}\\{project_id}-OverallModel.nwf.vbs")

# Run the NWD save script
run_vbs(f"{base_vbs_path}\\{project_id}-OverallModel.nwd.vbs")
