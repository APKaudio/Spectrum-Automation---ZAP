# main_app.py

import sys
import subprocess
import os
import tkinter as tk
from tkinter import messagebox

# Add the directory containing this script to sys.path
# This allows Python to find the 'utils' package when running main_app.py directly.
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

# Import the main application class and TextRedirector from user_interface
# Now that the script_dir is in sys.path, 'utils' can be imported as a package.
from utils.user_interface import App, TextRedirector

def check_and_install_dependencies():
    """
    Checks for required Python packages and installs them if missing.
    Returns True if all dependencies are met or successfully installed, False otherwise.
    """
    required_packages = [
        "pyvisa",
        "numpy",
        "pandas",
        "plotly"
    ]

    missing_packages = []
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            missing_packages.append(package)

    if missing_packages:
        print(f"Detected missing packages: {', '.join(missing_packages)}")
        response = messagebox.askyesno(
            "Missing Dependencies",
            f"The following Python packages are required but not found:\n{', '.join(missing_packages)}\n\n"
            "Do you want to install them now? This requires an internet connection."
        )
        if response:
            try:
                print("Attempting to install missing packages...")
                pip_path = [sys.executable, "-m", "pip"]
                subprocess.check_call(pip_path + ["install"] + missing_packages)
                print("Successfully installed missing packages.")
                return True
            except subprocess.CalledProcessError as e:
                messagebox.showerror(
                    "Installation Error",
                    f"Failed to install packages. Please install them manually using pip.\\nError: {e}"
                )
                print(f"❌ Failed to install packages: {e}")
                return False
            except Exception as e:
                messagebox.showerror(
                    "Installation Error",
                    f"An unexpected error occurred during installation: {e}"
                )
                print(f"❌ An unexpected error occurred during installation: {e}")
                return False
        else:
            print("User declined to install missing dependencies.")
            messagebox.showwarning(
                "Dependencies Not Met",
                "Critical dependencies are missing. The application may not function correctly."
            )
            return False
    return True

# The actual entry point of the script
if __name__ == '__main__':
    # Ensure dependencies are installed before running the app
    if check_and_install_dependencies():
        app = App()
        app.mainloop()
    else:
        print("Critical dependencies missing. Please install them to run the application.")
        messagebox.showerror("Dependency Error", "Critical dependencies missing. Please install them to run the application.")
