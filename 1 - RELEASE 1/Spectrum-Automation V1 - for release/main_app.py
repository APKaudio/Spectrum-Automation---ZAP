# main_app.py
#
# This is the main entry point for the RF Spectrum Analyzer Controller application.
# It handles initial setup, checks for and installs necessary Python dependencies,
# and then launches the main graphical user interface (GUI).
# This file ensures that the application environment is ready before starting the UI.
#
# Author: Anthony Peter Kuzub
# Blog: www.Like.audio (Contributor to this project)
#
# Professional services for customizing and tailoring this software to your specific
# application can be negotiated. There is no change to use, modify, or fork this software.
#
# Build Log: https://like.audio/category/software/spectrum-scanner/
# Source Code: https://github.com/APKaudio/
#
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
    This function is crucial for ensuring the application has all necessary libraries
    (pyvisa, numpy, pandas, plotly) to function correctly. If dependencies are missing,
    it prompts the user to install them via pip.

    Inputs: None
    Process:
        1. Defines a list of `required_packages`.
        2. Iterates through the `required_packages`, attempting to `__import__` each.
        3. If an `ImportError` occurs, the package name is added to `missing_packages`.
        4. If `missing_packages` is not empty, it displays a `messagebox.askyesno` dialog
           to the user, asking if they want to install the missing packages.
        5. If the user agrees, it attempts to install them using `subprocess.check_call` with `pip`.
        6. Handles `subprocess.CalledProcessError` (pip installation failure) and other `Exception` types.
        7. If the user declines, it shows a `messagebox.showwarning` about missing critical dependencies.
    Outputs:
        bool: True if all dependencies are met or successfully installed; False otherwise.
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
    # Ensures that all necessary Python packages are installed before the application starts.
    # This prevents runtime errors due to missing dependencies and provides a smoother user experience.
    if check_and_install_dependencies():
        app = App()
        app.mainloop()
    else:
        print("Critical dependencies missing. Please install them to run the application.")
        messagebox.showerror("Dependency Error", "Critical dependencies missing. Please install them to run the application.")
