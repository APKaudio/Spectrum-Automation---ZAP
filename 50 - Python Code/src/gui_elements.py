# src/gui_elements.py
import tkinter as tk
import sys
from tkinter import scrolledtext, TclError

class TextRedirector(object):
    """
    A class to redirect standard output (stdout) and standard error (stderr)
    to a Tkinter scrolled text widget. This allows all print statements and
    error messages from the application's backend to be displayed directly
    within the GUI's console area, providing real-time feedback to the user.
    """
    def __init__(self, widget, tag="stdout"):
        """
        Initializes the TextRedirector.

        Inputs:
            widget (tk.scrolledtext.ScrolledText): The Tkinter scrolled text widget
                                                  where output will be displayed.
            tag (str, optional): A tag for text formatting within the widget. Defaults to "stdout".
        Process:
            1. Stores the provided `widget` and `tag`.
            2. Initializes `last_char_was_cr` to False, used for handling carriage returns for line overwriting.
        Outputs: None
        """
        self.widget = widget
        self.tag = tag
        self.last_char_was_cr = False

    def write(self, str_val):
        """
        Writes the given string value to the Tkinter scrolled text widget.
        Handles carriage returns (`\r`) to overwrite the current line,
        useful for progress bars or dynamic console updates.

        Inputs:
            str_val (str): The string to write to the console.
        Process:
            1. Sets the widget state to `tk.NORMAL` to allow editing.
            2. Checks if `str_val` contains `\r`.
            3. If `\r` is present, splits the string and handles line deletion
               for overwriting if the previous character was also a carriage return.
            4. Inserts the string (or parts of it) into the widget at the end.
            5. Scrolls the widget to the end to show the latest output.
            6. Sets the widget state back to `tk.DISABLED` to prevent user editing.
            7. Updates Tkinter idle tasks to ensure immediate display.
        Outputs: None
        """
        self.widget.config(state=tk.NORMAL)
        
        if '\r' in str_val:
            parts = str_val.split('\r')
            for i, part in enumerate(parts):
                if self.last_char_was_cr and i == 0:
                    self.widget.delete("end-1c linestart", "end-1c")
                self.widget.insert(tk.END, part, (self.tag,))
                self.widget.see(tk.END)
                if i < len(parts) - 1:
                    self.last_char_was_cr = True
                else:
                    self.last_char_was_cr = False
        else:
            self.widget.insert(tk.END, str_val, (self.tag,))
            self.widget.see(tk.END)
            self.last_char_was_cr = False

        self.widget.config(state=tk.DISABLED)
        self.widget.update_idletasks()

    def flush(self):
        """
        Required method for file-like objects. Does nothing in this implementation.
        """
        pass

def print_art():
    """
    Prints an ASCII art logo to the console output. This function is called
    during application startup to provide a visual brand element.

    Inputs: None
    Process:
        1. Uses a series of `print()` statements to output the multi-line ASCII art.
    Outputs: None (prints to console)
    """
    print("                                                                                               ")
    print("                                               $              $$$$$                     $$ $$$$")
    print("                                               $$$            $$   $$$$$$               $$  $$ ")
    print("                                               $$$$           $$         $$$$$          $$ $$  ")
    print("                                  $$           $$ $$          $$             $$$$$      $$$$   ")
    print("                       $$$$$$$$$$$$            $$  $$$        $                  $$$    $$$$   ")
    print("             $$$$$$$$$        $$$              $$   $$$      $$                    $$   $$$    ")
    print("   $$$$$$$$$               $$$                 $$     $$     $$                $$$$     $$     ")
    print("                         $$$                   $$$$$$$$$$$   $$          $$$$$$         $$     ")
    print("                       $$$                     $$       $$$  $  $$$$$$$$                $      ")
    print("                     $$$                       $$         $$$$$                                ")
    print("                   $$$                         $$           $$                        $ $$     ")
    print("                 $$$                $$$$$$$                 $$                        $$$      ")
    print("              $$$$            $$$$$$                        $$                  $$$$$$$$$$     ")
    print("            $$$        $$$$$$$                                   $$$$$$$$$$$$$$                ")
    print("          $$$   $$$$$$$                             $$$$$$$$$$$                                ")
    print("        $$$$$$$$                  $$$$$$$$             $$$$                                    ")
    print("      $$$              $$$$$$$$$$  $$$$              $$$$$$$$                                  ")
    print("           $$$$$$$$$$$          $$$         $$$$$$$$                                           ")
    print(" $$$$$$$$$$                  $$$    $$$$$$$$                                                   ")
    print("                         $$$$$$$$$$                                                            ")
    print("                      $$$$$                        ")
    print("    ")
    print("    ")
    print("    ")
    print("Software created for  https://zimbelaudio.com/ike-zimbel/    ")
    print("A Colaboration betweeen Ike Zimbel and Anthony P. Kuzub")
    print("    ")
    print("    ")
