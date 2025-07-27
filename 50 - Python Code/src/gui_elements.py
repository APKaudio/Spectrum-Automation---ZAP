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
        # The last_char_was_cr flag is no longer strictly needed for the simplified write,
        # but keeping it for now if future complex behavior is reintroduced.
        self.last_char_was_cr = False

    def write(self, str_val):
        """
        Writes the given string value to the Tkinter scrolled text widget.
        This simplified version will always append the string and then a newline,
        effectively making every "print" statement appear on a new line.

        Inputs:
            str_val (str): The string to write to the console.
        Process:
            1. Inserts the string value at the end of the widget.
            2. Appends a newline character.
            3. Scrolls to the end of the text widget to show the latest output.
        Outputs: None
        """
        self.widget.insert(tk.END, str_val)
        # Ensure a newline is always added if the string doesn't already end with one,
        # to prevent "run-on" lines from different print statements.
        if not str_val.endswith('\n'):
            self.widget.insert(tk.END, '\n')
        self.widget.see(tk.END) # Always scroll to the end
        self.widget.update_idletasks() # Ensure the display updates immediately


    def flush(self):
        """
        Required for file-like objects. Ensures that output is processed.
        """
        pass # Tkinter widget updates are handled by .see(tk.END) and .update_idletasks()

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