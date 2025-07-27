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
            3. Configures a tag to control line spacing.
        Outputs: None
        """
        self.widget = widget
        self.tag = tag
        self.last_char_was_cr = False

        # Configure a tag to control line spacing
        # spacing1: extra space above a line
        # spacing3: extra space below a line
        # We set them to 0 to try and minimize perceived double-spacing.
        # This explicitly tells the Text widget to not add extra space between lines.
        self.widget.tag_configure(self.tag, spacing1=0, spacing3=0)


    def write(self, str_val):
        """
        Writes the given string value to the Tkinter scrolled text widget.
        It ensures that each logical 'print' statement results in exactly one newline
        in the console and applies the configured tag for spacing.

        Inputs:
            str_val (str): The string to write to the console.
        Process:
            1. Strips any trailing newline characters from the input string.
            2. Inserts the processed string value at the end of the widget, applying the tag.
            3. Appends exactly one newline character.
            4. Scrolls to the end of the text widget to show the latest output.
            5. Updates Tkinter's idle tasks to ensure immediate display.
        Outputs: None
        """
        # Strip any trailing newlines from the input string to avoid double spacing
        # if the original print statement already added one.
        str_val = str_val.rstrip('\n')
        self.widget.insert(tk.END, str_val, self.tag) # Apply the tag here
        self.widget.insert(tk.END, '\n') # Always add exactly one newline
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
        2. Each `print()` call is now configured to not add its own newline,
           relying solely on the `TextRedirector` to manage line breaks.
        3. Explicit blank lines in the ASCII art are now `print("", end='')`
           to ensure they are truly empty lines without any spaces.
    Outputs: None (prints to console)
    """
    # By adding end='', we prevent print() from adding its own newline,
    # allowing TextRedirector to control the spacing.
    # Changed lines that were just spaces to empty strings for true blank lines.
    print("", end='') # This was a line of spaces, now truly empty
    print("", end='') # This was a line of spaces, now truly empty
    print("", end='') # This was a line of spaces, now truly empty
    print("                                               $              $$$$$                     $$ $$$$", end='')
    print("                                               $$$            $$   $$$$$$               $$  $$ ", end='')
    print("                                               $$$$           $$         $$$$$          $$ $$  ", end='')
    print("                                  $$           $$ $$          $$             $$$$$      $$$$   ", end='')
    print("                       $$$$$$$$$$$$            $$  $$$        $                  $$$    $$$$   ", end='')
    print("             $$$$$$$$$        $$$              $$   $$$      $$                    $$   $$$    ", end='')
    print("   $$$$$$$$$               $$$                 $$     $$     $$                $$$$     $$     ", end='')
    print("                         $$$                   $$$$$$$$$$$   $$          $$$$$$         $$     ", end='')
    print("                       $$$                     $$       $$$  $  $$$$$$$$                $      ", end='')
    print("                     $$$                       $$         $$$$$                                ", end='')
    print("                   $$$                         $$           $$                        $ $$     ", end='')
    print("                 $$$                $$$$$$$                 $$                        $$$      ", end='')
    print("              $$$$            $$$$$$                        $$                  $$$$$$$$$$     ", end='')
    print("            $$$        $$$$$$$                                   $$$$$$$$$$$$$$                ", end='')
    print("          $$$   $$$$$$$                             $$$$$$$$$$$                                ", end='')
    print("        $$$$$$$$                  $$$$$$$$             $$$$                                    ", end='')
    print("      $$$              $$$$$$$$$$  $$$$              $$$$$$$$                                  ", end='')
    print("           $$$$$$$$$$$          $$$         $$$$$$$$                                           ", end='')
    print(" $$$$$$$$$$                  $$$    $$$$$$$$                                                   ", end='')
    print("                         $$$$$$$$$$                                                            ", end='')
    print("                      $$$$$                        ", end='')
    print("", end='') # Was "    ", now truly empty
    print("", end='') # Was "    ", now truly empty
    print("", end='') # Was "    ", now truly empty
    print("Software created for  https://zimbelaudio.com/ike-zimbel/    ", end='')
    print("A Colaboration betweeen Ike Zimbel and Anthony P. Kuzub", end='')
    print("", end='') # Was "    ", now truly empty
    print("", end='') # Was "    ", now truly empty
