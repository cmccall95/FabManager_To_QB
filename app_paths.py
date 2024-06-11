# path_util.py
import sys
import os

def get_base_path():
    """Get the base directory path for the application."""
    if getattr(sys, 'frozen', False):
        # The application is running as a standalone executable
        return sys._MEIPASS
    else:
        # The application is running as a plain script
        return os.path.dirname(os.path.abspath(__file__))



