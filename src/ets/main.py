"""
ETS Processor - Main Entry Point

This is the main entry point for the ETS Processor application.
"""

import sys
import os

# Add project root to path for both frozen and unfrozen environments
if getattr(sys, 'frozen', False):
    # Running as compiled executable
    base_path = sys._MEIPASS
    sys.path.insert(0, base_path)
    sys.path.insert(0, os.path.join(base_path, 'src'))
else:
    # Running as normal Python script
    base_path = os.path.dirname(os.path.abspath(__file__))
    sys.path.insert(0, os.path.abspath(os.path.join(base_path, '..', '..')))

# Silence output during startup for frozen executable
if getattr(sys, 'frozen', False):
    # Save original stdout/stderr
    original_stdout = sys.stdout
    original_stderr = sys.stderr
    
    # Redirect to null
    sys.stdout = open(os.devnull, 'w')
    sys.stderr = open(os.devnull, 'w')
    
    def enable_output():
        """Restore original output streams"""
        sys.stdout.close()
        sys.stderr.close()
        sys.stdout = original_stdout
        sys.stderr = original_stderr

from src.ets.gui import main as ets_gui_main

if __name__ == "__main__":
    if getattr(sys, 'frozen', False):
        # In frozen mode, run main and then restore output
        ets_gui_main()
        enable_output()
    else:
        # In development mode, just run main
        ets_gui_main()