#!/usr/bin/env python3
"""
Launcher script for Carbon Model GUI Application

This script can be used to run the GUI directly or as entry point
for PyInstaller executable.
"""

import sys
import os
from pathlib import Path

# Add project root to path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

try:
    from gui.carbon_model_gui import main
    
    if __name__ == "__main__":
        main()
except ImportError as e:
    print(f"Error importing GUI module: {e}")
    print("\nPlease make sure all dependencies are installed:")
    print("  pip install -r requirements.txt")
    input("\nPress Enter to exit...")
    sys.exit(1)
except Exception as e:
    print(f"Error starting application: {e}")
    import traceback
    traceback.print_exc()
    input("\nPress Enter to exit...")
    sys.exit(1)

