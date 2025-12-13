#!/usr/bin/env python3
"""
Package GUI Application for Distribution

Creates a portable zip folder with executable and all required files.
"""

import os
import shutil
import subprocess
import sys
from pathlib import Path


def create_executable():
    """Create standalone executable using PyInstaller."""
    print("="*70)
    print("PACKAGING CARBON MODEL GUI APPLICATION")
    print("="*70)
    print()
    
    # Check if PyInstaller is installed
    try:
        import PyInstaller
    except ImportError:
        print("PyInstaller not found. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("✓ PyInstaller installed")
        print()
    
    # Create spec file
    print("Creating PyInstaller spec file...")
    spec_content = """# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['gui/run_gui.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'pandas',
        'numpy',
        'scipy',
        'xlsxwriter',
        'openpyxl',
        'tkinter',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.ttk',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Carbon Model Tool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # No console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Can add icon file here if available
)
"""
    
    spec_file = Path("carbon_model_tool.spec")
    with open(spec_file, 'w') as f:
        f.write(spec_content)
    print(f"✓ Created {spec_file}")
    print()
    
    # Build executable
    print("Building executable (this may take 5-10 minutes)...")
    print()
    try:
        subprocess.check_call([
            "pyinstaller",
            "--clean",
            "carbon_model_tool.spec"
        ])
        print()
        print("✓ Executable created successfully!")
        print()
    except subprocess.CalledProcessError as e:
        print(f"✗ Error building executable: {e}")
        return False
    
    return True


def create_zip_package():
    """Create zip folder with executable and documentation."""
    print("="*70)
    print("CREATING PORTABLE ZIP PACKAGE")
    print("="*70)
    print()
    
    # Determine executable name based on platform
    import platform
    if platform.system() == "Windows":
        exe_name = "Carbon Model Tool.exe"
        exe_path = Path("dist") / exe_name
    elif platform.system() == "Darwin":  # macOS
        exe_name = "Carbon Model Tool.app"
        exe_path = Path("dist") / exe_name
    else:  # Linux
        exe_name = "Carbon Model Tool"
        exe_path = Path("dist") / exe_name
    
    if not exe_path.exists():
        print(f"✗ Executable not found: {exe_path}")
        print("Please run create_executable() first.")
        return False
    
    # Create package directory
    package_dir = Path("Carbon Model Tool - Portable")
    if package_dir.exists():
        shutil.rmtree(package_dir)
    package_dir.mkdir()
    
    print(f"Creating package in: {package_dir}")
    print()
    
    # Copy executable
    print(f"Copying executable...")
    if exe_path.is_dir():  # .app bundle on Mac
        shutil.copytree(exe_path, package_dir / exe_name)
    else:
        shutil.copy2(exe_path, package_dir / exe_name)
    print(f"✓ Copied {exe_name}")
    
    # Create README
    print("Creating README.txt...")
    readme_content = """═══════════════════════════════════════════════════════
  CARBON MODEL ANALYSIS TOOL
  Simple Instructions for Non-Technical Users
═══════════════════════════════════════════════════════

HOW TO USE:

1. Extract this zip file to any folder on your computer

2. Prepare your data:
   - Open Excel
   - Create a file with your project data
   - Save it as an Excel file (.xlsx)

3. Run the analysis:
   - Double-click "Carbon Model Tool.exe" (or .app on Mac)
   - Click "Browse" and select your Excel data file
   - (Optional) Adjust analysis options
   - Click "Run Analysis" button
   - Wait for "Analysis complete!" message

4. View results:
   - Click "Yes" when asked to open results
   - Or manually open "results.xlsx" file
   - Check all sheets for complete analysis

WHAT YOU GET:

The tool generates a comprehensive Excel file with:
  • Inputs & Assumptions
  • Valuation Schedule (20-year cash flows)
  • Summary & Results (key metrics, risk analysis)
  • Monte Carlo Results (if enabled)
  • Volatility Charts (if enabled)

ANALYSIS OPTIONS:

☑ Run Monte Carlo Simulation
  - Runs 5,000 simulations by default
  - Shows risk distribution of returns
  - Takes 1-3 minutes

☑ Use GBM (Geometric Brownian Motion)
  - Advanced price volatility modeling
  - More realistic price scenarios

☑ Generate Charts
  - Creates visual charts of results
  - Embedded in Excel output

TROUBLESHOOTING:

Problem: "No file selected"
Solution: Make sure you clicked "Browse" and selected an Excel file

Problem: Analysis takes a long time
Solution: This is normal! Large analyses can take 2-5 minutes.
         Reduce "Simulations" number to speed up.

Problem: Error message appears
Solution: Check that your Excel file has the required columns:
         - Year
         - Carbon Credits Gross
         - Base Carbon Price
         - Project Implementation Costs

Problem: Results file not created
Solution: Check that you have write permissions in the folder
         Try running as administrator (Windows) or with sudo (Mac/Linux)

SYSTEM REQUIREMENTS:

• Windows 10/11, macOS 10.14+, or Linux
• No Python installation needed
• No additional software required

SUPPORT:

For questions or issues, contact your team lead.

═══════════════════════════════════════════════════════
"""
    
    readme_file = package_dir / "README.txt"
    with open(readme_file, 'w') as f:
        f.write(readme_content)
    print("✓ Created README.txt")
    
    # Create sample data template (optional)
    print("Creating sample data template...")
    try:
        # Copy a sample file if it exists
        sample_file = Path("Analyst_Model_Test_OCC.xlsx")
        if sample_file.exists():
            shutil.copy2(sample_file, package_dir / "sample_data.xlsx")
            print("✓ Included sample_data.xlsx")
    except:
        print("  (No sample data file found - skipping)")
    
    print()
    print("="*70)
    print("PACKAGE CREATED SUCCESSFULLY!")
    print("="*70)
    print()
    print(f"📦 Package location: {package_dir.absolute()}")
    print()
    print("Next steps:")
    print("1. Test the executable in the package folder")
    print("2. Zip the entire 'Carbon Model Tool - Portable' folder")
    print("3. Share the zip file with your colleagues")
    print()
    print("Package contents:")
    for item in package_dir.iterdir():
        if item.is_file():
            size = item.stat().st_size / (1024 * 1024)  # MB
            print(f"  • {item.name} ({size:.1f} MB)")
        else:
            print(f"  • {item.name}/ (folder)")
    print()
    
    return True


def main():
    """Main packaging function."""
    print()
    print("="*70)
    print("CARBON MODEL GUI - PACKAGING TOOL")
    print("="*70)
    print()
    
    # Step 1: Create executable
    if not create_executable():
        print("Failed to create executable. Exiting.")
        return
    
    print()
    
    # Step 2: Create zip package
    if not create_zip_package():
        print("Failed to create package. Exiting.")
        return
    
    print()
    print("="*70)
    print("ALL DONE!")
    print("="*70)
    print()
    print("Your portable package is ready to share!")
    print()


if __name__ == "__main__":
    main()

