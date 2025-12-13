# GUI Application - Packaging & Distribution Guide

## 🎯 Overview

This guide explains how to package the GUI application as a portable zip folder for distribution to colleagues.

---

## 📦 What Gets Created

After packaging, you'll have:

```
Carbon Model Tool - Portable/
├── Carbon Model Tool.exe (or .app on Mac)
├── README.txt
└── (optional) sample_data.xlsx
```

**Total size:** ~50-100 MB (depending on platform)

---

## 🚀 Quick Start: Create Package

### **Step 1: Run Packaging Script**

```bash
python3 package_gui.py
```

This will:
1. ✅ Create standalone executable using PyInstaller
2. ✅ Package everything into a portable folder
3. ✅ Include README with instructions
4. ✅ Ready to zip and share!

### **Step 2: Test the Package**

1. Go to `Carbon Model Tool - Portable/` folder
2. Double-click the executable
3. Test with sample data
4. Verify it works correctly

### **Step 3: Create Zip File**

1. Right-click `Carbon Model Tool - Portable` folder
2. Select "Compress" (Mac) or "Send to > Compressed folder" (Windows)
3. Name it: `Carbon_Model_Tool_v1.0.zip`

### **Step 4: Distribute**

**Option A: Email** (if under 25 MB)
- Attach zip file
- Send to colleagues

**Option B: Shared Drive**
- Upload zip to company shared drive
- Share link with colleagues

**Option C: File Sharing Service**
- Upload to Dropbox/OneDrive/Google Drive
- Share download link

---

## 📋 Detailed Packaging Steps

### **Prerequisites**

1. **Install PyInstaller:**
   ```bash
   pip install pyinstaller
   ```

2. **Install all dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

### **Manual Packaging (if script doesn't work)**

1. **Create executable:**
   ```bash
   pyinstaller --onefile --windowed --name "Carbon Model Tool" gui/run_gui.py
   ```

2. **Find executable:**
   - Windows: `dist/Carbon Model Tool.exe`
   - Mac: `dist/Carbon Model Tool.app`
   - Linux: `dist/Carbon Model Tool`

3. **Create package folder:**
   ```bash
   mkdir "Carbon Model Tool - Portable"
   ```

4. **Copy files:**
   ```bash
   cp "dist/Carbon Model Tool.exe" "Carbon Model Tool - Portable/"
   cp README_GUI.txt "Carbon Model Tool - Portable/README.txt"
   ```

5. **Zip the folder:**
   - Right-click folder → Compress/Zip
   - Name: `Carbon_Model_Tool_v1.0.zip`

---

## 🎨 GUI Features

### **Main Window:**
- ✅ Professional header with title
- ✅ Input file browser
- ✅ Output file selection (optional)
- ✅ Analysis options (checkboxes)
- ✅ Progress bar with status updates
- ✅ Run Analysis button
- ✅ Help button

### **User Experience:**
- ✅ Clean, modern interface
- ✅ Real-time progress updates
- ✅ Friendly error messages
- ✅ Option to open results automatically
- ✅ Help window with instructions

---

## 📝 What Colleagues Need to Do

### **1. Receive Zip File**
- Download from email/shared drive
- Extract to any folder

### **2. Run Application**
- Double-click executable
- GUI window opens

### **3. Select Data File**
- Click "Browse" button
- Choose their Excel data file

### **4. Run Analysis**
- (Optional) Adjust options
- Click "Run Analysis"
- Wait for completion

### **5. View Results**
- Click "Yes" to open results
- Or manually open `results.xlsx`

**That's it!** No Python, no command line, no technical knowledge needed.

---

## 🔧 Troubleshooting Packaging

### **Problem: PyInstaller not found**
```bash
pip install pyinstaller
```

### **Problem: Import errors during packaging**
- Make sure all dependencies are installed
- Check that all modules are in correct locations
- Verify `requirements.txt` is up to date

### **Problem: Executable is too large**
- This is normal (50-100 MB)
- All Python dependencies are bundled
- No way to make it smaller without removing features

### **Problem: Executable doesn't run on other computers**
- Make sure you're packaging for the correct platform
- Windows: Package on Windows
- Mac: Package on Mac
- Linux: Package on Linux

### **Problem: Antivirus flags executable**
- This is a false positive (common with PyInstaller)
- Options:
  1. Sign the executable (requires certificate)
  2. Add to antivirus whitelist
  3. Explain to colleagues it's safe

---

## 📊 File Size Estimates

| Component | Size |
|-----------|------|
| Python Runtime | ~30 MB |
| Dependencies (pandas, numpy, etc.) | ~40 MB |
| Application Code | ~5 MB |
| **Total** | **~75 MB** |

**Compressed (zip):** ~25-35 MB

---

## 🎯 Distribution Checklist

Before sharing with colleagues:

- [ ] Test executable on clean system (no Python installed)
- [ ] Verify all features work
- [ ] Check README.txt is clear
- [ ] Test with sample data
- [ ] Create zip file
- [ ] Test zip extraction
- [ ] Write email with instructions

---

## 📧 Sample Distribution Email

```
Subject: Carbon Model Analysis Tool - Ready to Use!

Hi Team,

I've created a simple tool for running carbon model analysis. 
No Python knowledge needed - just double-click and go!

HOW TO USE:
1. Download the attached zip file
2. Extract to any folder
3. Double-click "Carbon Model Tool.exe"
4. Select your Excel data file
5. Click "Run Analysis"
6. View results in Excel!

The tool includes:
- Full DCF analysis
- Monte Carlo simulation
- Risk assessment
- Professional Excel output

Let me know if you have any questions!

Best,
[Your Name]
```

---

## ✅ Success Criteria

The package is successful if:
- ✅ Colleagues can run it without Python
- ✅ No technical knowledge required
- ✅ Clear instructions provided
- ✅ Works on their computers
- ✅ Results are accurate

---

## 🔄 Updates

To update the tool:

1. Make changes to code
2. Run `package_gui.py` again
3. Test new version
4. Create new zip file
5. Share with version number (v1.1, v1.2, etc.)

---

**Ready to package? Run: `python3 package_gui.py`** 🚀

