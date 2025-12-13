# How to Start the GUI Application

## 🚀 Quick Start

### **In a New Terminal:**

```bash
cd /path/to/carbon_model_template
python3 gui/run_gui.py
```

That's it! The GUI window will open.

---

## 📋 Step-by-Step

1. **Open Terminal** (or Command Prompt on Windows)

2. **Navigate to project folder:**
   ```bash
   cd /path/to/carbon_model_template
   ```

3. **Run the GUI:**
   ```bash
   python3 gui/run_gui.py
   ```

4. **GUI window opens** - you're ready to use it!

---

## 🖥️ Alternative: Direct Path

If you're already in a different directory, you can run:

```bash
python3 /path/to/carbon_model_template/gui/run_gui.py
```

---

## ✅ Verify It Works

After running, you should see:
- GUI window opens
- Title: "Carbon Model Analysis Tool"
- File browser section
- Options section
- Run Analysis button

---

## 🆘 Troubleshooting

**Problem: "command not found: python3"**
- Try: `python gui/run_gui.py` (without the "3")
- Or install Python 3.8+

**Problem: "No module named 'tkinter'"**
- Install tkinter: `pip install tk` (Linux)
- Or use system package manager

**Problem: Import errors**
- Make sure you're in the project root directory
- Install dependencies: `pip install -r requirements.txt`

---

**That's it! Just two commands and you're running!** 🎯

