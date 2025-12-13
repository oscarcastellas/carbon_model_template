# Multi-File Support - Documentation

## 🎯 Overview

The Carbon Model Tool now supports loading data from **multiple file types**:
- ✅ **Excel** (.xlsx, .xls) - Primary format
- ✅ **Word** (.docx, .doc) - Extracts tables and text
- ✅ **PDF** (.pdf) - Extracts tables and text

This allows you to combine data from different sources into a single analysis.

---

## 📊 How It Works

### **File Type Detection**

The system automatically detects file type from extension:
- `.xlsx`, `.xls` → Excel (uses existing robust loader)
- `.docx`, `.doc` → Word (extracts tables and text)
- `.pdf` → PDF (extracts tables and text)

### **Data Extraction**

**For Excel:**
- Uses existing `DataLoader` with full robust handling
- Supports transposed formats, multiple sheets, etc.

**For Word/PDF:**
- Extracts all tables from document
- Converts tables to pandas DataFrames
- Extracts text content
- Searches for key-value pairs (e.g., "WACC: 8%")
- Finds best matching data table automatically

### **Table Selection**

The system automatically finds the best data table by scoring:
- ✅ Has "Year" column → +10 points
- ✅ Has financial columns (price, cost, credit) → +5 points each
- ✅ Has ~20 rows (for 20-year model) → +10 points
- ✅ Has numeric data → +5 points

**Highest scoring table is used automatically!**

---

## 🖥️ GUI Usage

### **Step 1: Add Files**

1. Click "➕ Add Files..." button
2. Select one or more files:
   - Excel files (.xlsx, .xls)
   - Word documents (.docx, .doc)
   - PDF files (.pdf)
3. Files appear in the list

### **Step 2: Remove Files (Optional)**

1. Select file(s) in the list
2. Click "➖ Remove Selected" button

### **Step 3: Run Analysis**

1. Click "▶ Run Analysis"
2. System automatically:
   - Detects file types
   - Extracts data from each file
   - Combines data intelligently
   - Runs analysis

---

## 🔍 Data Extraction Details

### **Word Documents (.docx, .doc)**

**What gets extracted:**
- All tables → Converted to DataFrames
- All text → Searched for key-value pairs
- Assumptions → Automatically extracted

**Example patterns found:**
- "WACC: 8%" → Extracted as `wacc: 0.08`
- "Investment: $20,000,000" → Extracted as `rubicon_investment_total: 20000000`
- "Streaming: 48%" → Extracted as `streaming_percentage: 0.48`

### **PDF Files (.pdf)**

**What gets extracted:**
- All tables → Converted to DataFrames
- All text → Searched for key-value pairs
- Uses `pdfplumber` (better) or `PyPDF2` (fallback)

**Example patterns found:**
- Same as Word documents
- Financial terms automatically recognized

### **Excel Files (.xlsx, .xls)**

**What gets extracted:**
- Full robust data loading (existing functionality)
- Multiple sheets support
- Transposed format detection
- Assumption extraction from "Assumptions" sheet

---

## 🎯 Use Cases

### **Use Case 1: Excel + Word Contract**

**Scenario:**
- Excel file has financial data
- Word document has contract terms (assumptions)

**How to use:**
1. Add both files to GUI
2. System extracts:
   - Data from Excel
   - Assumptions from Word
3. Combines automatically

### **Use Case 2: Multiple Excel Files**

**Scenario:**
- File 1: Price forecasts
- File 2: Volume data
- File 3: Cost data

**How to use:**
1. Add all three Excel files
2. System combines data intelligently
3. Uses best matching tables

### **Use Case 3: PDF Report**

**Scenario:**
- PDF contains project summary with tables

**How to use:**
1. Add PDF file
2. System extracts tables
3. Finds best matching data table
4. Runs analysis

---

## ⚙️ Technical Details

### **Dependencies**

**Required:**
- `python-docx` - For Word documents
- `PyPDF2` or `pdfplumber` - For PDF files

**Installation:**
```bash
pip install python-docx PyPDF2 pdfplumber
```

### **File Processing Order**

1. **Detect file type** from extension
2. **Extract data** using appropriate method
3. **Score tables** to find best match
4. **Combine data** from multiple files
5. **Merge assumptions** from all sources
6. **Validate** required columns exist
7. **Run analysis**

### **Error Handling**

- ✅ Missing files → Warning, skip file
- ✅ Unsupported format → Warning, skip file
- ✅ No data table found → Clear error message
- ✅ Missing columns → Attempts automatic mapping
- ✅ Extraction errors → Graceful fallback

---

## 📋 Supported File Formats

| Format | Extension | Status | Notes |
|--------|-----------|--------|-------|
| Excel | .xlsx, .xls | ✅ Full Support | Robust loader with all features |
| Word | .docx | ✅ Supported | Extracts tables and text |
| Word | .doc | ⚠️ Limited | May not work (old format) |
| PDF | .pdf | ✅ Supported | Extracts tables and text |

---

## 🔧 Column Mapping

The system automatically maps common column names:

**Carbon Credits:**
- "Carbon Credits", "Credits", "Gross Credits", "Tonnage", "Tons"

**Carbon Price:**
- "Carbon Price", "Price", "Price per Ton", "Price/Ton"

**Project Costs:**
- "Project Costs", "Costs", "CAPEX", "Capital", "Implementation Costs"

---

## 💡 Tips

1. **Best Results:**
   - Use Excel for primary data (most robust)
   - Use Word/PDF for assumptions or supplementary data

2. **Table Format:**
   - Tables with clear headers work best
   - Year column helps identification
   - ~20 rows ideal for 20-year model

3. **Multiple Files:**
   - System combines intelligently
   - Excel data takes priority
   - Assumptions merged from all sources

4. **Troubleshooting:**
   - If table not found, check table has headers
   - Ensure Year column exists (or can be inferred)
   - Verify financial columns are present

---

## 🚀 Example Workflow

### **Scenario: Excel + PDF**

1. **Excel file** (`data.xlsx`):
   - Contains: Year, Credits, Price, Costs
   - Primary data source

2. **PDF file** (`contract.pdf`):
   - Contains: Contract terms
   - Has table with assumptions

3. **GUI Steps:**
   - Add both files
   - Click "Run Analysis"
   - System:
     - Extracts data from Excel
     - Extracts assumptions from PDF
     - Combines automatically
     - Runs analysis

4. **Result:**
   - Complete analysis with data from Excel
   - Assumptions from PDF automatically applied

---

## ✅ Summary

**New Capabilities:**
- ✅ Load from Excel, Word, or PDF
- ✅ Multiple files at once
- ✅ Automatic data extraction
- ✅ Intelligent table selection
- ✅ Assumption extraction
- ✅ Seamless combination

**User Experience:**
- ✅ Just add files and click "Run"
- ✅ No manual data copying needed
- ✅ Works with messy, unstructured files
- ✅ Automatic error handling

**This makes the tool even more powerful and user-friendly!** 🎯

