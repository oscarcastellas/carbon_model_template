# Interactive Excel Analysis Guide

## 🎯 Overview

The Excel output now includes an **Interactive Analysis** sheet that allows you to:
1. **Adjust GBM and Monte Carlo parameters** directly in Excel
2. **Run Python script** to update results automatically
3. **View all charts** embedded in Excel
4. **No need to edit Python code** - everything is in Excel!

---

## 📊 Excel File Structure

### **Sheet 1: Inputs & Assumptions**
- All model assumptions
- GBM parameters
- Monte Carlo settings

### **Sheet 2: Valuation Schedule**
- 20-year cash flow table
- All calculations as formulas

### **Sheet 3: Summary & Results**
- Key metrics
- Risk assessment
- Breakeven analysis
- Monte Carlo summary

### **Sheet 4: Monte Carlo Results**
- Full simulation results
- Histogram charts
- Statistical summary

### **Sheet 5: Interactive Analysis** ⭐ **NEW!**
- **Adjustable parameters**
- Current results display
- Instructions for running analysis

### **Sheet 6: Volatility Charts** ⭐ **NEW!**
- All 5 volatility charts embedded
- High-resolution images
- Ready for presentations

---

## 🔧 How to Use Interactive Analysis

### **Step 1: Open Excel File**

```bash
# Generate the Excel file first (if not already done)
python3 examples/generate_excel_with_charts.py
```

Open `carbon_model_with_charts.xlsx`

### **Step 2: Adjust Parameters**

Go to **"Interactive Analysis"** sheet and modify:

#### **Base Financial Assumptions**
- WACC
- Rubicon Investment Total
- Investment Tenor
- Initial Streaming Percentage

#### **GBM Parameters**
- Use GBM Method (Yes/No)
- GBM Drift (μ) - Expected Return
- GBM Volatility (σ) - Price Volatility

#### **Monte Carlo Parameters**
- Number of Simulations
- Price Growth Base (Mean)
- Price Growth Std Dev
- Volume Multiplier Base
- Volume Std Dev

### **Step 3: Run Analysis**

Save the Excel file, then run:

```bash
python3 examples/run_interactive_analysis.py
```

This script will:
1. ✅ Read your adjusted parameters from Excel
2. ✅ Run full Monte Carlo simulation
3. ✅ Generate updated charts
4. ✅ Update all Excel sheets with new results

### **Step 4: View Results**

Open the updated Excel file:
- **"Summary & Results"** - Updated metrics
- **"Monte Carlo Results"** - New simulation results
- **"Volatility Charts"** - Updated charts
- **"Interactive Analysis"** - Shows current results

---

## 📋 Parameter Guide

### **GBM Drift (μ)**
- **What it is**: Expected annual price return
- **Typical range**: 0.02 to 0.05 (2% to 5%)
- **Example**: 0.03 = 3% expected annual growth

### **GBM Volatility (σ)**
- **What it is**: Annual price volatility
- **Typical range**: 0.10 to 0.25 (10% to 25%)
- **Example**: 0.15 = 15% annual volatility
- **Higher = More risk**

### **Number of Simulations**
- **What it is**: How many scenarios to run
- **Recommended**: 5,000 (balance of accuracy and speed)
- **Minimum**: 1,000 for reliable results
- **More = More accurate but slower**

### **Price Growth Base**
- **What it is**: Mean annual price growth (if not using GBM)
- **Typical range**: 0.02 to 0.05

### **Volume Std Dev**
- **What it is**: Volume delivery uncertainty
- **Typical range**: 0.10 to 0.20 (10% to 20%)
- **Higher = More volume risk**

---

## 🎨 Viewing Charts in Excel

### **Volatility Charts Sheet**

All 5 charts are embedded:

1. **Price Volatility Paths**
   - Shows multiple GBM paths
   - Base forecast vs. stochastic outcomes

2. **Price Distribution Over Time**
   - Distribution at Years 5, 10, 15, 20
   - Shows uncertainty over time

3. **Returns Distribution**
   - IRR and NPV histograms
   - Risk distribution visualization

4. **Volatility Heatmap**
   - Percentile heatmap
   - Price ranges over time

5. **Correlation Analysis**
   - Relationship between volatility and returns
   - 4 scatter plots

**To view**: Simply open the "Volatility Charts" sheet!

---

## 🔄 Workflow Example

### **Scenario 1: Test Higher Volatility**

1. Open Excel → "Interactive Analysis" sheet
2. Change **GBM Volatility** from 15% to 25%
3. Save Excel file
4. Run: `python3 examples/run_interactive_analysis.py`
5. Open updated Excel → Check "Summary & Results"
6. Compare P10/P90 IRR ranges (should be wider)

### **Scenario 2: Test Lower Drift**

1. Open Excel → "Interactive Analysis" sheet
2. Change **GBM Drift** from 3% to 2%
3. Save Excel file
4. Run: `python3 examples/run_interactive_analysis.py`
5. Check Mean IRR (should be lower)

### **Scenario 3: Compare GBM vs. Growth-Rate**

1. Set **Use GBM Method** = No
2. Run analysis
3. Set **Use GBM Method** = Yes
4. Run analysis again
5. Compare results in "Summary & Results"

---

## 💡 Tips

### **Parameter Adjustment**
- **Start with small changes** to see impact
- **Change one parameter at a time** for clarity
- **Save Excel before running script**

### **Results Interpretation**
- **P10 IRR**: Downside risk (worst 10% of scenarios)
- **P90 IRR**: Upside potential (best 10% of scenarios)
- **Wider P10-P90 range** = Higher volatility risk

### **Performance**
- **5,000 simulations**: ~1-3 minutes
- **10,000 simulations**: ~2-5 minutes
- **1,000 simulations**: ~30 seconds (less accurate)

### **Charts**
- Charts are **high resolution** (300 DPI)
- Can be **copied from Excel** for presentations
- **Automatically updated** when you run analysis

---

## 🚨 Troubleshooting

### **"Excel file not found"**
```bash
# Generate the Excel file first
python3 examples/generate_excel_with_charts.py
```

### **"Parameters not reading correctly"**
- Make sure you're editing the **"Interactive Analysis"** sheet
- Parameters are in **Column B** (next to labels in Column A)
- **Save Excel file** before running script

### **"Charts not showing"**
- Charts are in **"Volatility Charts"** sheet
- Make sure charts were generated in `volatility_charts/` folder
- Re-run: `python3 examples/generate_excel_with_charts.py`

### **"Results not updating"**
- Make sure Excel file is **closed** when running script
- Check that script completed successfully
- Open Excel file again to see updates

---

## 📊 Quick Reference

### **Generate Excel with Charts**
```bash
python3 examples/generate_excel_with_charts.py
```

### **Run Interactive Analysis**
```bash
python3 examples/run_interactive_analysis.py
```

### **Excel File Location**
```
carbon_model_with_charts.xlsx
```

### **Charts Location**
```
volatility_charts/
```

---

## ✅ Summary

**You now have:**
- ✅ **Interactive Excel sheet** for parameter adjustment
- ✅ **All charts embedded** in Excel
- ✅ **Simple Python script** to update results
- ✅ **No need to edit Python code** - everything in Excel!

**Workflow:**
1. Adjust parameters in Excel
2. Run Python script
3. View updated results and charts

**That's it!** 🎯

