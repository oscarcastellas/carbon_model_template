# User Guide - Carbon Model Template

## Overview

The Carbon Model Template is a comprehensive financial analysis tool for carbon credit streaming deals. It provides both a simple GUI for non-technical users and Python scripts for advanced analysis.

## Quick Start

### Step 1: Run GUI Application

1. Launch the GUI:
   ```bash
   python3 gui/run_gui.py
   ```

2. Upload your data file (Excel, Word, or PDF)

3. Review extracted assumptions (WACC, investment amount, streaming percentage, etc.)

4. Click "Run Analysis"

5. The GUI will generate an Excel file with:
   - Auto-populated data in "Inputs & Assumptions", "Valuation Schedule", and "Summary & Results" sheets
   - Professional charts for investor presentations
   - Empty analysis sheets ready for your inputs

### Step 2: Advanced Analysis (Optional)

After the GUI generates your Excel file:

1. **Open the Excel file**

2. **Navigate to analysis sheets:**
   - **Deal Valuation**: Find optimal purchase price for target IRR
   - **Monte Carlo Results**: Run stochastic simulations
   - **Sensitivity Analysis**: Analyze sensitivity to key variables
   - **Breakeven Analysis**: Calculate breakeven points

3. **Fill in input cells** in the analysis sheet you want to use

4. **Save the Excel file**

5. **Run the corresponding Python script from Terminal:**
   ```bash
   # Deal Valuation
   python3 scripts/run_deal_valuation_from_excel.py path/to/your_file.xlsx
   
   # Monte Carlo Simulation
   python3 scripts/run_monte_carlo_from_excel.py path/to/your_file.xlsx
   
   # Sensitivity Analysis
   python3 scripts/run_sensitivity_from_excel.py path/to/your_file.xlsx
   
   # Breakeven Analysis
   python3 scripts/run_breakeven_from_excel.py path/to/your_file.xlsx
   ```

6. **Results and charts are automatically written to Excel!**

## Excel File Structure

### Presentation Sheets (Auto-populated)

1. **Inputs & Assumptions**
   - All model inputs and assumptions
   - Professional charts showing key assumptions

2. **Valuation Schedule**
   - 20-year detailed cash flow table
   - All calculations are Excel formulas
   - Charts showing cash flow trends

3. **Summary & Results**
   - Key financial metrics (NPV, IRR, Payback)
   - Risk assessment (flags and scores)
   - Presentation-ready charts

4. **Analysis** (Separator)
   - Blank sheet separating presentation from analysis modules

### Analysis Sheets (User-driven)

5. **Deal Valuation**
   - Input cells: Target IRR, Purchase Price, Streaming %, Investment Tenor
   - Output: Maximum Purchase Price, Actual IRR, NPV, Required Streaming
   - Chart: Purchase price vs IRR visualization

6. **Monte Carlo Results**
   - Input cells: Number of simulations, GBM settings, Volume settings
   - Output: Statistical results (mean, std dev, percentiles)
   - Charts: IRR and NPV histograms

7. **Sensitivity Analysis**
   - Input cells: Variable ranges for sensitivity analysis
   - Output: 2D sensitivity table
   - Chart: Sensitivity heatmap

8. **Breakeven Analysis**
   - Input cells: Target metric (price, volume, or streaming)
   - Output: Breakeven values
   - Chart: Breakeven visualization

## Tips

- **All calculations use Excel formulas** - you can modify inputs and see results update automatically
- **Charts are embedded** - no external files needed
- **Python scripts read and write directly to Excel** - no manual copy/paste needed
- **Master template is automatically used** - the GUI finds and uses it automatically

## Troubleshooting

### Python script can't find Excel file
- Make sure you provide the full path to the Excel file
- Check that the Excel file is saved and closed before running the script

### Results not appearing in Excel
- Make sure the Excel file is closed when running the Python script
- Check that you filled in all required input cells
- Verify the sheet names match (they should be: "Deal Valuation", "Monte Carlo Results", "Sensitivity Analysis", "Breakeven Analysis")

### Charts not appearing
- Charts are embedded as images in the Excel file
- Make sure you have write permissions for the Excel file
- Check that matplotlib is installed: `pip install matplotlib`

## Next Steps

- See `docs/MASTER_TEMPLATE_SIMPLIFIED_WORKFLOW_PLAN.md` for detailed implementation plan
- See `examples/basic_usage.py` for Python API usage
- See `docs/PROJECT_STRUCTURE.md` for project architecture

