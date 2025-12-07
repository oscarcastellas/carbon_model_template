# How to Use the Carbon Model Template

## 🚀 Quick Start

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Run the Full Test

```bash
python3 test_excel.py
```

This will:
- ✅ Load your Excel data (`Analyst_Model_Test_OCC.xlsx`)
- ✅ Run DCF analysis (NPV, IRR)
- ✅ Perform goal-seeking for target IRR
- ✅ Run sensitivity analysis
- ✅ Run Monte Carlo simulation (5,000 runs)
- ✅ Export comprehensive Excel report with histogram charts

**Output:** `carbon_model_results.xlsx` with 5 sheets

---

## 📖 Complete Usage Guide

### Basic Workflow

```python
from carbon_model_template import CarbonModelGenerator

# Step 1: Initialize with assumptions
model = CarbonModelGenerator(
    wacc=0.08,                          # 8% discount rate
    rubicon_investment_total=20_000_000, # $20M investment
    investment_tenor=5,                  # 5 years deployment
    streaming_percentage_initial=0.48,   # 48% initial streaming
    
    # Monte Carlo assumptions (optional)
    price_growth_base=0.03,              # 3% mean price growth
    price_growth_std_dev=0.02,          # 2% price volatility
    volume_multiplier_base=1.0,         # 100% of base volume
    volume_std_dev=0.15                 # 15% volume volatility
)

# Step 2: Load your data
model.load_data("your_data_file.xlsx")

# Step 3: Run DCF analysis
results = model.run_dcf()
print(f"NPV: ${results['npv']:,.2f}")
print(f"IRR: {results['irr']:.2%}")

# Step 4: Goal-seeking (find streaming % for target IRR)
goal = model.find_target_irr_stream(target_irr=0.20)  # 20% target
print(f"Required streaming: {goal['streaming_percentage']:.2%}")

# Step 5: Sensitivity analysis
sensitivity = model.run_sensitivity_table(
    credit_range=[0.9, 1.0, 1.1],      # Credit volume multipliers
    price_range=[0.8, 1.0, 1.2]        # Price multipliers
)

# Step 6: Monte Carlo simulation
mc_results = model.run_monte_carlo(
    simulations=5000,
    use_percentage_variation=False     # Use growth-rate-based variation
)

# Step 7: Export to Excel (with all formulas!)
model.export_model_to_excel("my_results.xlsx")
```

---

## 📊 Input Data Format

### Required Data

Your Excel/CSV file should contain:

1. **Time Series**: Years 1 to 20
2. **Carbon Credits Issued (Gross)**: Annual tonnage of credits
3. **Project Implementation Costs**: Annual CAPEX in USD
4. **Base Carbon Price**: Price per ton in USD (your price forecast)

### Flexible Data Handling

The data loader handles:
- ✅ **Transposed formats**: Labels in rows, years in columns
- ✅ **Various column names**: "Carbon Credits", "Credits Issued", "Gross Credits", etc.
- ✅ **Multiple sheets**: Automatically searches common sheet names
- ✅ **Messy data**: Cleans headers and handles missing values
- ✅ **Assumption extraction**: Can extract assumptions from dedicated sheets

### Example Data Structure

**Standard Format:**
```
Year | Carbon Credits Gross | Project Implementation Costs | Base Carbon Price
-----|---------------------|------------------------------|------------------
  1  |   10,000,000        |    -130,130,374              |      0.00
  2  |    0                |        -628,486               |     40.00
  3  |    0                |      -1,333,184               |     41.60
...
```

**Transposed Format (also supported):**
```
                    | Year 1 | Year 2 | Year 3 | ...
--------------------|--------|--------|--------|----
Carbon Credits      | 10M    | 0      | 0      | ...
Project Costs       | -130M  | -628K  | -1.3M  | ...
Carbon Price        | 0.00   | 40.00  | 41.60  | ...
```

---

## 📈 Output Excel File

The exported Excel file contains **ALL calculations as formulas** for full transparency:

### Sheet 1: Inputs & Assumptions
- All model inputs (WACC, Investment, Tenor, Streaming %)
- Monte Carlo assumptions
- Yellow highlighting for input cells

### Sheet 2: Valuation Schedule
- 20-year detailed cash flow table
- **All formulas linked to inputs:**
  - Rubicon Share = Credits × Streaming %
  - Revenue = Share × Price
  - Investment Drawdown = IF(Year <= Tenor, -Investment/Tenor, 0)
  - Net Cash Flow = Revenue + Investment
  - Discount Factor = 1/(1+WACC)^(Year-1)
  - Present Value = Net CF × Discount Factor
  - Cumulative formulas
- Green highlighting for formula cells
- Bold totals row

### Sheet 3: Summary & Results
- **NPV** = SUM(Valuation Schedule PV column)
- **IRR** = IRR(Valuation Schedule Cash Flow column)
- **Payback Period** = MATCH formula
- Target metrics linked to inputs
- Monte Carlo summary statistics

### Sheet 4: Sensitivity Analysis
- 2D IRR sensitivity table
- Credit Volume vs. Carbon Price multipliers
- Formatted as percentages

### Sheet 5: Monte Carlo Results
- Summary statistics (Mean, P10, P90, Std Dev)
- Full 5,000 simulation results (IRR and NPV columns)
- **IRR Distribution Histogram Chart**
- **NPV Distribution Histogram Chart**

---

## 🎯 Key Features

### 1. Formula-Based Excel Output
- **Every calculation is a formula**, not a hardcoded value
- Full transparency and auditability
- Change inputs → see results update automatically
- Perfect for external review

### 2. Monte Carlo Simulation
- Uses **YOUR original price forecasts** as the base
- Applies stochastic variations around your assumptions
- Two variation methods:
  - **Growth-rate-based** (default): Models uncertainty in price growth
  - **Percentage-based**: Models uniform percentage uncertainty
- Dual-variable stochastic modeling (price + volume)

### 3. Robust Data Handling
- Handles messy, unstructured data
- Automatic column detection
- Transposed format support
- Assumption extraction from Excel

### 4. Comprehensive Analysis
- DCF analysis (NPV, IRR)
- Goal-seeking optimization
- Sensitivity analysis
- Payback period calculation
- Monte Carlo risk assessment

---

## 🔧 Advanced Usage

### Providing Assumptions

**Option 1: Direct initialization**
```python
model = CarbonModelGenerator(
    wacc=0.08,
    rubicon_investment_total=20_000_000,
    investment_tenor=5,
    streaming_percentage_initial=0.48
)
```

**Option 2: Extract from Excel**
```python
model = CarbonModelGenerator()
model.load_data_with_assumptions("data_with_assumptions.xlsx")
# Assumptions extracted from "Assumptions" sheet
```

**Option 3: Set after initialization**
```python
model = CarbonModelGenerator()
model.set_assumptions({
    'wacc': 0.08,
    'rubicon_investment_total': 20_000_000,
    'investment_tenor': 5,
    'streaming_percentage_initial': 0.48
})
```

### Monte Carlo Options

**Growth-Rate-Based Variation (Default)**
```python
mc_results = model.run_monte_carlo(
    simulations=5000,
    use_percentage_variation=False  # Default
)
```
- Applies stochastic deviations to growth rates from your price curve
- Preserves your original forecast pattern
- Best for modeling price growth uncertainty

**Percentage-Based Variation**
```python
mc_results = model.run_monte_carlo(
    simulations=5000,
    use_percentage_variation=True
)
```
- Applies percentage multipliers directly to prices
- Models uniform percentage uncertainty
- Best for uniform price volatility

---

## 📚 Additional Documentation

- **`README.md`**: Project overview and installation
- **`ARCHITECTURE.md`**: System architecture and design
- **`EXCEL_FORMULA_GUIDE.md`**: Detailed Excel formula documentation
- **`MONTE_CARLO_GUIDE.md`**: Monte Carlo simulation guide
- **`example_usage.py`**: Code examples
- **`example_assumptions.py`**: Assumption handling examples

---

## 🐛 Troubleshooting

### Import Errors
```bash
# Make sure you're in the project directory
cd /path/to/carbon_model_template
python3 test_excel.py
```

### Data Loading Issues
- Check that your Excel file has the required columns
- The loader handles transposed formats automatically
- Check console output for column detection messages

### IRR Calculation Warnings
- Some cash flow patterns don't have valid IRRs
- This is normal for certain scenarios
- The model continues with NaN values where appropriate

### Monte Carlo Results
- Ensure you have valid price forecasts in your data
- Check that `price_growth_std_dev` and `volume_std_dev` are set
- Review the Monte Carlo guide for parameter selection

---

## 🎓 Best Practices

1. **Start with the test script**: Run `test_excel.py` to verify everything works
2. **Use your price forecasts**: The Monte Carlo uses your original forecasts as the base
3. **Set appropriate volatility**: 
   - Price: 1-3% for stable markets, 3-5% for volatile
   - Volume: 10-20% for delivery uncertainty
4. **Run enough simulations**: 5,000 provides stable statistics
5. **Review Excel formulas**: All calculations are transparent and auditable

---

## 📝 Example Scripts

### Quick Test (No Monte Carlo)
```bash
python3 quick_test.py
```
Fast test without Monte Carlo simulation

### Full Test (With Monte Carlo)
```bash
python3 test_excel.py
```
Complete test with all features including Monte Carlo

---

## 🚀 Ready to Use!

The template is ready for production use. Simply:
1. Replace `Analyst_Model_Test_OCC.xlsx` with your data
2. Adjust assumptions in the test script
3. Run the analysis
4. Review the comprehensive Excel output

All calculations are formula-based, fully transparent, and ready for external review!
