# Carbon Model Template

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Code Style](https://img.shields.io/badge/Code%20Style-PEP8-orange.svg)](https://www.python.org/dev/peps/pep-0008/)

A comprehensive, production-ready Python framework for rapid DCF (Discounted Cash Flow) and Carbon Streaming financial analysis. This tool prioritizes robust data handling, automated financial calculations, and **fully formula-based Excel outputs** for complete transparency and auditability.

## 🎯 Project Highlights

- **Simple GUI Application**: Professional graphical interface for non-technical users - no Python knowledge required!
- **Formula-Based Excel Output**: Every calculation is an Excel formula, ensuring full transparency and auditability for external review
- **Monte Carlo Simulation**: Dual-variable stochastic modeling with GBM (Geometric Brownian Motion) support and histogram visualizations
- **Advanced Risk Analysis**: Automated risk flagging, scoring, and breakeven analysis
- **Robust Data Handling**: Intelligently processes messy, unstructured Excel/CSV files with automatic format detection
- **Modular Architecture**: Clean separation of concerns for easy maintenance and extension
- **100% Local Processing**: All data stays on your machine - no external services or data transmission

## 🛠️ Technologies Used

- **Python 3.8+** - Core language
- **pandas** - Data manipulation and analysis
- **numpy** - Numerical computations
- **scipy** - Scientific computing (optimization, IRR calculation)
- **openpyxl** - Excel file reading
- **xlsxwriter** - Excel file writing with formulas and charts

## ✨ Key Features

### Core Financial Analysis
- ✅ **DCF Analysis**: NPV, IRR, cash flow calculations
- ✅ **Goal-Seeking Optimization**: Find optimal streaming percentage for target IRR
- ✅ **Sensitivity Analysis**: 2D sensitivity tables for risk assessment
- ✅ **Payback Period Calculation**: Simple and discounted payback

### Advanced Risk Modeling
- ✅ **Monte Carlo Simulation**: 5,000+ simulations with statistical analysis
  - **GBM Support**: Industry-standard Geometric Brownian Motion for price volatility
  - **Growth-Rate Method**: Alternative method respecting your price forecasts
- ✅ **Risk Flagging**: Automatic red/yellow/green risk indicators
- ✅ **Risk Scoring**: 0-100 risk score with component breakdown
- ✅ **Breakeven Analysis**: Calculate breakeven price, volume, and streaming percentage

### Data & Export
- ✅ **Robust Data Loading**: Handles transposed formats, various column names, messy data
- ✅ **Assumption Extraction**: Automatically extracts assumptions from Excel files
- ✅ **Excel Export**: Multi-sheet output with all formulas and histogram charts

## 🚀 Quick Start

### **Option 1: GUI Application (Recommended for Non-Technical Users)**

**For Colleagues:**
1. Download the portable zip package
2. Extract and double-click `Carbon Model Tool.exe`
3. Select your Excel data file
4. Click "Run Analysis"
5. View results in Excel with auto-populated data and charts!

**For Developers:**
```bash
# Test GUI locally
python3 gui/run_gui.py

# Create package for distribution
python3 package_gui.py
```

**Advanced Analysis Workflow:**
After the GUI generates the Excel file:
1. Open the Excel file
2. Navigate to analysis sheets (Deal Valuation, Monte Carlo, Sensitivity, Breakeven)
3. Fill in the input cells with your assumptions
4. Save the Excel file
5. Run Python scripts from Terminal:
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
6. Results and charts are automatically written to Excel!

### **Option 2: Python API (For Technical Users)**

```bash
# Clone the repository
git clone https://github.com/oscarcastellas/carbon_model_template.git
cd carbon_model_template

# Install dependencies
pip install -r requirements.txt
```

### Basic Usage

```python
from carbon_model_generator import CarbonModelGenerator

# Initialize with assumptions
model = CarbonModelGenerator(
    wacc=0.08,
    rubicon_investment_total=20_000_000,
    investment_tenor=5,
    streaming_percentage_initial=0.48
)

# Load data
model.load_data("your_data.xlsx")

# Run DCF analysis
results = model.run_dcf()
print(f"NPV: ${results['npv']:,.2f}")
print(f"IRR: {results['irr']:.2%}")

# Goal-seeking
goal = model.find_target_irr_stream(target_irr=0.20)

# Risk analysis (automatic after DCF)
risk_flags = model.flag_risks()
risk_score = model.calculate_risk_score()
print(f"Risk Level: {risk_flags['risk_level']}")
print(f"Risk Score: {risk_score['overall_risk_score']}/100")

# Breakeven analysis
breakeven = model.calculate_breakeven(metric='all')

# Monte Carlo simulation (with GBM)
mc_results = model.run_monte_carlo(
    simulations=5000,
    use_gbm=True,              # Use Geometric Brownian Motion
    gbm_drift=0.03,            # 3% expected return
    gbm_volatility=0.15        # 15% volatility
)

# Export to Excel (all formulas!)
model.export_model_to_excel("results.xlsx")
```

### Run Tests

```bash
# Full test with Monte Carlo (takes 2-5 minutes)
python3 tests/generate_full_excel.py

# Quick test without Monte Carlo
python3 tests/quick_test.py

# Test productivity tools
python3 tests/test_productivity_tools.py
```

## 📊 Output

The tool generates a comprehensive Excel file with 8 sheets:

### Presentation Sheets (Auto-populated by GUI)
1. **Inputs & Assumptions** - All model inputs with professional charts
2. **Valuation Schedule** - 20-year detailed cash flow table with formulas and visualizations
3. **Summary & Results** - Key financial metrics, risk assessment, and presentation charts
4. **Analysis** - Separator sheet for advanced analysis modules

### Analysis Sheets (User-driven workflow)
5. **Deal Valuation** - Back-solver to find optimal purchase price for target IRR
6. **Monte Carlo Results** - Stochastic simulation with GBM support and histogram charts
7. **Sensitivity Analysis** - 2D sensitivity tables for risk assessment
8. **Breakeven Analysis** - Breakeven calculations for price, volume, and streaming percentage

**Workflow:**
1. GUI auto-populates sheets 1-3 with data and presentation charts
2. User fills input cells in analysis sheets (5-8)
3. User runs Python scripts from Terminal to execute analysis
4. Python scripts write results and charts directly to Excel

**All calculations are Excel formulas** - no hardcoded values. Change inputs and see results update automatically!

## 📁 Project Structure

```
carbon_model_template/
├── carbon_model_generator.py  # Main orchestrator class
├── core/                      # Core financial calculations
│   ├── dcf.py                 # DCF calculations
│   ├── irr.py                 # IRR calculations
│   └── payback.py             # Payback period
├── analysis/                  # Analysis & optimization
│   ├── goal_seeker.py         # Goal-seeking optimization
│   ├── sensitivity.py         # Sensitivity analysis
│   ├── monte_carlo.py         # Monte Carlo simulation
│   └── gbm_simulator.py       # GBM price simulator
├── risk/                      # Risk analysis tools
│   ├── flagger.py             # Risk flagging
│   └── scorer.py              # Risk scoring
├── valuation/                 # Valuation & deal analysis
│   └── breakeven.py           # Breakeven calculator
├── data/                      # Data handling
│   └── loader.py              # Data loading
├── export/                    # Export & reporting
│   ├── excel.py               # Excel export
│   ├── template_based_export.py  # Template-based export with charts
│   └── presentation_charts.py    # Chart generation for presentation sheets
├── excel_integration/         # Excel integration utilities
│   └── chart_generator.py     # Chart generation for analysis sheets
├── scripts/                   # Python scripts for Terminal execution
│   ├── run_deal_valuation_from_excel.py
│   ├── run_monte_carlo_from_excel.py
│   ├── run_sensitivity_from_excel.py
│   └── run_breakeven_from_excel.py
├── templates/                 # Master template
│   ├── create_master_template.py
│   └── master_template_with_interactive_modules.xlsx
├── tests/                     # Test scripts
├── examples/                  # Example scripts
└── docs/                      # Documentation
```

## 🔒 Data Privacy

**100% Local Processing** - All data stays on your machine:
- ❌ No cloud services
- ❌ No API calls
- ❌ No data transmission
- ❌ No internet connection required

See `docs/DATA_PRIVACY.md` for details.

## 📚 Documentation

- **Getting Started**: See `examples/basic_usage.py`
- **Project Structure**: See `docs/PROJECT_STRUCTURE.md`
- **Master Template Workflow**: See `docs/MASTER_TEMPLATE_SIMPLIFIED_WORKFLOW_PLAN.md`
- **GBM Implementation**: See `docs/GBM_IMPLEMENTATION.md`
- **Advanced Modules**: See `docs/ADVANCED_MODULES_PLAN.md`

## 🎓 Advanced Features

### GBM (Geometric Brownian Motion)

Use industry-standard GBM for sophisticated price volatility modeling:

```python
# Monte Carlo with GBM
mc_results = model.run_monte_carlo(
    simulations=5000,
    use_gbm=True,
    gbm_drift=0.03,        # Expected annual return
    gbm_volatility=0.15     # Annual volatility
)
```

### Risk Analysis

Automatic risk assessment with detailed flags:

```python
# Risk flags are automatically calculated after DCF
risk_summary = model.get_risk_summary()
print(risk_summary)

# Get detailed risk score
risk_score = model.calculate_risk_score()
print(f"Overall: {risk_score['overall_risk_score']}/100")
print(f"Financial: {risk_score['financial_risk']}/100")
print(f"Volume: {risk_score['volume_risk']}/100")
```

### Breakeven Analysis

Find breakeven points for negotiation:

```python
# Calculate all breakeven points
breakeven = model.calculate_breakeven(metric='all')
print(f"Breakeven Price: ${breakeven['breakeven_price']['breakeven_price']:,.2f}/ton")
print(f"Breakeven Volume: {breakeven['breakeven_volume']['breakeven_volume_multiplier']:.2%}")
```

## 🤝 Contributing

This is a template project designed for customization. Feel free to:
- Add new modules following the existing structure
- Extend existing calculators
- Improve data handling
- Add new export formats

## 📄 License

MIT License - see LICENSE file for details

## 👤 Author

Oscar Castellas-Cartwright

---

**Built for rapid carbon credit streaming deal analysis and investment decision support.**
