# Template Structure Guide

This document explains the template structure and how to customize it for your use case.

## 📁 Project Structure

```
carbon_model_template/
├── carbon_model_generator.py    # Main orchestrator class
│
├── data/                         # Data handling modules
│   ├── __init__.py
│   └── data_loader.py           # Robust data ingestion
│
├── calculators/                  # Financial calculation modules
│   ├── __init__.py
│   ├── dcf_calculator.py        # DCF calculations
│   ├── irr_calculator.py        # IRR calculation
│   ├── goal_seeker.py           # Goal-seeking optimization
│   ├── sensitivity_analyzer.py  # Sensitivity analysis
│   ├── payback_calculator.py   # Payback period
│   └── monte_carlo.py          # Monte Carlo simulation
│
├── reporting/                    # Output and reporting modules
│   ├── __init__.py
│   └── excel_exporter.py       # Excel export with formulas
│
├── test_excel.py                # Full test script (recommended)
├── quick_test.py                # Quick test (no Monte Carlo)
├── example_usage.py             # Usage examples
├── example_assumptions.py       # Assumption handling examples
│
├── requirements.txt             # Python dependencies
├── setup.py                     # Package setup (optional)
│
├── README.md                    # Main entry point
├── HOW_TO_USE.md               # Complete usage guide
├── ARCHITECTURE.md              # System architecture
├── EXCEL_FORMULA_GUIDE.md      # Excel output documentation
├── MONTE_CARLO_GUIDE.md        # Monte Carlo guide
│
└── Analyst_Model_Test_OCC.xlsx  # Example data file
```

## 🎯 Core Components

### Main Class: `CarbonModelGenerator`

The main orchestrator that coordinates all modules. Use this for most use cases.

**Key Methods:**
- `load_data()` - Load and clean data
- `run_dcf()` - Run DCF analysis
- `find_target_irr_stream()` - Goal-seeking
- `run_sensitivity_table()` - Sensitivity analysis
- `run_monte_carlo()` - Monte Carlo simulation
- `export_model_to_excel()` - Export to Excel

### Data Module: `data/data_loader.py`

Handles all data ingestion and cleaning.

**Features:**
- CSV and Excel file support
- Transposed format detection
- Automatic column mapping
- Assumption extraction

### Calculators Module: `calculators/`

Individual calculation modules for specific tasks.

**Modules:**
- `dcf_calculator.py` - DCF calculations
- `irr_calculator.py` - IRR with fallbacks
- `goal_seeker.py` - Optimization
- `sensitivity_analyzer.py` - Sensitivity tables
- `payback_calculator.py` - Payback periods
- `monte_carlo.py` - Stochastic simulation

### Reporting Module: `reporting/excel_exporter.py`

Creates comprehensive Excel output with formulas.

**Features:**
- Formula-based calculations
- Multi-sheet output
- Histogram charts
- Professional formatting

## 🔧 Customization Guide

### 1. Replace Example Data

Replace `Analyst_Model_Test_OCC.xlsx` with your data file.

**Required columns:**
- Years 1-20
- Carbon Credits Issued (Gross)
- Project Implementation Costs
- Base Carbon Price

### 2. Adjust Assumptions

Modify assumptions in your script:

```python
model = CarbonModelGenerator(
    wacc=0.08,                    # Your discount rate
    rubicon_investment_total=20_000_000,  # Your investment
    investment_tenor=5,           # Your deployment period
    streaming_percentage_initial=0.48,  # Your initial streaming %
    
    # Monte Carlo assumptions
    price_growth_base=0.03,       # Your price growth assumption
    price_growth_std_dev=0.02,   # Your price volatility
    volume_multiplier_base=1.0,   # Your volume assumption
    volume_std_dev=0.15          # Your volume volatility
)
```

### 3. Customize Test Scripts

Copy `test_excel.py` and modify for your workflow:

```python
# Your custom script
from carbon_model_template import CarbonModelGenerator

model = CarbonModelGenerator(
    # Your assumptions
)

model.load_data("your_data.xlsx")
# ... your analysis workflow
model.export_model_to_excel("your_output.xlsx")
```

### 4. Extend Modules

Add new functionality by extending existing modules:

```python
# Example: Add new calculator
from calculators.dcf_calculator import DCFCalculator

class CustomCalculator(DCFCalculator):
    def calculate_custom_metric(self, data):
        # Your custom calculation
        pass
```

## 📝 File Naming Conventions

- **Main class**: `carbon_model_generator.py`
- **Modules**: Lowercase with underscores (`data_loader.py`)
- **Test scripts**: `test_*.py` or descriptive names
- **Documentation**: UPPERCASE with underscores (`HOW_TO_USE.md`)

## 🚀 Deployment Checklist

- [ ] Replace example data file
- [ ] Update assumptions in test scripts
- [ ] Test with your data
- [ ] Review Excel output
- [ ] Customize for your workflow
- [ ] Document any customizations

## 📚 Documentation Files

- **README.md**: Main entry point and overview
- **HOW_TO_USE.md**: Complete usage guide
- **ARCHITECTURE.md**: System design and architecture
- **EXCEL_FORMULA_GUIDE.md**: Excel output details
- **MONTE_CARLO_GUIDE.md**: Monte Carlo simulation guide
- **TEMPLATE_STRUCTURE.md**: This file

## 🔍 Key Design Principles

1. **Modularity**: Each module has a single responsibility
2. **Transparency**: All calculations are formulas in Excel
3. **Robustness**: Handles messy data gracefully
4. **Extensibility**: Easy to add new features
5. **Documentation**: Comprehensive guides for all features

## 💡 Tips for Customization

1. **Start Simple**: Use the main class first, then dive into modules
2. **Test Incrementally**: Test each feature as you customize
3. **Keep Formulas**: Maintain formula-based Excel output
4. **Document Changes**: Note any customizations you make
5. **Version Control**: Use git to track your changes

---

**Ready to customize!** Start with the test scripts and modify as needed.

