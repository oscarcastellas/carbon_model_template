# Carbon Model Template

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Code Style](https://img.shields.io/badge/Code%20Style-PEP8-orange.svg)](https://www.python.org/dev/peps/pep-0008/)

A comprehensive, production-ready Python framework for rapid DCF (Discounted Cash Flow) and Carbon Streaming financial analysis. This tool prioritizes robust data handling, automated financial calculations, and **fully formula-based Excel outputs** for complete transparency and auditability.

## 🎯 Project Highlights

- **Formula-Based Excel Output**: Every calculation is an Excel formula, ensuring full transparency and auditability for external review
- **Monte Carlo Simulation**: Dual-variable stochastic modeling with histogram visualizations for risk assessment
- **Robust Data Handling**: Intelligently processes messy, unstructured Excel/CSV files with automatic format detection
- **Modular Architecture**: Clean separation of concerns for easy maintenance and extension
- **Production Ready**: Comprehensive error handling, documentation, and testing

## 🛠️ Technologies Used

- **Python 3.8+** - Core language
- **pandas** - Data manipulation and analysis
- **numpy** - Numerical computations
- **scipy** - Scientific computing (optimization, IRR calculation)
- **openpyxl** - Excel file reading
- **xlsxwriter** - Excel file writing with formulas and charts

## ✨ Key Features

- ✅ **DCF Analysis**: NPV, IRR, cash flow calculations
- ✅ **Goal-Seeking Optimization**: Find optimal streaming percentage for target IRR
- ✅ **Sensitivity Analysis**: 2D sensitivity tables for risk assessment
- ✅ **Monte Carlo Simulation**: 5,000+ simulations with statistical analysis
- ✅ **Payback Period Calculation**: Simple and discounted payback
- ✅ **Excel Export**: Multi-sheet output with all formulas and histogram charts
- ✅ **Robust Data Loading**: Handles transposed formats, various column names, messy data
- ✅ **Assumption Extraction**: Automatically extracts assumptions from Excel files

## 🚀 Quick Start

### Installation

```bash
# Clone the repository
git clone https://github.com/yourusername/carbon_model_template.git
cd carbon_model_template

# Install dependencies
pip install -r requirements.txt
```

### Basic Usage

```python
from carbon_model_template import CarbonModelGenerator

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

# Monte Carlo simulation
mc_results = model.run_monte_carlo(simulations=5000)

# Export to Excel (all formulas!)
model.export_model_to_excel("results.xlsx")
```

### Run Tests

```bash
# Full test with Monte Carlo (takes 2-5 minutes)
python3 test_excel.py

# Quick test without Monte Carlo
python3 quick_test.py
```

## 📊 Output

The tool generates a comprehensive Excel file with 5 sheets:

1. **Inputs & Assumptions** - All model inputs (highlighted for easy modification)
2. **Valuation Schedule** - 20-year detailed cash flow table with formulas
3. **Summary & Results** - Key financial metrics (NPV, IRR, Payback) as formulas
4. **Sensitivity Analysis** - 2D IRR sensitivity table
5. **Monte Carlo Results** - Full simulation results with histogram charts

**All calculations are Excel formulas** - no hardcoded values. Change inputs and see results update automatically!

## 📁 Project Structure

```
carbon_model_template/
├── carbon_model_generator.py    # Main orchestrator class
├── data/
│   └── data_loader.py           # Robust data ingestion
├── calculators/
│   ├── dcf_calculator.py        # DCF calculations
│   ├── irr_calculator.py        # IRR with fallback strategies
│   ├── goal_seeker.py           # Optimization
│   ├── sensitivity_analyzer.py  # Sensitivity analysis
│   ├── payback_calculator.py   # Payback period
│   └── monte_carlo.py          # Monte Carlo simulation
├── reporting/
│   └── excel_exporter.py       # Excel export with formulas
├── test_excel.py               # Full test script
├── quick_test.py               # Quick test
└── requirements.txt            # Dependencies
```

## 📚 Documentation

- **[HOW_TO_USE.md](HOW_TO_USE.md)** - Complete usage guide with examples
- **[EXCEL_FORMULA_GUIDE.md](EXCEL_FORMULA_GUIDE.md)** - Detailed Excel formula documentation
- **[MONTE_CARLO_GUIDE.md](MONTE_CARLO_GUIDE.md)** - Monte Carlo simulation guide
- **[ARCHITECTURE.md](ARCHITECTURE.md)** - System architecture and design
- **[TEMPLATE_STRUCTURE.md](TEMPLATE_STRUCTURE.md)** - Customization guide

## 🎓 Technical Achievements

### Formula-Based Excel Generation
- Implemented comprehensive Excel formula generation using `xlsxwriter`
- All financial calculations are transparent Excel formulas
- Enables full auditability for external reviewers
- Supports complex formulas: NPV, IRR, cumulative calculations, conditional logic

### Advanced Data Processing
- Intelligent column detection with flexible pattern matching
- Automatic handling of transposed data formats
- Robust error handling for messy, unstructured data
- Assumption extraction from multiple Excel sheet formats

### Monte Carlo Simulation
- Dual-variable stochastic modeling (price + volume volatility)
- Uses original price forecasts as base (not replacement)
- Two variation methods: growth-rate-based and percentage-based
- Statistical analysis with percentiles (P10, P90) and standard deviations
- Histogram chart generation for visualization

### Modular Architecture
- Clean separation of concerns
- Each module has single responsibility
- Easy to extend and maintain
- Comprehensive error handling throughout

## 📝 Input Data Format

Your Excel/CSV file should contain:
- **Years 1-20**: Time series data
- **Carbon Credits Issued (Gross)**: Annual tonnage
- **Project Implementation Costs**: Annual CAPEX (USD)
- **Base Carbon Price**: Price per ton (USD) - your price forecast

The data loader automatically handles:
- Various column name formats
- Transposed layouts (labels in rows, years in columns)
- Multiple Excel sheets
- Messy/unstructured data

## 🎯 Use Cases

- **Carbon Credit Streaming Analysis**: Evaluate streaming agreements and terms
- **DCF Modeling**: Comprehensive discounted cash flow analysis
- **Risk Assessment**: Monte Carlo simulation for probabilistic outcomes
- **Sensitivity Analysis**: Understand impact of key variables
- **Goal-Seeking**: Find optimal terms for target returns

## 🔧 Requirements

- Python 3.8+
- pandas >= 1.5.0
- numpy >= 1.23.0
- scipy >= 1.9.0
- openpyxl >= 3.0.0
- xlsxwriter >= 3.0.0

## 📖 Example Scripts

- `test_excel.py` - Full test with Monte Carlo simulation
- `quick_test.py` - Quick test without Monte Carlo
- `example_usage.py` - Usage examples
- `example_assumptions.py` - Assumption handling examples

## 🔍 Key Differentiators

1. **Formula-Based Excel**: Every calculation is transparent and auditable
2. **Robust Data Handling**: Works with messy, unstructured data automatically
3. **Monte Carlo with Your Forecasts**: Uses your price forecasts as the base, not replacement
4. **Production Ready**: Comprehensive error handling and documentation
5. **Modular Design**: Easy to extend and enhance

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🤝 Contributing

This is a portfolio project. For questions or suggestions, please open an issue.

## 📧 Contact

For questions or feedback, please open an issue on GitHub.

---

**Ready to use!** Start with `python3 test_excel.py` to see everything in action.
