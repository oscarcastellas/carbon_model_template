# Project Structure

## 📁 New Folder Organization

The project has been reorganized for clarity and scalability:

```
carbon_model_template/
├── carbon_model_generator.py  # Main orchestrator class
├── __init__.py
├── README.md
├── requirements.txt
├── setup.py
│
├── core/                      # Core financial calculations
│   ├── __init__.py
│   ├── dcf.py                 # DCF calculations
│   ├── irr.py                 # IRR calculations
│   └── payback.py             # Payback period
│
├── analysis/                  # Analysis & optimization
│   ├── __init__.py
│   ├── goal_seeker.py         # Goal-seeking optimization
│   ├── sensitivity.py         # Sensitivity analysis
│   ├── monte_carlo.py         # Monte Carlo simulation
│   └── gbm_simulator.py       # GBM price simulator (NEW)
│
├── risk/                      # Risk analysis tools
│   ├── __init__.py
│   ├── flagger.py             # Risk flagging
│   └── scorer.py              # Risk scoring
│
├── valuation/                 # Valuation & deal analysis
│   ├── __init__.py
│   ├── breakeven.py           # Breakeven calculator
│   └── deal_solver.py         # Deal valuation (future)
│
├── data/                      # Data handling
│   ├── __init__.py
│   └── loader.py              # Data loading
│
├── export/                    # Export & reporting
│   ├── __init__.py
│   └── excel.py               # Excel export
│
├── tests/                     # Test scripts
│   ├── test_basic.py
│   ├── test_excel.py
│   └── test_productivity.py
│
├── examples/                  # Example scripts
│   ├── basic_usage.py
│   └── assumptions.py
│
└── docs/                      # Documentation
    ├── ARCHITECTURE.md
    ├── HOW_TO_USE.md
    └── ...
```

## 🎯 Benefits

1. **Clear Separation**: Each folder has a single, clear purpose
2. **Scalable**: Easy to add new modules in the right place
3. **Intuitive**: Folder names describe their contents
4. **Professional**: Clean structure for collaboration
5. **Maintainable**: Easy to find and update code

## 📦 Module Categories

- **core/**: Fundamental financial calculations (DCF, IRR, Payback)
- **analysis/**: Advanced analysis tools (Monte Carlo, Sensitivity, GBM)
- **risk/**: Risk assessment and scoring
- **valuation/**: Deal valuation and breakeven analysis
- **data/**: Data loading and processing
- **export/**: Output generation (Excel, reports)

## 🔄 Migration Notes

All imports have been updated to use the new structure. The old `calculators/`, `data/`, and `reporting/` folders are deprecated but kept for backward compatibility during transition.

