# Project Reorganization Plan

## Current Issues

1. **Root directory clutter**: Test scripts, examples, docs all mixed together
2. **calculators/ folder too large**: 9 modules mixing different concerns
3. **Unclear separation**: Core calculations vs. risk analysis vs. productivity tools
4. **Hard to scale**: Adding new modules makes it more confusing

## Proposed Structure

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
│   └── gbm_simulator.py       # NEW: GBM price simulator
│
├── risk/                      # Risk analysis tools
│   ├── __init__.py
│   ├── flagger.py             # Risk flagging
│   └── scorer.py              # Risk scoring
│
├── valuation/                 # Valuation & deal analysis
│   ├── __init__.py
│   ├── breakeven.py           # Breakeven calculator
│   └── deal_solver.py         # NEW: Deal valuation (future)
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
    ├── ADVANCED_MODULES_PLAN.md
    └── ...
```

## Benefits

1. **Clear separation**: Each folder has a single, clear purpose
2. **Scalable**: Easy to add new modules in the right place
3. **Intuitive**: Folder names describe their contents
4. **Professional**: Clean structure for collaboration
5. **Maintainable**: Easy to find and update code

## Migration Plan

1. Create new folder structure
2. Move files with updated imports
3. Update all import statements
4. Update __init__.py files
5. Test all functionality
6. Update documentation

