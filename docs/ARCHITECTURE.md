# Architecture Overview

## Modular Structure

The CarbonModelGenerator package is organized into focused modules, each handling a specific responsibility. This design makes it easy to enhance, test, and maintain individual components.

```
carbon_model_template/
├── __init__.py                 # Package initialization and exports
├── carbon_model_generator.py   # Main orchestrator class
├── data_loader.py              # Data ingestion and cleaning
├── dcf_calculator.py           # DCF financial calculations
├── irr_calculator.py           # IRR calculation engine
├── goal_seeker.py              # Optimization for target IRR
├── example_usage.py            # Usage examples
├── requirements.txt            # Dependencies
└── README.md                   # Documentation
```

## Module Responsibilities

### 1. DataLoader (`data_loader.py`)
**Purpose**: Handle all data ingestion and cleaning operations.

**Key Features**:
- Reads CSV and Excel files
- Intelligent column detection with flexible matching
- Header cleaning and standardization
- Data validation and error handling
- Time series standardization (ensures 20 years)

**When to Enhance**:
- Add support for new file formats (JSON, Parquet, etc.)
- Improve column detection algorithms
- Add data validation rules
- Support different time series lengths

### 2. DCFCalculator (`dcf_calculator.py`)
**Purpose**: Perform all Discounted Cash Flow calculations.

**Key Features**:
- Calculates share of credits
- Revenue calculations
- Investment cash flow modeling
- NPV calculation
- Present value discounting
- Cumulative metrics

**When to Enhance**:
- Add fee calculations
- Support different investment deployment schedules
- Add sensitivity analysis
- Support multiple revenue streams

### 3. IRRCalculator (`irr_calculator.py`)
**Purpose**: Calculate Internal Rate of Return with robust fallback strategies.

**Key Features**:
- Primary method: Brent's algorithm (brentq)
- Fallback method: fsolve
- Handles edge cases gracefully
- Configurable tolerance

**When to Enhance**:
- Add additional IRR calculation methods
- Improve convergence for edge cases
- Add IRR validation and diagnostics
- Support modified IRR (MIRR)

### 4. GoalSeeker (`goal_seeker.py`)
**Purpose**: Optimize streaming percentage to achieve target IRR.

**Key Features**:
- Uses optimization to find optimal streaming percentage
- Validates feasibility before optimization
- Error handling for unachievable targets
- Configurable tolerance

**When to Enhance**:
- Add alternative optimization algorithms
- Support multi-objective optimization
- Add constraint handling
- Support optimization for other parameters

### 5. CarbonModelGenerator (`carbon_model_generator.py`)
**Purpose**: Main orchestrator class that coordinates all modules.

**Key Features**:
- Provides simple, unified API
- Manages module lifecycle
- Stores results and state
- Maintains backward compatibility

**When to Enhance**:
- Add batch processing capabilities
- Add result export functionality
- Add visualization methods
- Add reporting features

## Data Flow

```
User Input (CSV/Excel)
    ↓
DataLoader
    ↓
Cleaned DataFrame
    ↓
CarbonModelGenerator
    ↓
DCFCalculator → IRRCalculator
    ↓
Results (NPV, IRR, Cash Flows)
    ↓
GoalSeeker (if needed)
    ↓
Optimal Streaming Percentage
```

## Benefits of Modular Design

1. **Easy Enhancement**: Each module can be improved independently
2. **Testability**: Modules can be tested in isolation
3. **Reusability**: Modules can be used independently or together
4. **Maintainability**: Clear separation of concerns
5. **Extensibility**: Easy to add new features without breaking existing code

## Example: Enhancing a Module

### Scenario: Add support for JSON data files

**File to modify**: `data_loader.py`

**Changes needed**:
1. Add JSON reading in `load_file()` method
2. Add JSON-specific parsing if needed
3. No changes required in other modules

### Scenario: Add fee calculations to DCF

**File to modify**: `dcf_calculator.py`

**Changes needed**:
1. Add fee calculation method
2. Update `calculate_net_cash_flow()` to include fees
3. No changes required in other modules

### Scenario: Add new IRR calculation method

**File to modify**: `irr_calculator.py`

**Changes needed**:
1. Add new calculation method (e.g., `calculate_irr_newton()`)
2. Add to fallback chain in `calculate_irr()`
3. No changes required in other modules

## Testing Strategy

Each module can be tested independently:

```python
# Test DataLoader
from data_loader import DataLoader
loader = DataLoader()
data = loader.load_data("test_data.csv")
assert len(data) == 20

# Test DCFCalculator
from dcf_calculator import DCFCalculator
calc = DCFCalculator(wacc=0.08, ...)
results = calc.run_dcf(data, 0.48)
assert 'npv' in results

# Test IRRCalculator
from irr_calculator import IRRCalculator
irr_calc = IRRCalculator()
irr = irr_calc.calculate_irr(cash_flows)
assert not np.isnan(irr)
```

## Future Enhancements

Potential enhancements that are now easier with the modular structure:

1. **DataLoader**: Support for API data sources, database connections
2. **DCFCalculator**: Monte Carlo simulations, scenario analysis
3. **IRRCalculator**: Multiple IRR calculation methods, diagnostics
4. **GoalSeeker**: Multi-parameter optimization, constraint handling
5. **CarbonModelGenerator**: Batch processing, result visualization, reporting

