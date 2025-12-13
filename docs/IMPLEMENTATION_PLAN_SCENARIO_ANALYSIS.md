# Implementation Plan: Scenario Analysis Generator (Inside Excel)

## 🎯 Overview

**Purpose**: Auto-generate common scenarios (Base, Upside, Downside, Stress) directly in Excel with one-click scenario switching. This saves hours of manual scenario building and ensures consistent methodology across all deals.

**Status**: Not implemented  
**Priority**: MEDIUM-HIGH  
**Development Time**: 2 days

---

## 📋 Current State Analysis

### **What We Have:**
- ✅ `DCFCalculator` that runs single scenario
- ✅ Excel export with formulas
- ✅ Inputs & Assumptions sheet
- ✅ Valuation Schedule with formulas

### **What We Need:**
- ❌ Multiple scenario definitions (Base, Upside, Downside, Stress)
- ❌ Scenario switching mechanism in Excel
- ❌ Scenario comparison table
- ❌ Scenario-specific results sheets

---

## 🏗️ Technical Architecture

### **Module Structure**

```
analysis/
├── __init__.py
├── goal_seeker.py          # Existing
├── sensitivity.py          # Existing
├── monte_carlo.py          # Existing
└── scenario_generator.py  # NEW: Scenario generator
```

### **Class Design**

```python
class ScenarioGenerator:
    """
    Generates standard scenarios for carbon credit deal analysis.
    
    Scenarios:
    - Base Case: Current assumptions
    - Upside: Optimistic assumptions (price +20%, volume +10%)
    - Downside: Pessimistic assumptions (price -20%, volume -10%)
    - Stress: Worst-case assumptions (price -40%, volume -20%)
    """
    
    def __init__(
        self,
        base_assumptions: Dict
    ):
        """
        Initialize the Scenario Generator.
        
        Parameters:
        -----------
        base_assumptions : Dict
            Base case assumptions (from model)
        """
        self.base_assumptions = base_assumptions
    
    def generate_scenarios(
        self,
        custom_scenarios: Optional[Dict[str, Dict]] = None
    ) -> Dict[str, Dict]:
        """
        Generate all standard scenarios.
        
        Parameters:
        -----------
        custom_scenarios : Dict[str, Dict], optional
            Custom scenario definitions
            
        Returns:
        --------
        dict
            Dictionary of scenario name -> scenario assumptions
        """
    
    def get_scenario_assumptions(
        self,
        scenario_name: str
    ) -> Dict:
        """
        Get assumptions for a specific scenario.
        
        Parameters:
        -----------
        scenario_name : str
            Scenario name ('Base', 'Upside', 'Downside', 'Stress')
            
        Returns:
        --------
        dict
            Scenario assumptions
        """
```

---

## 🔧 Implementation Details

### **Step 1: Scenario Definitions**

```python
def generate_scenarios(
    self,
    custom_scenarios: Optional[Dict[str, Dict]] = None
) -> Dict[str, Dict]:
    """
    Generate all standard scenarios.
    
    Scenario definitions:
    - Base: 100% of all assumptions
    - Upside: Price +20%, Volume +10%, Lower WACC
    - Downside: Price -20%, Volume -10%, Higher WACC
    - Stress: Price -40%, Volume -20%, Higher WACC, Longer Payback
    """
    scenarios = {}
    
    # Base Case
    scenarios['Base'] = {
        'name': 'Base Case',
        'description': 'Current assumptions',
        'price_multiplier': 1.0,
        'volume_multiplier': 1.0,
        'wacc_adjustment': 0.0,
        'streaming_adjustment': 0.0,
        'assumptions': self.base_assumptions.copy()
    }
    
    # Upside Case
    scenarios['Upside'] = {
        'name': 'Upside Case',
        'description': 'Optimistic assumptions: Price +20%, Volume +10%',
        'price_multiplier': 1.20,
        'volume_multiplier': 1.10,
        'wacc_adjustment': -0.01,  # Lower WACC (better financing)
        'streaming_adjustment': 0.0,
        'assumptions': self._apply_scenario_multipliers(
            self.base_assumptions.copy(),
            price_multiplier=1.20,
            volume_multiplier=1.10,
            wacc_adjustment=-0.01
        )
    }
    
    # Downside Case
    scenarios['Downside'] = {
        'name': 'Downside Case',
        'description': 'Pessimistic assumptions: Price -20%, Volume -10%',
        'price_multiplier': 0.80,
        'volume_multiplier': 0.90,
        'wacc_adjustment': 0.01,  # Higher WACC (worse financing)
        'streaming_adjustment': 0.0,
        'assumptions': self._apply_scenario_multipliers(
            self.base_assumptions.copy(),
            price_multiplier=0.80,
            volume_multiplier=0.90,
            wacc_adjustment=0.01
        )
    }
    
    # Stress Case
    scenarios['Stress'] = {
        'name': 'Stress Case',
        'description': 'Worst-case assumptions: Price -40%, Volume -20%',
        'price_multiplier': 0.60,
        'volume_multiplier': 0.80,
        'wacc_adjustment': 0.02,  # Much higher WACC
        'streaming_adjustment': 0.0,
        'assumptions': self._apply_scenario_multipliers(
            self.base_assumptions.copy(),
            price_multiplier=0.60,
            volume_multiplier=0.80,
            wacc_adjustment=0.02
        )
    }
    
    # Add custom scenarios if provided
    if custom_scenarios:
        scenarios.update(custom_scenarios)
    
    return scenarios

def _apply_scenario_multipliers(
    self,
    assumptions: Dict,
    price_multiplier: float = 1.0,
    volume_multiplier: float = 1.0,
    wacc_adjustment: float = 0.0
) -> Dict:
    """
    Apply scenario multipliers to assumptions.
    
    Parameters:
    -----------
    assumptions : Dict
        Base assumptions
    price_multiplier : float
        Price multiplier (e.g., 1.20 for +20%)
    volume_multiplier : float
        Volume multiplier (e.g., 0.90 for -10%)
    wacc_adjustment : float
        WACC adjustment (e.g., 0.01 for +1%)
        
    Returns:
    --------
    dict
        Adjusted assumptions
    """
    adjusted = assumptions.copy()
    
    # Adjust WACC
    if 'wacc' in adjusted:
        adjusted['wacc'] = adjusted['wacc'] + wacc_adjustment
    
    # Note: Price and volume multipliers are applied in Excel formulas
    # We store them here for reference
    adjusted['_scenario_price_multiplier'] = price_multiplier
    adjusted['_scenario_volume_multiplier'] = volume_multiplier
    
    return adjusted
```

### **Step 2: Excel Scenario Switching Mechanism**

The key is to use Excel's **Data Validation** and **IF formulas** to switch between scenarios.

**Approach:**
1. Create a "Scenario Selector" cell in Inputs sheet
2. Use Data Validation dropdown (Base, Upside, Downside, Stress)
3. Use IF formulas in Valuation Schedule to apply scenario multipliers
4. All formulas automatically update when scenario changes

### **Step 3: Excel Formula Structure**

```excel
# In Valuation Schedule sheet:

# Price column (applies scenario multiplier)
=IF(Inputs!$B$2="Base", Base_Price,
 IF(Inputs!$B$2="Upside", Base_Price * 1.20,
 IF(Inputs!$B$2="Downside", Base_Price * 0.80,
 IF(Inputs!$B$2="Stress", Base_Price * 0.60, Base_Price))))

# Volume column (applies scenario multiplier)
=IF(Inputs!$B$2="Base", Credits_Gross,
 IF(Inputs!$B$2="Upside", Credits_Gross * 1.10,
 IF(Inputs!$B$2="Downside", Credits_Gross * 0.90,
 IF(Inputs!$B$2="Stress", Credits_Gross * 0.80, Credits_Gross))))

# WACC (applies scenario adjustment)
=IF(Inputs!$B$2="Base", WACC_Base,
 IF(Inputs!$B$2="Upside", WACC_Base - 0.01,
 IF(Inputs!$B$2="Downside", WACC_Base + 0.01,
 IF(Inputs!$B$2="Stress", WACC_Base + 0.02, WACC_Base))))
```

**Better approach using VLOOKUP:**
```excel
# Create scenario lookup table in hidden sheet or Inputs sheet
# Then use:
=Base_Price * VLOOKUP(Inputs!$B$2, Scenario_Table, 2, FALSE)
=Credits_Gross * VLOOKUP(Inputs!$B$2, Scenario_Table, 3, FALSE)
=WACC_Base + VLOOKUP(Inputs!$B$2, Scenario_Table, 4, FALSE)
```

---

## 🔗 Integration Points

### **1. Extend Excel Exporter**

Add scenario functionality to `export/excel.py`:

```python
def _write_inputs_sheet(
    self,
    workbook: xlsxwriter.Workbook,
    sheet: xlsxwriter.Worksheet,
    formats: Dict,
    assumptions: Dict,
    target_streaming_percentage: float,
    target_irr: float,
    scenarios: Optional[Dict[str, Dict]] = None  # NEW
) -> None:
    """
    Write Inputs & Assumptions sheet with scenario selector.
    """
    # ... existing code ...
    
    # Add Scenario Selector section
    if scenarios:
        row = self._write_scenario_selector(
            sheet, formats, row, scenarios
        )
    
    # Add Scenario Definitions table
    if scenarios:
        row = self._write_scenario_definitions(
            sheet, formats, row, scenarios
        )

def _write_scenario_selector(
    self,
    sheet: xlsxwriter.Worksheet,
    formats: Dict,
    start_row: int,
    scenarios: Dict[str, Dict]
) -> int:
    """
    Write Scenario Selector section.
    
    Creates:
    - Dropdown cell for scenario selection
    - Scenario definitions table
    """
    row = start_row
    
    # Section header
    sheet.write(row, 0, 'Scenario Analysis', formats['section_header'])
    row += 1
    
    # Scenario selector
    sheet.write(row, 0, 'Active Scenario:', formats['label'])
    
    # Create dropdown using data validation
    # Note: xlsxwriter doesn't support data validation directly
    # We'll create a helper table and use VLOOKUP
    scenario_cell = xlsxwriter.utility.xl_rowcol_to_cell(row, 1)
    sheet.write(row, 1, 'Base', formats['input_cell'])
    
    # Add note about changing scenario
    row += 1
    sheet.write(row, 0, 'Note: Change scenario in cell above', formats['note'])
    row += 1
    
    return row

def _write_scenario_definitions(
    self,
    sheet: xlsxwriter.Worksheet,
    formats: Dict,
    start_row: int,
    scenarios: Dict[str, Dict]
) -> int:
    """
    Write Scenario Definitions table.
    
    Creates lookup table for scenario multipliers.
    """
    row = start_row
    
    # Table header
    sheet.write(row, 0, 'Scenario', formats['table_header'])
    sheet.write(row, 1, 'Price Mult.', formats['table_header'])
    sheet.write(row, 2, 'Volume Mult.', formats['table_header'])
    sheet.write(row, 3, 'WACC Adj.', formats['table_header'])
    sheet.write(row, 4, 'Description', formats['table_header'])
    row += 1
    
    # Write scenario data
    for scenario_name, scenario_data in scenarios.items():
        sheet.write(row, 0, scenario_name, formats['text'])
        sheet.write(row, 1, scenario_data['price_multiplier'], formats['number'])
        sheet.write(row, 2, scenario_data['volume_multiplier'], formats['number'])
        sheet.write(row, 3, scenario_data['wacc_adjustment'], formats['number'])
        sheet.write(row, 4, scenario_data['description'], formats['text'])
        row += 1
    
    row += 1  # Empty row
    
    return row
```

### **2. Modify Valuation Schedule Formulas**

Update `_write_valuation_schedule_with_formulas()` to use scenario multipliers:

```python
def _write_valuation_schedule_with_formulas(
    self,
    workbook: xlsxwriter.Workbook,
    sheet: xlsxwriter.Worksheet,
    formats: Dict,
    valuation_schedule: pd.DataFrame,
    inputs_sheet: xlsxwriter.Worksheet,
    scenarios: Optional[Dict[str, Dict]] = None  # NEW
) -> None:
    """
    Write Valuation Schedule with scenario-aware formulas.
    """
    # ... existing code ...
    
    # When writing price column, use scenario multiplier
    if scenarios:
        # Price formula with scenario multiplier
        price_formula = f"={base_price_cell} * VLOOKUP({scenario_selector_cell}, Scenario_Table, 2, FALSE)"
    else:
        price_formula = f"={base_price_cell}"
    
    # When writing volume column, use scenario multiplier
    if scenarios:
        volume_formula = f"={base_credits_cell} * VLOOKUP({scenario_selector_cell}, Scenario_Table, 3, FALSE)"
    else:
        volume_formula = f"={base_credits_cell}"
    
    # When writing WACC, use scenario adjustment
    if scenarios:
        wacc_formula = f"={base_wacc_cell} + VLOOKUP({scenario_selector_cell}, Scenario_Table, 4, FALSE)"
    else:
        wacc_formula = f"={base_wacc_cell}"
```

### **3. Add Scenario Comparison Sheet**

```python
def _write_scenario_comparison_sheet(
    self,
    workbook: xlsxwriter.Workbook,
    sheet: xlsxwriter.Worksheet,
    formats: Dict,
    inputs_sheet: xlsxwriter.Worksheet,
    valuation_sheet: xlsxwriter.Worksheet,
    scenarios: Dict[str, Dict]
) -> None:
    """
    Write Scenario Comparison sheet.
    
    Shows key metrics for each scenario side-by-side.
    """
    # Header
    sheet.write(0, 0, 'Scenario Comparison', formats['header'])
    
    # Table header
    row = 2
    sheet.write(row, 0, 'Metric', formats['table_header'])
    col = 1
    for scenario_name in scenarios.keys():
        sheet.write(row, col, scenario_name, formats['table_header'])
        col += 1
    row += 1
    
    # Key metrics for each scenario
    metrics = [
        ('NPV', 'Summary & Results!B3'),
        ('IRR', 'Summary & Results!B5'),
        ('Payback Period', 'Summary & Results!B7'),
        ('Total Revenue', 'Valuation Schedule!SUM(Revenue_Column)'),
        # ... more metrics
    ]
    
    for metric_name, formula_base in metrics:
        sheet.write(row, 0, metric_name, formats['label'])
        col = 1
        for scenario_name in scenarios.keys():
            # Formula that switches scenario and recalculates
            # This is complex - may need to use INDIRECT or helper cells
            sheet.write(row, col, f"={formula_base}", formats['number'])
            col += 1
        row += 1
```

---

## 📊 Development Steps

### **Day 1: Core Scenario Generator**

1. **Create `analysis/scenario_generator.py`**
   - Implement `ScenarioGenerator` class
   - Implement `generate_scenarios()`
   - Implement `_apply_scenario_multipliers()`
   - Define standard scenarios (Base, Upside, Downside, Stress)

2. **Define Scenario Rules**
   - Base: 100% of assumptions
   - Upside: Price +20%, Volume +10%, WACC -1%
   - Downside: Price -20%, Volume -10%, WACC +1%
   - Stress: Price -40%, Volume -20%, WACC +2%

3. **Unit Tests**
   - Test scenario generation
   - Test multiplier application
   - Test custom scenarios

### **Day 2: Excel Integration**

1. **Extend Excel Exporter**
   - Add scenario selector to Inputs sheet
   - Add scenario definitions table
   - Modify Valuation Schedule formulas to use scenario multipliers
   - Add Scenario Comparison sheet

2. **Implement Scenario Switching**
   - Use VLOOKUP for scenario multipliers
   - Ensure all formulas update when scenario changes
   - Test scenario switching in Excel

3. **Testing & Documentation**
   - Test with real data
   - Verify formulas work correctly
   - Create usage guide
   - Update documentation

---

## 🧪 Testing Strategy

### **Unit Tests**

```python
def test_scenario_generation():
    """Test scenario generation."""
    base_assumptions = {
        'wacc': 0.08,
        'base_carbon_price': 50.0
    }
    generator = ScenarioGenerator(base_assumptions)
    scenarios = generator.generate_scenarios()
    
    assert 'Base' in scenarios
    assert 'Upside' in scenarios
    assert 'Downside' in scenarios
    assert 'Stress' in scenarios
    
    # Verify multipliers
    assert scenarios['Upside']['price_multiplier'] == 1.20
    assert scenarios['Downside']['price_multiplier'] == 0.80

def test_scenario_assumptions():
    """Test scenario assumption application."""
    generator = ScenarioGenerator(base_assumptions)
    upside = generator.get_scenario_assumptions('Upside')
    
    # Verify WACC adjustment
    assert upside['assumptions']['wacc'] == base_assumptions['wacc'] - 0.01
```

### **Excel Integration Tests**

```python
def test_excel_scenario_switching():
    """Test scenario switching in Excel."""
    model = CarbonModelGenerator(...)
    model.load_data("test_data.xlsx")
    model.run_dcf()
    
    # Generate scenarios
    from .analysis.scenario_generator import ScenarioGenerator
    generator = ScenarioGenerator(model.get_assumptions())
    scenarios = generator.generate_scenarios()
    
    # Export with scenarios
    model.export_model_to_excel(
        "test_output.xlsx",
        scenarios=scenarios
    )
    
    # Verify Excel file
    assert os.path.exists("test_output.xlsx")
    # Verify scenario selector exists
    # Verify formulas use scenario multipliers
```

---

## 📝 Excel Sheet Design

### **Inputs & Assumptions Sheet - Scenario Section**

**Layout:**
```
Row N: "Scenario Analysis"
Row N+1: "Active Scenario:" | [Dropdown: Base/Upside/Downside/Stress]
Row N+2: Empty
Row N+3: "Scenario Definitions"
Row N+4: "Scenario" | "Price Mult." | "Volume Mult." | "WACC Adj." | "Description"
Row N+5: "Base" | 1.00 | 1.00 | 0.00 | "Current assumptions"
Row N+6: "Upside" | 1.20 | 1.10 | -0.01 | "Optimistic assumptions"
Row N+7: "Downside" | 0.80 | 0.90 | 0.01 | "Pessimistic assumptions"
Row N+8: "Stress" | 0.60 | 0.80 | 0.02 | "Worst-case assumptions"
```

### **Valuation Schedule Sheet - Scenario-Aware Formulas**

**Price Column:**
```excel
=Base_Price * VLOOKUP(Inputs!$B$N, Scenario_Table, 2, FALSE)
```

**Volume Column:**
```excel
=Credits_Gross * VLOOKUP(Inputs!$B$N, Scenario_Table, 3, FALSE)
```

**WACC:**
```excel
=Base_WACC + VLOOKUP(Inputs!$B$N, Scenario_Table, 4, FALSE)
```

### **Scenario Comparison Sheet (NEW)**

**Layout:**
```
Row 1: "Scenario Comparison"
Row 2: Empty
Row 3: "Metric" | "Base" | "Upside" | "Downside" | "Stress"
Row 4: "NPV" | [Formula] | [Formula] | [Formula] | [Formula]
Row 5: "IRR" | [Formula] | [Formula] | [Formula] | [Formula]
Row 6: "Payback Period" | [Formula] | [Formula] | [Formula] | [Formula]
...
```

**Note:** Scenario comparison formulas are complex because they need to switch scenarios. Consider using helper cells or a more sophisticated approach.

---

## 🚀 Success Criteria

1. ✅ Can generate standard scenarios (Base, Upside, Downside, Stress)
2. ✅ Scenario selector works in Excel
3. ✅ Valuation Schedule formulas use scenario multipliers
4. ✅ All calculations update when scenario changes
5. ✅ Scenario Comparison sheet shows all scenarios
6. ✅ Custom scenarios can be added
7. ✅ Unit tests pass
8. ✅ Excel integration tests pass
9. ✅ Documentation complete

---

## 📋 Checklist

- [ ] Create `analysis/scenario_generator.py`
- [ ] Implement `ScenarioGenerator` class
- [ ] Define standard scenarios
- [ ] Extend Excel exporter with scenario selector
- [ ] Modify Valuation Schedule formulas
- [ ] Add Scenario Comparison sheet
- [ ] Write unit tests
- [ ] Write Excel integration tests
- [ ] Test scenario switching
- [ ] Create usage guide
- [ ] Update documentation

---

## 💡 Alternative Approach: Multiple Sheets

Instead of scenario switching in one sheet, we could create **separate sheets for each scenario**:

- "Valuation Schedule - Base"
- "Valuation Schedule - Upside"
- "Valuation Schedule - Downside"
- "Valuation Schedule - Stress"

**Pros:**
- Simpler formulas (no IF/VLOOKUP)
- Can see all scenarios at once
- Easier to understand

**Cons:**
- More sheets to manage
- Larger file size
- More complex Excel structure

**Recommendation:** Start with scenario switching in one sheet (simpler), then consider multiple sheets if users prefer.

---

## 🔄 Next Steps After Implementation

1. **Deal Comparison Dashboard** (future enhancement)
2. **Portfolio Risk Aggregator** (future enhancement)
3. **Investment Committee Report Generator** (future enhancement)

---

**Ready to start building?** This plan provides a complete roadmap for implementing the Scenario Analysis Generator module inside Excel.

