# Implementation Plan: Buffer Pool Risk Calculator

## 🎯 Overview

**Purpose**: Calculate buffer pool percentage based on project risks and automatically reduce sellable credits by buffer amount. This makes the model more realistic by accounting for the industry-standard practice of holding credits in reserve.

**Status**: Planned but not implemented  
**Priority**: HIGH  
**Development Time**: 2-3 days

---

## 📋 Current State Analysis

### **What We Have:**
- ✅ `DCFCalculator` that calculates revenue from credits
- ✅ `RiskFlagger` that identifies risks
- ✅ `RiskScoreCalculator` that scores risks
- ✅ Revenue calculation: `Credits × Price × Streaming %`

### **What We Need:**
- ❌ Buffer pool percentage calculation
- ❌ Risk-based buffer allocation
- ❌ Integration into DCF flow (reduce sellable credits)
- ❌ Excel export with buffer pool details

---

## 🏗️ Technical Architecture

### **Module Structure**

```
risk/
├── __init__.py
├── flagger.py          # Existing
├── scorer.py           # Existing
└── buffer_pool.py      # NEW: Buffer pool calculator
```

### **Class Design**

```python
class BufferPoolCalculator:
    """
    Calculates buffer pool percentage and adjusts sellable credits.
    
    Buffer pools are standard in carbon projects to account for:
    - Reversal risks (credits invalidated)
    - Verification failures
    - Project underperformance
    - Market risks
    """
    
    def __init__(
        self,
        base_buffer_percentage: float = 0.20,  # 20% default
        risk_multiplier: float = 1.0
    ):
        """
        Initialize the Buffer Pool Calculator.
        
        Parameters:
        -----------
        base_buffer_percentage : float
            Base buffer percentage (default: 20%)
        risk_multiplier : float
            Multiplier for risk-adjusted buffer (default: 1.0)
        """
        self.base_buffer_percentage = base_buffer_percentage
        self.risk_multiplier = risk_multiplier
    
    def calculate_buffer_percentage(
        self,
        risk_flags: Dict,
        risk_score: Dict,
        project_type: Optional[str] = None,
        verification_status: Optional[str] = None
    ) -> Dict:
        """
        Calculate buffer pool percentage based on risks.
        
        Parameters:
        -----------
        risk_flags : Dict
            Risk flags from RiskFlagger
        risk_score : Dict
            Risk score from RiskScoreCalculator
        project_type : str, optional
            Project type (e.g., 'forestry', 'agriculture', 'technology')
        verification_status : str, optional
            Verification status (e.g., 'verified', 'pending', 'unverified')
            
        Returns:
        --------
        dict
            Dictionary containing:
            - 'buffer_percentage': Calculated buffer percentage
            - 'base_buffer': Base buffer percentage
            - 'risk_adjustment': Risk-based adjustment
            - 'project_type_adjustment': Project type adjustment
            - 'verification_adjustment': Verification status adjustment
            - 'total_adjustment': Total adjustment factor
            - 'sellable_percentage': Percentage of credits that are sellable (1 - buffer)
        """
    
    def adjust_credits_for_buffer(
        self,
        credits: pd.Series,
        buffer_percentage: float
    ) -> pd.Series:
        """
        Reduce credits by buffer pool amount.
        
        Parameters:
        -----------
        credits : pd.Series
            Gross credits (before buffer)
        buffer_percentage : float
            Buffer percentage (0.0 to 1.0)
            
        Returns:
        --------
        pd.Series
            Sellable credits (after buffer deduction)
        """
        sellable_percentage = 1.0 - buffer_percentage
        return credits * sellable_percentage
```

---

## 🔧 Implementation Details

### **Step 1: Buffer Percentage Calculation Logic**

```python
def calculate_buffer_percentage(
    self,
    risk_flags: Dict,
    risk_score: Dict,
    project_type: Optional[str] = None,
    verification_status: Optional[str] = None
) -> Dict:
    """
    Calculate buffer pool percentage based on multiple factors.
    
    Buffer calculation formula:
    buffer = base_buffer × (1 + risk_adjustment) × project_type_factor × verification_factor
    
    Where:
    - base_buffer: Default buffer (e.g., 20%)
    - risk_adjustment: Based on risk score (0-100 → 0% to +50% adjustment)
    - project_type_factor: Based on project type (forestry: 1.2x, tech: 0.8x)
    - verification_factor: Based on verification status (verified: 0.9x, unverified: 1.3x)
    """
    # Start with base buffer
    base_buffer = self.base_buffer_percentage
    
    # Risk adjustment: Higher risk = higher buffer
    overall_risk_score = risk_score.get('overall_risk_score', 50)  # 0-100
    risk_adjustment = (overall_risk_score / 100) * 0.5  # 0% to 50% adjustment
    risk_adjusted_buffer = base_buffer * (1 + risk_adjustment)
    
    # Project type adjustment
    project_type_factors = {
        'forestry': 1.2,      # Higher reversal risk
        'agriculture': 1.1,   # Moderate risk
        'technology': 0.8,    # Lower risk
        'blue_carbon': 1.3,   # Higher risk (newer)
        'soil_carbon': 1.0,   # Standard
        None: 1.0            # Default
    }
    project_type_factor = project_type_factors.get(project_type, 1.0)
    
    # Verification status adjustment
    verification_factors = {
        'verified': 0.9,      # Lower buffer (already verified)
        'pending': 1.0,        # Standard
        'unverified': 1.3,   # Higher buffer (not yet verified)
        None: 1.0            # Default
    }
    verification_factor = verification_factors.get(verification_status, 1.0)
    
    # Calculate final buffer
    final_buffer = risk_adjusted_buffer * project_type_factor * verification_factor
    
    # Cap buffer at reasonable limits (5% to 50%)
    final_buffer = max(0.05, min(0.50, final_buffer))
    
    # Calculate sellable percentage
    sellable_percentage = 1.0 - final_buffer
    
    return {
        'buffer_percentage': final_buffer,
        'base_buffer': base_buffer,
        'risk_adjustment': risk_adjustment,
        'risk_adjusted_buffer': risk_adjusted_buffer,
        'project_type_adjustment': project_type_factor,
        'verification_adjustment': verification_factor,
        'total_adjustment': project_type_factor * verification_factor,
        'sellable_percentage': sellable_percentage,
        'risk_score_used': overall_risk_score
    }
```

### **Step 2: Risk-Based Buffer Factors**

```python
def _calculate_risk_adjustment(
    self,
    risk_flags: Dict,
    risk_score: Dict
) -> float:
    """
    Calculate risk adjustment factor for buffer pool.
    
    Uses both risk flags and risk score to determine adjustment.
    """
    # Base adjustment from risk score (0-100 → 0% to 50%)
    overall_risk_score = risk_score.get('overall_risk_score', 50)
    base_adjustment = (overall_risk_score / 100) * 0.5
    
    # Additional adjustment from specific risk flags
    red_flags = risk_flags.get('red_flags', [])
    yellow_flags = risk_flags.get('yellow_flags', [])
    
    # Red flags increase buffer more
    red_flag_adjustment = len(red_flags) * 0.05  # +5% per red flag
    
    # Yellow flags increase buffer moderately
    yellow_flag_adjustment = len(yellow_flags) * 0.02  # +2% per yellow flag
    
    # Total adjustment
    total_adjustment = base_adjustment + red_flag_adjustment + yellow_flag_adjustment
    
    # Cap at 100% adjustment (double the base buffer)
    return min(1.0, total_adjustment)
```

### **Step 3: Integration into DCF Calculator**

Modify `core/dcf.py`:

```python
def calculate_share_of_credits(
    self,
    data: pd.DataFrame,
    streaming_percentage: float,
    buffer_percentage: Optional[float] = None
) -> pd.Series:
    """
    Calculate Rubicon's share of carbon credits (after buffer pool).
    
    Parameters:
    -----------
    data : pd.DataFrame
        Input data with 'carbon_credits_gross' column
    streaming_percentage : float
        Percentage of credits Rubicon receives (0.0 to 1.0)
    buffer_percentage : float, optional
        Buffer pool percentage (if None, no buffer applied)
        
    Returns:
    --------
    pd.Series
        Rubicon's share of credits (after buffer)
    """
    # Calculate gross credits
    gross_credits = data['carbon_credits_gross']
    
    # Apply buffer pool if provided
    if buffer_percentage is not None:
        sellable_percentage = 1.0 - buffer_percentage
        sellable_credits = gross_credits * sellable_percentage
    else:
        sellable_credits = gross_credits
    
    # Apply streaming percentage
    rubicon_share = sellable_credits * streaming_percentage
    
    return rubicon_share
```

---

## 🔗 Integration Points

### **1. Extend `CarbonModelGenerator`**

Add buffer pool calculation to `carbon_model_generator.py`:

```python
def run_dcf(
    self,
    streaming_percentage: Optional[float] = None,
    apply_buffer_pool: bool = True
) -> Dict:
    """
    Run DCF analysis with optional buffer pool adjustment.
    
    Parameters:
    -----------
    streaming_percentage : float, optional
        Streaming percentage (uses initial if not provided)
    apply_buffer_pool : bool
        Whether to apply buffer pool adjustment (default: True)
        
    Returns:
    --------
    dict
        DCF results with buffer pool information
    """
    if self.data is None:
        raise ValueError("Data not loaded. Please call load_data() first.")
    
    if streaming_percentage is None:
        streaming_percentage = self._streaming_percentage_initial
    
    # Calculate buffer pool if requested
    buffer_results = None
    buffer_percentage = None
    
    if apply_buffer_pool:
        # First run DCF to get risk flags/score
        initial_results = self.dcf_calculator.run_dcf(
            self.data, 
            streaming_percentage
        )
        
        # Calculate risk flags and score
        self.risk_flags = self.flag_risks()
        self.risk_score = self.calculate_risk_score()
        
        # Calculate buffer pool
        from .risk.buffer_pool import BufferPoolCalculator
        buffer_calculator = BufferPoolCalculator()
        buffer_results = buffer_calculator.calculate_buffer_percentage(
            risk_flags=self.risk_flags,
            risk_score=self.risk_score,
            project_type=None,  # Can be added as parameter
            verification_status=None  # Can be added as parameter
        )
        buffer_percentage = buffer_results['buffer_percentage']
    
    # Run DCF with buffer pool adjustment
    results = self.dcf_calculator.run_dcf(
        self.data,
        streaming_percentage,
        buffer_percentage=buffer_percentage
    )
    
    # Add buffer pool info to results
    if buffer_results:
        results['buffer_pool'] = buffer_results
        results['buffer_percentage'] = buffer_percentage
        results['sellable_percentage'] = buffer_results['sellable_percentage']
    
    return results
```

### **2. Excel Export Integration**

Add buffer pool section to `export/excel.py`:

```python
def _write_summary_results_sheet(
    self,
    workbook: xlsxwriter.Workbook,
    sheet: xlsxwriter.Worksheet,
    formats: Dict,
    valuation_schedule: pd.DataFrame,
    inputs_sheet: xlsxwriter.Worksheet,
    actual_irr: float,
    payback_period: Optional[float],
    monte_carlo_results: Optional[Dict],
    risk_flags: Optional[Dict],
    risk_score: Optional[Dict],
    breakeven_results: Optional[Dict],
    buffer_pool_results: Optional[Dict] = None  # NEW
) -> None:
    """
    Write Summary & Results sheet with buffer pool information.
    """
    # ... existing code ...
    
    # Add Buffer Pool section
    if buffer_pool_results:
        row = self._write_buffer_pool_section(
            sheet, formats, row, buffer_pool_results, inputs_sheet
        )

def _write_buffer_pool_section(
    self,
    sheet: xlsxwriter.Worksheet,
    formats: Dict,
    start_row: int,
    buffer_pool_results: Dict,
    inputs_sheet: xlsxwriter.Worksheet
) -> int:
    """
    Write Buffer Pool section to Summary sheet.
    
    Returns:
    --------
    int
        Next row after section
    """
    row = start_row
    
    # Section header
    sheet.write(row, 0, 'Buffer Pool Analysis', formats['section_header'])
    row += 1
    
    # Buffer percentage
    sheet.write(row, 0, 'Buffer Pool Percentage:', formats['label'])
    sheet.write(row, 1, buffer_pool_results['buffer_percentage'], formats['percentage'])
    row += 1
    
    # Sellable percentage
    sheet.write(row, 0, 'Sellable Credits Percentage:', formats['label'])
    sheet.write(row, 1, buffer_pool_results['sellable_percentage'], formats['percentage'])
    row += 1
    
    # Base buffer
    sheet.write(row, 0, 'Base Buffer:', formats['label'])
    sheet.write(row, 1, buffer_pool_results['base_buffer'], formats['percentage'])
    row += 1
    
    # Risk adjustment
    sheet.write(row, 0, 'Risk Adjustment:', formats['label'])
    sheet.write(row, 1, f"{buffer_pool_results['risk_adjustment']:.1%}", formats['percentage'])
    row += 1
    
    # Project type adjustment
    sheet.write(row, 0, 'Project Type Factor:', formats['label'])
    sheet.write(row, 1, buffer_pool_results['project_type_adjustment'], formats['number'])
    row += 1
    
    # Verification adjustment
    sheet.write(row, 0, 'Verification Factor:', formats['label'])
    sheet.write(row, 1, buffer_pool_results['verification_adjustment'], formats['number'])
    row += 1
    
    # Risk score used
    sheet.write(row, 0, 'Risk Score Used:', formats['label'])
    sheet.write(row, 1, buffer_pool_results['risk_score_used'], formats['number'])
    row += 1
    
    row += 1  # Empty row
    
    return row
```

---

## 📊 Development Steps

### **Day 1: Core Buffer Pool Logic**

1. **Create `risk/buffer_pool.py`**
   - Implement `BufferPoolCalculator` class
   - Implement `calculate_buffer_percentage()`
   - Implement `_calculate_risk_adjustment()`
   - Implement `adjust_credits_for_buffer()`

2. **Define Buffer Calculation Rules**
   - Base buffer: 20%
   - Risk adjustment: 0% to 50% based on risk score
   - Project type factors
   - Verification status factors

3. **Unit Tests**
   - Test buffer calculation with different risk scores
   - Test project type adjustments
   - Test verification status adjustments
   - Test edge cases

### **Day 2: DCF Integration**

1. **Modify `DCFCalculator`**
   - Add `buffer_percentage` parameter to `calculate_share_of_credits()`
   - Update `run_dcf()` to accept buffer percentage
   - Ensure buffer is applied before streaming percentage

2. **Extend `CarbonModelGenerator`**
   - Add `apply_buffer_pool` parameter to `run_dcf()`
   - Calculate buffer pool after initial risk assessment
   - Pass buffer percentage to DCF calculator
   - Store buffer results

3. **Integration Tests**
   - Test DCF with buffer pool
   - Verify credits are reduced correctly
   - Verify revenue is adjusted
   - Test with different risk scenarios

### **Day 3: Excel Export & Testing**

1. **Update Excel Exporter**
   - Add `_write_buffer_pool_section()` method
   - Integrate into Summary & Results sheet
   - Add buffer pool details to Inputs sheet

2. **Comprehensive Testing**
   - Test with multiple datasets
   - Test with different risk levels
   - Test Excel export
   - Verify formulas work

3. **Documentation**
   - Update README
   - Document buffer pool calculation methodology
   - Create usage examples

---

## 🧪 Testing Strategy

### **Unit Tests**

```python
def test_buffer_calculation_low_risk():
    """Test buffer calculation with low risk."""
    calculator = BufferPoolCalculator()
    risk_flags = {'red_flags': [], 'yellow_flags': []}
    risk_score = {'overall_risk_score': 20}  # Low risk
    
    results = calculator.calculate_buffer_percentage(
        risk_flags, risk_score
    )
    
    assert results['buffer_percentage'] < 0.25  # Should be low

def test_buffer_calculation_high_risk():
    """Test buffer calculation with high risk."""
    calculator = BufferPoolCalculator()
    risk_flags = {'red_flags': ['Low IRR', 'High Payback'], 'yellow_flags': []}
    risk_score = {'overall_risk_score': 80}  # High risk
    
    results = calculator.calculate_buffer_percentage(
        risk_flags, risk_score
    )
    
    assert results['buffer_percentage'] > 0.30  # Should be high

def test_project_type_adjustment():
    """Test project type adjustment."""
    calculator = BufferPoolCalculator()
    risk_flags = {'red_flags': []}
    risk_score = {'overall_risk_score': 50}
    
    # Forestry should have higher buffer
    forestry_results = calculator.calculate_buffer_percentage(
        risk_flags, risk_score, project_type='forestry'
    )
    
    # Technology should have lower buffer
    tech_results = calculator.calculate_buffer_percentage(
        risk_flags, risk_score, project_type='technology'
    )
    
    assert forestry_results['buffer_percentage'] > tech_results['buffer_percentage']
```

### **Integration Tests**

```python
def test_dcf_with_buffer_pool():
    """Test DCF calculation with buffer pool."""
    model = CarbonModelGenerator(...)
    model.load_data("test_data.xlsx")
    
    # Run DCF with buffer pool
    results = model.run_dcf(apply_buffer_pool=True)
    
    # Verify buffer pool is applied
    assert 'buffer_pool' in results
    assert 'buffer_percentage' in results
    assert results['buffer_percentage'] > 0
    
    # Verify credits are reduced
    # (Compare with results without buffer pool)
```

---

## 📝 Excel Sheet Design

### **Summary & Results Sheet - Buffer Pool Section**

**Layout:**
```
Row N: "Buffer Pool Analysis"
Row N+1: "Buffer Pool Percentage:" | [Value] | "%"
Row N+2: "Sellable Credits Percentage:" | [Value] | "%"
Row N+3: "Base Buffer:" | [Value] | "%"
Row N+4: "Risk Adjustment:" | [Value] | "%"
Row N+5: "Project Type Factor:" | [Value]
Row N+6: "Verification Factor:" | [Value]
Row N+7: "Risk Score Used:" | [Value]
```

**Inputs & Assumptions Sheet:**
- Add buffer pool parameters section
- Allow user to override buffer percentage if needed
- Show calculated buffer percentage

---

## 🚀 Success Criteria

1. ✅ Buffer pool percentage calculated based on risks
2. ✅ Credits reduced by buffer amount in DCF
3. ✅ Revenue calculations reflect buffer pool
4. ✅ Excel export includes buffer pool details
5. ✅ Risk-based adjustments work correctly
6. ✅ Project type adjustments work correctly
7. ✅ Verification status adjustments work correctly
8. ✅ Unit tests pass
9. ✅ Integration tests pass
10. ✅ Documentation complete

---

## 📋 Checklist

- [ ] Create `risk/buffer_pool.py`
- [ ] Implement `BufferPoolCalculator` class
- [ ] Implement buffer calculation logic
- [ ] Modify `DCFCalculator` to accept buffer percentage
- [ ] Extend `CarbonModelGenerator.run_dcf()` with buffer pool
- [ ] Add Excel export functionality
- [ ] Write unit tests
- [ ] Write integration tests
- [ ] Test with real data
- [ ] Update documentation

---

## 🔄 Next Steps After Implementation

1. **Scenario Analysis Generator** (next module)
2. **Deal Comparison Dashboard** (future enhancement)
3. **Portfolio Risk Aggregator** (future enhancement)

---

**Ready to start building?** This plan provides a complete roadmap for implementing the Buffer Pool Risk Calculator module.

