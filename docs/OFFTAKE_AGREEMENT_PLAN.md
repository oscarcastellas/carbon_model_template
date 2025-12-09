# Offtake Agreement Integration Module - Development Plan

## 🎯 Objective

Create a module that integrates **Rubicon's specific offtake agreements** and **offtake contracts** as variables in the carbon model, allowing for realistic contract-based revenue modeling instead of simple streaming percentages.

---

## 📋 Current State vs. Desired State

### **Current Model**
- Uses simple **streaming percentage** (e.g., 48% of credits)
- Assumes **spot market pricing** for all credits
- No contract structure or terms
- Revenue = Credits × Price × Streaming %

### **Desired Model**
- **Multiple offtake contracts** with different terms
- **Contract-specific pricing** (fixed, floor, cap, indexed)
- **Volume commitments** (minimum/maximum)
- **Contract duration** and **renewal terms**
- **Take-or-pay provisions**
- **Contract risk assessment**

---

## 🏗️ Module Architecture

### **1. Core Module: `contracts/offtake_manager.py`**

**Purpose**: Manages all offtake contracts and calculates contract-based revenue

**Key Components**:
- `OfftakeContract` class (individual contract)
- `OfftakeManager` class (manages multiple contracts)
- Contract validation and conflict resolution
- Revenue calculation engine

### **2. Contract Structure**

```python
class OfftakeContract:
    """
    Represents a single offtake agreement.
    """
    def __init__(
        self,
        contract_id: str,
        counterparty: str,
        start_year: int,
        end_year: int,
        volume_commitment_min: float,  # Minimum tons/year
        volume_commitment_max: float,   # Maximum tons/year
        pricing_type: str,              # 'fixed', 'floor', 'cap', 'indexed', 'hybrid'
        base_price: float,              # Base price per ton
        price_floor: Optional[float],  # Minimum price
        price_cap: Optional[float],    # Maximum price
        indexation_rate: Optional[float],  # Annual price escalation
        take_or_pay: bool,             # Take-or-pay provision
        take_or_pay_percentage: float,  # % of committed volume
        penalty_rate: Optional[float],  # Penalty for non-delivery
        bonus_rate: Optional[float],   # Bonus for over-delivery
        priority: int = 1              # Contract priority (1 = highest)
    ):
        pass
```

### **3. Integration Points**

#### **A. Data Input Module: `contracts/loader.py`**
- Load contracts from Excel/JSON/YAML
- Parse contract terms
- Validate contract data
- Handle multiple contract formats

#### **B. Revenue Calculator: `contracts/revenue_calculator.py`**
- Calculate contract-based revenue
- Handle volume allocation across contracts
- Apply pricing mechanisms
- Calculate penalties/bonuses

#### **C. Risk Assessment: `contracts/risk_analyzer.py`**
- Counterparty risk scoring
- Contract concentration risk
- Volume commitment risk
- Pricing risk (fixed vs. indexed)

#### **D. Excel Export: `export/excel.py` (enhancement)**
- New sheet: "Offtake Contracts"
- Contract terms display
- Revenue breakdown by contract
- Contract risk metrics

---

## 📊 Contract Types to Support

### **1. Fixed Price Contract**
- Fixed price per ton for entire duration
- Simple, predictable revenue
- Example: $50/ton for 10 years

### **2. Price Floor Contract**
- Minimum price guaranteed
- Can benefit from spot price increases
- Example: Floor at $45/ton, but can go higher

### **3. Price Cap Contract**
- Maximum price limit
- Protects buyer from price spikes
- Example: Cap at $60/ton, but can go lower

### **4. Indexed Price Contract**
- Price linked to market index
- Annual escalation (CPI, carbon index)
- Example: Base $50/ton + 2% annual escalation

### **5. Hybrid Contract**
- Combination of above
- Example: Floor $45, Cap $65, indexed to carbon market

### **6. Take-or-Pay Contract**
- Buyer commits to minimum volume
- Penalty if volume not delivered
- Example: 80% take-or-pay = buyer pays for 80% even if not delivered

---

## 🔄 Integration with Existing Model

### **1. DCF Integration**

**Current Flow**:
```
Credits → Streaming % → Revenue
```

**New Flow**:
```
Credits → Contract Allocation → Contract Pricing → Revenue
```

**Changes Needed**:
- Modify `core/dcf.py` to accept `OfftakeManager`
- Replace simple streaming calculation with contract-based revenue
- Handle multiple contracts and volume allocation

### **2. Monte Carlo Integration**

**Enhancements**:
- Contract pricing in stochastic scenarios
- Volume commitment fulfillment
- Contract default risk
- Counterparty risk scenarios

### **3. Risk Analysis Integration**

**New Risk Metrics**:
- Contract concentration risk (too many credits to one buyer)
- Counterparty credit risk
- Contract duration risk (short-term vs. long-term)
- Pricing risk (fixed vs. indexed)

### **4. Excel Export Enhancement**

**New Sheet: "Offtake Contracts"**
- Contract summary table
- Revenue breakdown by contract
- Contract terms and conditions
- Risk metrics

---

## 📁 Proposed File Structure

```
contracts/
├── __init__.py
├── offtake_manager.py      # Main contract manager
├── contract.py             # Individual contract class
├── loader.py               # Contract data loader
├── revenue_calculator.py   # Contract revenue calculation
├── risk_analyzer.py        # Contract risk analysis
└── examples/
    ├── sample_contracts.json
    ├── sample_contracts.xlsx
    └── contract_template.yaml
```

---

## 🎯 Implementation Phases

### **Phase 1: Core Contract Structure** (Week 1)
- [ ] Create `OfftakeContract` class
- [ ] Implement basic contract validation
- [ ] Support fixed price contracts
- [ ] Unit tests

### **Phase 2: Contract Manager** (Week 1-2)
- [ ] Create `OfftakeManager` class
- [ ] Handle multiple contracts
- [ ] Volume allocation logic
- [ ] Contract priority system

### **Phase 3: Revenue Calculator** (Week 2)
- [ ] Implement all pricing types (fixed, floor, cap, indexed)
- [ ] Take-or-pay calculation
- [ ] Penalty/bonus logic
- [ ] Integration with DCF

### **Phase 4: Data Loading** (Week 2-3)
- [ ] Excel contract loader
- [ ] JSON/YAML contract loader
- [ ] Contract validation
- [ ] Error handling

### **Phase 5: Risk Analysis** (Week 3)
- [ ] Counterparty risk scoring
- [ ] Contract concentration analysis
- [ ] Volume commitment risk
- [ ] Integration with existing risk module

### **Phase 6: Excel Export** (Week 3-4)
- [ ] New "Offtake Contracts" sheet
- [ ] Contract summary table
- [ ] Revenue breakdown
- [ ] Risk metrics display

### **Phase 7: Monte Carlo Integration** (Week 4)
- [ ] Contract pricing in MC scenarios
- [ ] Volume fulfillment simulation
- [ ] Contract default scenarios
- [ ] Testing and validation

---

## 📝 Contract Data Format

### **Option 1: Excel Format**

**Sheet: "Offtake Contracts"**
| Contract ID | Counterparty | Start Year | End Year | Min Volume | Max Volume | Pricing Type | Base Price | Price Floor | Price Cap | Indexation | Take-or-Pay | Priority |
|-------------|--------------|------------|----------|------------|------------|--------------|------------|-------------|-----------|-------------|-------------|----------|
| CONTRACT_01 | Buyer A | 1 | 10 | 10000 | 15000 | fixed | 50 | - | - | 0% | 80% | 1 |
| CONTRACT_02 | Buyer B | 5 | 15 | 5000 | 8000 | floor | 45 | 45 | 65 | 2% | 90% | 2 |

### **Option 2: JSON Format**

```json
{
  "contracts": [
    {
      "contract_id": "CONTRACT_01",
      "counterparty": "Buyer A",
      "start_year": 1,
      "end_year": 10,
      "volume_commitment": {
        "min_tons_per_year": 10000,
        "max_tons_per_year": 15000
      },
      "pricing": {
        "type": "fixed",
        "base_price": 50.0,
        "price_floor": null,
        "price_cap": null,
        "indexation_rate": 0.0
      },
      "terms": {
        "take_or_pay": true,
        "take_or_pay_percentage": 0.80,
        "penalty_rate": null,
        "bonus_rate": null
      },
      "priority": 1
    }
  ]
}
```

### **Option 3: YAML Format** (Human-readable)

```yaml
contracts:
  - contract_id: CONTRACT_01
    counterparty: Buyer A
    duration:
      start_year: 1
      end_year: 10
    volume:
      min_tons_per_year: 10000
      max_tons_per_year: 15000
    pricing:
      type: fixed
      base_price: 50.0
    terms:
      take_or_pay: true
      take_or_pay_percentage: 80%
    priority: 1
```

---

## 🔧 Technical Implementation Details

### **1. Volume Allocation Algorithm**

**Priority-Based Allocation**:
1. Allocate to highest priority contracts first
2. Fill minimum commitments
3. Distribute remaining volume
4. Handle over-commitment scenarios

**Example**:
- Total credits: 20,000 tons
- Contract 1 (Priority 1): Min 10,000, Max 15,000
- Contract 2 (Priority 2): Min 5,000, Max 8,000
- Allocation: Contract 1 gets 12,000, Contract 2 gets 8,000

### **2. Revenue Calculation Logic**

```python
def calculate_contract_revenue(
    contract: OfftakeContract,
    allocated_volume: float,
    spot_price: float,
    year: int
) -> float:
    """
    Calculate revenue for a contract in a given year.
    """
    # Determine effective price based on pricing type
    if contract.pricing_type == 'fixed':
        effective_price = contract.base_price
    elif contract.pricing_type == 'floor':
        effective_price = max(contract.price_floor, spot_price)
    elif contract.pricing_type == 'cap':
        effective_price = min(contract.price_cap, spot_price)
    elif contract.pricing_type == 'indexed':
        escalation = (1 + contract.indexation_rate) ** (year - contract.start_year)
        effective_price = contract.base_price * escalation
    # ... handle hybrid, etc.
    
    # Apply take-or-pay
    if contract.take_or_pay:
        committed_volume = contract.volume_commitment_min * contract.take_or_pay_percentage
        if allocated_volume < committed_volume:
            # Penalty for under-delivery
            penalty = (committed_volume - allocated_volume) * contract.penalty_rate
            revenue = committed_volume * effective_price - penalty
        else:
            revenue = allocated_volume * effective_price
    else:
        revenue = allocated_volume * effective_price
    
    return revenue
```

### **3. Integration with DCF**

**Modify `core/dcf.py`**:
```python
def run_dcf(
    self,
    data: pd.DataFrame,
    streaming_percentage: Optional[float] = None,
    offtake_manager: Optional[OfftakeManager] = None  # NEW
) -> Dict:
    """
    Run DCF with optional offtake contracts.
    """
    if offtake_manager:
        # Use contract-based revenue
        revenue = offtake_manager.calculate_total_revenue(
            credits=data['carbon_credits_gross'],
            spot_prices=data['base_carbon_price'],
            years=data.index
        )
    else:
        # Use simple streaming (backward compatible)
        revenue = credits * price * streaming_percentage
    
    # ... rest of DCF calculation
```

---

## 📊 Excel Output Enhancements

### **New Sheet: "Offtake Contracts"**

**Section 1: Contract Summary**
- Table of all contracts
- Key terms (duration, volume, pricing)
- Status indicators

**Section 2: Revenue Breakdown**
- Revenue by contract by year
- Total contract revenue vs. spot revenue
- Contract utilization rates

**Section 3: Risk Metrics**
- Contract concentration
- Counterparty risk scores
- Volume commitment fulfillment
- Pricing risk analysis

---

## 🎯 Use Cases

### **Use Case 1: Single Fixed-Price Contract**
- Simple 10-year fixed price contract
- All credits go to one buyer
- Predictable revenue stream

### **Use Case 2: Multiple Contracts with Priority**
- Contract 1: High-priority, fixed price, 10,000 tons
- Contract 2: Lower priority, floor price, 5,000 tons
- Remaining credits: Spot market
- Volume allocation based on priority

### **Use Case 3: Take-or-Pay Analysis**
- Buyer commits to 80% of volume
- Penalty if not delivered
- Risk assessment of fulfillment

### **Use Case 4: Contract vs. Spot Comparison**
- Compare contract revenue vs. spot market
- Assess value of contracts
- Optimize contract mix

---

## 🔍 Risk Analysis Enhancements

### **New Risk Flags**

1. **Contract Concentration Risk**
   - >80% of credits to one buyer = RED
   - >60% to one buyer = YELLOW
   - Diversified = GREEN

2. **Volume Commitment Risk**
   - Cannot fulfill minimum commitments = RED
   - Tight margin on commitments = YELLOW
   - Comfortable margin = GREEN

3. **Pricing Risk**
   - All fixed-price contracts in rising market = YELLOW
   - All indexed contracts = GREEN
   - Mix of pricing types = GREEN

4. **Counterparty Risk**
   - Low credit rating counterparties = YELLOW/RED
   - High credit rating = GREEN

---

## 🚀 Quick Start Example

```python
from contracts.offtake_manager import OfftakeManager
from contracts.loader import load_contracts_from_excel

# Load contracts
contracts = load_contracts_from_excel("rubicon_contracts.xlsx")
manager = OfftakeManager(contracts)

# Use in model
model = CarbonModelGenerator(...)
model.load_data("data.xlsx")

# Run DCF with contracts
dcf_results = model.run_dcf(offtake_manager=manager)

# Export to Excel (includes contract analysis)
model.export_model_to_excel("results.xlsx")
```

---

## 📋 Testing Strategy

### **Unit Tests**
- Contract validation
- Revenue calculation for each pricing type
- Volume allocation logic
- Take-or-pay calculations

### **Integration Tests**
- DCF with contracts
- Monte Carlo with contracts
- Excel export with contracts
- Risk analysis with contracts

### **Example Data**
- Sample contract files (Excel, JSON, YAML)
- Test scenarios
- Expected results

---

## 🎓 Documentation Needs

1. **Contract Format Guide**: How to structure contract data
2. **Pricing Type Guide**: Explanation of each pricing mechanism
3. **Volume Allocation Guide**: How credits are allocated
4. **Risk Analysis Guide**: Understanding contract risks
5. **Integration Guide**: How to use with existing model

---

## ✅ Success Criteria

- [ ] Can load contracts from Excel/JSON/YAML
- [ ] Supports all pricing types (fixed, floor, cap, indexed, hybrid)
- [ ] Handles multiple contracts with priority
- [ ] Integrates seamlessly with DCF
- [ ] Works with Monte Carlo simulation
- [ ] Provides contract risk analysis
- [ ] Excel export includes contract details
- [ ] Backward compatible (works without contracts)
- [ ] Well-documented with examples

---

## 🔄 Future Enhancements (Post-MVP)

1. **Contract Optimization**: Find optimal contract mix
2. **Contract Negotiation Support**: Compare contract terms
3. **Counterparty Database**: Store counterparty information
4. **Contract Renewal Modeling**: Model contract extensions
5. **Market-Based Contract Pricing**: Value contracts vs. spot
6. **Contract Portfolio Analysis**: Analyze entire portfolio

---

## 📞 Questions to Clarify

1. **Contract Data Source**: Where do contracts come from? (Excel, database, API?)
2. **Pricing Mechanisms**: What pricing types does Rubicon actually use?
3. **Volume Allocation**: How are credits allocated across contracts?
4. **Take-or-Pay**: Are take-or-pay provisions common?
5. **Counterparty Data**: Do we have counterparty credit ratings?
6. **Contract Renewals**: How are contract renewals handled?
7. **Priority System**: How is contract priority determined?

---

**This module will make the model significantly more realistic and valuable for Rubicon's specific use case!** 🎯

