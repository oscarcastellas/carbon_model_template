# GBM Usage Guide

## 📊 Where Are GBM Outputs?

GBM results are included in the Excel export in **multiple places**:

### **Sheet 1: Inputs & Assumptions**
- **Monte Carlo Assumptions Section** shows:
  - Use GBM Method: Yes/No
  - GBM Drift (μ): Your drift parameter
  - GBM Volatility (σ): Your volatility parameter
  - Number of Simulations
  - All other Monte Carlo parameters

### **Sheet 3: Summary & Results**
- **Monte Carlo Simulation Summary** shows:
  - Mean IRR, P10 IRR, P90 IRR
  - Mean NPV, P10 NPV, P90 NPV
  - Standard deviations

### **Sheet 5: Monte Carlo Results**
- **Top of sheet** shows:
  - **Simulation Method**: "GBM (Geometric Brownian Motion)" or "Growth-Rate Based"
  - **GBM Drift (μ)**: Your drift parameter (if GBM was used)
  - **GBM Volatility (σ)**: Your volatility parameter (if GBM was used)
- **Summary Statistics**: Mean, percentiles, standard deviations
- **Full Results Table**: All 5,000 simulated IRRs and NPVs
- **Histogram Charts**: Visual distribution of results

---

## ⚙️ How to Configure GBM Analysis

### **Method 1: Easy Configuration Script (Recommended)**

Use `analysis_config.py` - the simplest way to configure and run:

```python
from analysis_config import AnalysisConfig

config = AnalysisConfig()

# Set GBM parameters
config.use_gbm = True
config.gbm_drift = 0.03      # 3% expected return
config.gbm_volatility = 0.15  # 15% volatility
config.simulations = 5000

# Run analysis
config.run_analysis()
```

**Or use the example script:**
```bash
python3 examples/run_gbm_analysis.py
```

### **Method 2: Direct API Usage**

```python
from carbon_model_generator import CarbonModelGenerator

model = CarbonModelGenerator(
    wacc=0.08,
    rubicon_investment_total=20_000_000,
    investment_tenor=5,
    streaming_percentage_initial=0.48
)

model.load_data("data.xlsx")
model.run_dcf()

# Run Monte Carlo with GBM
mc_results = model.run_monte_carlo(
    simulations=5000,
    use_gbm=True,              # Enable GBM
    gbm_drift=0.03,            # 3% expected return
    gbm_volatility=0.15,       # 15% volatility
    volume_multiplier_base=1.0,
    volume_std_dev=0.15
)

# Export to Excel (includes GBM parameters)
model.export_model_to_excel("results.xlsx")
```

---

## 🎯 Quick Configuration Guide

### **GBM Parameters Explained**

1. **GBM Drift (μ)**: Expected annual return
   - Example: 0.03 = 3% expected annual price increase
   - Typical range: 0.02 to 0.05 (2% to 5%)

2. **GBM Volatility (σ)**: Annual price volatility
   - Example: 0.15 = 15% annual volatility
   - Typical range: 0.10 to 0.25 (10% to 25%)
   - Higher = more uncertainty/risk

3. **Number of Simulations**: How many scenarios to run
   - Default: 5,000
   - More = more accurate but slower
   - Minimum: 1,000 for reliable results

### **Example Configurations**

#### **Conservative (Low Risk)**
```python
config.use_gbm = True
config.gbm_drift = 0.02      # 2% expected return
config.gbm_volatility = 0.10  # 10% volatility (low risk)
```

#### **Moderate (Balanced)**
```python
config.use_gbm = True
config.gbm_drift = 0.03      # 3% expected return
config.gbm_volatility = 0.15  # 15% volatility
```

#### **Aggressive (High Risk)**
```python
config.use_gbm = True
config.gbm_drift = 0.04      # 4% expected return
config.gbm_volatility = 0.25  # 25% volatility (high risk)
```

---

## 🔄 Changing Assumptions

### **Option 1: Edit `analysis_config.py`**

1. Open `analysis_config.py`
2. Find the configuration section (around line 200)
3. Modify the values:
   ```python
   config.use_gbm = True
   config.gbm_drift = 0.03      # Change this
   config.gbm_volatility = 0.15  # Change this
   ```
4. Run: `python3 analysis_config.py`

### **Option 2: Use the Example Script**

1. Open `examples/run_gbm_analysis.py`
2. Modify the configuration section
3. Run: `python3 examples/run_gbm_analysis.py`

### **Option 3: Create Your Own Script**

Copy `examples/run_gbm_analysis.py` and modify as needed.

---

## 📋 Excel Inputs Sheet Structure

The **Inputs & Assumptions** sheet now includes:

```
Base Financial Assumptions
├── WACC
├── Rubicon Investment Total
├── Investment Tenor
├── Initial Streaming Percentage
├── Target IRR
└── Target Streaming Percentage

Monte Carlo Assumptions
├── Price Growth Base (Mean)
├── Price Growth Std Dev
├── Volume Multiplier Base (Mean)
├── Volume Std Dev
├── Use GBM Method          ← NEW!
├── GBM Drift (μ)           ← NEW!
├── GBM Volatility (σ)      ← NEW!
└── Number of Simulations   ← NEW!
```

**All values are editable** - change them and the model will reflect your settings.

---

## 🎨 Excel Output Structure

### **Sheet 1: Inputs & Assumptions**
- All assumptions in one place
- GBM parameters clearly labeled
- Yellow highlighting for input cells

### **Sheet 3: Summary & Results**
- Risk Assessment (with all flags listed)
- Risk Score breakdown
- Breakeven Analysis
- Monte Carlo Summary

### **Sheet 5: Monte Carlo Results**
- **Method indicator** at top (GBM vs. Growth-Rate)
- **GBM parameters** shown if GBM was used
- Summary statistics
- Full 5,000 simulation results
- **Histogram charts** (IRR and NPV distributions)

---

## 💡 Tips

1. **Start with defaults**: Use the example script first to see how it works
2. **Adjust gradually**: Change one parameter at a time to see impact
3. **Compare methods**: Run both GBM and Growth-Rate to compare results
4. **Use random seed**: Set `random_seed=42` for reproducible results
5. **Check Excel**: All parameters are visible in Sheet 1

---

## 🚀 Quick Start

```bash
# Run GBM analysis with default settings
python3 examples/run_gbm_analysis.py
```

This will:
1. Show your configuration
2. Ask for confirmation
3. Run the full analysis
4. Generate Excel with all GBM outputs
5. Show where to find results

---

**All GBM outputs are in the Excel file - check Sheet 1 (Inputs) and Sheet 5 (Monte Carlo Results)!**

