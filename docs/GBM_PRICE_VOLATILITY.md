# GBM and Carbon Price Volatility

## 🎯 Yes! GBM is Specifically for Carbon Price Volatility

**Geometric Brownian Motion (GBM)** is a financial model designed to simulate **asset price volatility** - in this case, **carbon credit price volatility**.

---

## 📊 What GBM Models

GBM simulates how **carbon prices evolve over time** with **volatility**:

```
dS = μS dt + σS dW
```

Where:
- **S** = Carbon price per ton
- **μ (drift)** = Expected annual return (e.g., 3% growth)
- **σ (volatility)** = Annual price volatility (e.g., 15%)
- **dW** = Random shocks (market uncertainty)

---

## 🔍 How It Works

### **1. Price Volatility Simulation**

GBM generates **stochastic price paths** that:
- Start from your base price forecast
- Apply **random volatility** each year
- Model **uncertainty** in future carbon prices

**Example:**
- Year 1: $50/ton (your forecast)
- Year 2: $52/ton (could be higher or lower due to volatility)
- Year 3: $48/ton (random market movements)
- ... continues for 20 years

### **2. Volatility Parameters**

**GBM Drift (μ)**: Expected annual price growth
- Example: 3% = prices expected to grow 3% per year on average
- Represents long-term trend

**GBM Volatility (σ)**: Annual price volatility
- Example: 15% = prices can swing ±15% per year
- Represents **market uncertainty and risk**
- Higher volatility = more price risk

### **3. Monte Carlo Integration**

In Monte Carlo analysis, GBM:
1. Generates **5,000 different price paths**
2. Each path has **different volatility outcomes**
3. Calculates IRR/NPV for each scenario
4. Shows **risk distribution** of outcomes

---

## 💡 Why GBM for Carbon Prices?

### **Real-World Carbon Price Volatility**

Carbon prices are **highly volatile** due to:
- Policy changes
- Market demand fluctuations
- Regulatory uncertainty
- Economic conditions
- Supply/demand imbalances

### **GBM Captures This Volatility**

GBM models:
- ✅ **Random price movements** (market shocks)
- ✅ **Long-term trends** (drift)
- ✅ **Volatility clustering** (high volatility periods)
- ✅ **Uncertainty** in future prices

---

## 📈 Example: How Volatility Affects Results

### **Low Volatility (σ = 10%)**
```
Year 1: $50/ton
Year 2: $51/ton (±10% variation)
Year 3: $52/ton
→ More predictable, lower risk
```

### **High Volatility (σ = 25%)**
```
Year 1: $50/ton
Year 2: $60/ton (±25% variation - could be $37.50 or $62.50)
Year 3: $45/ton
→ Less predictable, higher risk
```

**Impact on IRR:**
- Low volatility: IRR range might be 16% - 19%
- High volatility: IRR range might be 12% - 22%
- **Higher volatility = wider risk distribution**

---

## 🎯 GBM vs. Growth-Rate Method

### **Growth-Rate Method**
- Applies **fixed growth rate** with small variations
- Less realistic for volatile markets
- Simpler but less accurate

### **GBM Method** ⭐
- Models **true price volatility** (random walk)
- More realistic for carbon markets
- Industry-standard for financial modeling
- Better captures **market uncertainty**

---

## 📊 What You See in Results

### **Monte Carlo Results Show:**

1. **Mean IRR**: Average across all volatility scenarios
2. **P10 IRR**: 10th percentile (downside risk)
3. **P90 IRR**: 90th percentile (upside potential)
4. **Standard Deviation**: Measure of volatility impact

**Example Output:**
```
Mean IRR: 17.8%
P10 IRR: 16.2%  ← Worst 10% of scenarios (high volatility downside)
P90 IRR: 19.4%  ← Best 10% of scenarios (high volatility upside)
Std Dev: 0.8%   ← Volatility impact
```

**Wider P10-P90 range = Higher price volatility risk**

---

## ⚙️ Configuring Volatility

### **Conservative (Low Volatility)**
```python
config.gbm_volatility = 0.10  # 10% annual volatility
```
- Lower risk
- More predictable outcomes
- Tighter IRR distribution

### **Moderate (Market Realistic)**
```python
config.gbm_volatility = 0.15  # 15% annual volatility
```
- Balanced risk
- Realistic for carbon markets
- Moderate IRR spread

### **Aggressive (High Volatility)**
```python
config.gbm_volatility = 0.25  # 25% annual volatility
```
- Higher risk
- More uncertainty
- Wider IRR distribution

---

## 🔬 Technical Details

### **GBM Formula**
```
S(t+1) = S(t) × exp((μ - σ²/2) × Δt + σ × √Δt × Z)
```

Where:
- **S(t)** = Price at time t
- **μ** = Drift (expected return)
- **σ** = Volatility
- **Z** = Random shock (standard normal)

### **Volatility Impact**
- **σ²/2 term**: Adjusts for volatility drag
- **σ × √Δt × Z**: Random volatility component
- **Higher σ** = More random price movements

---

## 📋 Summary

✅ **GBM is specifically designed for carbon price volatility**

✅ **Volatility (σ) parameter controls price uncertainty**

✅ **Monte Carlo shows how volatility affects IRR/NPV**

✅ **Higher volatility = wider risk distribution**

✅ **Industry-standard approach for financial modeling**

---

## 🚀 Quick Start

To model carbon price volatility:

```python
config.use_gbm = True
config.gbm_drift = 0.03      # 3% expected growth
config.gbm_volatility = 0.15  # 15% annual volatility ← This is the key!
```

**The `gbm_volatility` parameter directly controls carbon price volatility!**

---

**GBM = Carbon Price Volatility Simulator** 🎯

