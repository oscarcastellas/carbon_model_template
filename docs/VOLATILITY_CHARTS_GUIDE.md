# Carbon Price Volatility Charts - Complete Output Guide

## 📊 Full Chart Output Suite

The carbon price volatility module generates **5 comprehensive charts** that visualize:

1. **Price Volatility Paths** - How prices evolve with volatility
2. **Price Distribution Over Time** - Price uncertainty at key horizons
3. **Returns Distribution** - Impact on IRR and NPV
4. **Volatility Heatmap** - Percentile distribution visualization
5. **Correlation Analysis** - Relationship between volatility and returns

---

## 🎯 Chart 1: Price Volatility Paths

**File**: `volatility_charts/carbon_price_volatility_price_paths.png`

### What It Shows:
- **Base Forecast** (black line): Your original price forecast
- **50 Sample GBM Paths** (light blue): Individual stochastic price paths
- **P10-P90 Range** (shaded area): 80% confidence interval
- **Median Path** (dashed blue): Expected price path (P50)

### Key Insights:
- Visualizes **price uncertainty** over 20 years
- Shows how volatility creates **wide price ranges**
- Demonstrates **stochastic nature** of carbon prices
- Base forecast vs. actual volatility outcomes

### Use Cases:
- Present to stakeholders showing price risk
- Understand volatility impact on forecasts
- Compare base case vs. stochastic outcomes

---

## 📈 Chart 2: Price Distribution Over Time

**File**: `volatility_charts/carbon_price_volatility_price_distribution.png`

### What It Shows:
- **4 Histograms** showing price distribution at:
  - Year 5
  - Year 10
  - Year 15
  - Year 20
- **Mean** (red dashed line): Average price
- **Median** (green dashed line): 50th percentile price

### Key Insights:
- **Price uncertainty increases** over time
- Distribution becomes **wider** in later years
- Shows **volatility compounding** effect
- Standard deviation increases with time horizon

### Use Cases:
- Assess price risk at specific time horizons
- Understand long-term vs. short-term uncertainty
- Plan for different price scenarios

---

## 💰 Chart 3: Returns Distribution (IRR & NPV)

**File**: `volatility_charts/carbon_price_volatility_returns_distribution.png`

### What It Shows:
- **Left Panel**: IRR Distribution
  - Histogram of all simulated IRRs
  - Mean IRR (red line)
  - P10 IRR (orange line) - downside risk
  - P90 IRR (green line) - upside potential
- **Right Panel**: NPV Distribution
  - Histogram of all simulated NPVs
  - Mean NPV (red line)
  - P10 NPV (orange line) - downside
  - P90 NPV (green line) - upside

### Key Insights:
- **Risk distribution** of investment returns
- **Downside risk** (P10) vs. **upside potential** (P90)
- **Standard deviation** shows volatility impact
- Most likely outcomes vs. tail risks

### Use Cases:
- Risk assessment for investment decisions
- Present to investment committee
- Compare different volatility scenarios
- Understand probability of different outcomes

---

## 🔥 Chart 4: Volatility Heatmap

**File**: `volatility_charts/carbon_price_volatility_volatility_heatmap.png`

### What It Shows:
- **Heatmap** with price percentiles (P10, P25, P50, P75, P90) by year
- **Color coding**: Green = higher prices, Red = lower prices
- **Percentile values** shown in each cell

### Key Insights:
- **Price range** at each percentile over time
- **Volatility spread** (P90 - P10) increases over time
- **Median path** (P50) shows expected trajectory
- **Confidence intervals** at a glance

### Use Cases:
- Quick reference for price scenarios
- Understand percentile ranges
- Compare different volatility levels
- Visual risk assessment

---

## 🔗 Chart 5: Correlation Analysis

**File**: `volatility_charts/carbon_price_volatility_correlation_analysis.png`

### What It Shows:
**4 Scatter Plots** analyzing relationships:

1. **Price Volatility vs. IRR** (top left)
   - Correlation coefficient shown
   - Shows if higher volatility = higher/lower returns

2. **Price Volatility vs. NPV** (top right)
   - How volatility affects NPV
   - Correlation coefficient

3. **Final Price vs. IRR** (bottom left)
   - Relationship between end price and returns
   - Shows price impact on IRR

4. **Final Price vs. NPV** (bottom right)
   - Direct price-to-value relationship
   - Correlation coefficient

### Key Insights:
- **Correlation coefficients** show strength of relationships
- Understand **what drives returns**
- Identify **key risk factors**
- Price volatility impact on investment outcomes

### Use Cases:
- Understand drivers of returns
- Identify risk factors
- Optimize investment strategy
- Risk factor analysis

---

## 🚀 How to Generate Charts

### **Method 1: Quick Generation**

```bash
python3 examples/generate_volatility_charts.py
```

This will:
1. Run GBM analysis with default settings
2. Generate all 5 charts
3. Save to `volatility_charts/` directory

### **Method 2: Custom Configuration**

```python
from analysis.volatility_visualizer import VolatilityVisualizer
from analysis.gbm_simulator import GBMPriceSimulator

# Generate GBM paths
gbm_sim = GBMPriceSimulator()
paths = [gbm_sim.generate_gbm_path_from_base(
    base_prices=your_prices,
    drift=0.03,
    volatility=0.15
) for _ in range(1000)]

# Create visualizer
visualizer = VolatilityVisualizer(output_dir="my_charts")

# Generate specific chart
fig = visualizer.plot_price_paths(
    base_prices=your_prices,
    gbm_paths=paths
)
plt.show()
```

### **Method 3: Full Report**

```python
from analysis.volatility_visualizer import VolatilityVisualizer

visualizer = VolatilityVisualizer()
saved_files = visualizer.generate_full_report(
    base_prices=base_prices,
    gbm_paths=gbm_paths,
    monte_carlo_results=mc_results
)

# Access individual charts
print(saved_files['price_paths'])
print(saved_files['returns_distribution'])
# etc.
```

---

## 📋 Chart Specifications

### **Resolution**
- All charts saved at **300 DPI** (publication quality)
- High resolution for presentations and reports

### **Format**
- **PNG format** (easy to embed in documents)
- Can be converted to PDF, SVG, etc.

### **Sizes**
- **Price Paths**: 14×8 inches
- **Price Distribution**: 16×12 inches (4 subplots)
- **Returns Distribution**: 16×6 inches (2 subplots)
- **Heatmap**: 14×8 inches
- **Correlation**: 16×12 inches (4 subplots)

---

## 🎨 Customization Options

### **Change Colors**
```python
# In volatility_visualizer.py, modify:
ax.plot(..., color='your_color')
```

### **Change Chart Size**
```python
fig, ax = plt.subplots(figsize=(16, 10))  # Custom size
```

### **Add Annotations**
```python
ax.annotate('Key Point', xy=(x, y), ...)
```

### **Change Style**
```python
plt.style.use('seaborn-darkgrid')  # Different style
```

---

## 📊 Integration with Excel

The charts complement the Excel output:

- **Excel**: Detailed data tables, formulas, histograms
- **Charts**: Visual analysis, presentations, reports

**Best Practice**: Use both!
- Excel for detailed analysis
- Charts for presentations and quick insights

---

## 💡 Interpretation Guide

### **Reading Price Paths Chart**
- **Wider shaded area** = Higher volatility risk
- **Paths above base** = Upside scenarios
- **Paths below base** = Downside scenarios

### **Reading Distribution Charts**
- **Wider histogram** = More uncertainty
- **Skewed distribution** = Asymmetric risk
- **P10-P90 range** = 80% confidence interval

### **Reading Correlation Charts**
- **Correlation > 0.5** = Strong positive relationship
- **Correlation < -0.5** = Strong negative relationship
- **Correlation near 0** = No relationship

---

## 🎯 Use Cases

### **1. Investment Committee Presentation**
- Use **Price Paths** and **Returns Distribution** charts
- Show risk vs. return trade-off
- Highlight downside scenarios

### **2. Risk Assessment**
- Use **Price Distribution** and **Heatmap**
- Understand price uncertainty
- Assess tail risks

### **3. Scenario Analysis**
- Use **Correlation Analysis**
- Understand drivers of returns
- Identify key risk factors

### **4. Reporting**
- Include all charts in appendix
- Use for detailed analysis
- Support Excel findings

---

## ✅ Summary

**5 Comprehensive Charts** showing:

1. ✅ **Price Volatility Paths** - Stochastic price evolution
2. ✅ **Price Distribution** - Uncertainty over time
3. ✅ **Returns Distribution** - IRR and NPV risk
4. ✅ **Volatility Heatmap** - Percentile visualization
5. ✅ **Correlation Analysis** - Risk factor relationships

**All charts are:**
- High resolution (300 DPI)
- Publication quality
- Ready for presentations
- Saved in `volatility_charts/` directory

**Generate with:**
```bash
python3 examples/generate_volatility_charts.py
```

---

**Your complete carbon price volatility visualization suite!** 📊🎯

