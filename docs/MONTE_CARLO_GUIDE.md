# Monte Carlo Simulation Guide

## Overview

The Monte Carlo simulation uses **YOUR original carbon price forecasts** from the input data (e.g., `Analyst_Model_Test_OCC.xlsx`) as the **base/center** of the stochastic distribution. It then applies variability around those forecasts to assess probabilistic risk.

This ensures that:
- ✅ Your price forecast assumptions are respected
- ✅ The simulation builds sensitivities from your original forecasts
- ✅ Results are centered on your base case
- ✅ Uncertainty is modeled around your assumptions

---

## How It Works

### Price Path Generation

The simulation uses your original price forecasts and applies stochastic variations in two ways:

#### 1. Growth-Rate-Based Variation (Default)
- **Method**: Calculates the growth rates from your original price curve
- **Process**: Applies stochastic deviations to those growth rates
- **Result**: Price paths that follow your original forecast pattern with added variability
- **Use Case**: When you want to model uncertainty in price growth rates

**Example:**
- Your original forecast: Year 2 = $40, Year 3 = $41.60 (4% growth)
- Simulation might generate: Year 2 = $40, Year 3 = $42.75 (6.9% growth)
- The deviation (2.9%) is drawn from a normal distribution

#### 2. Percentage-Based Variation (Optional)
- **Method**: Applies percentage multipliers directly to your original prices
- **Process**: Each year's price is multiplied by a stochastic factor
- **Result**: Price paths that scale your original forecasts up/down
- **Use Case**: When you want to model uniform percentage uncertainty

**Example:**
- Your original forecast: Year 2 = $40
- Simulation might generate: Year 2 = $41.07 (2.7% multiplier)
- The multiplier is drawn from a normal distribution centered at 1.0

### Volume Path Generation

- **Method**: Applies stochastic multipliers to your original credit volumes
- **Process**: Each year's volume is multiplied by a factor from a normal distribution
- **Result**: Volume paths that scale your original volumes up/down
- **Use Case**: Models delivery uncertainty (permits, weather, operational delays)

**Example:**
- Your original forecast: Year 1 = 10,000,000 credits
- Simulation might generate: Year 1 = 8,822,388 credits (0.88x multiplier)
- The multiplier is drawn from a normal distribution (e.g., mean=1.0, std=0.15)

---

## Usage

### Basic Usage

```python
from carbon_model_template import CarbonModelGenerator

# Initialize model
model = CarbonModelGenerator(
    wacc=0.08,
    rubicon_investment_total=20_000_000,
    investment_tenor=5,
    streaming_percentage_initial=0.48,
    # Monte Carlo assumptions
    price_growth_base=0.03,        # 3% mean growth
    price_growth_std_dev=0.02,    # 2% volatility
    volume_multiplier_base=1.0,   # 100% of base volume
    volume_std_dev=0.15            # 15% volume volatility
)

# Load data (with your price forecasts)
model.load_data("Analyst_Model_Test_OCC.xlsx")

# Run Monte Carlo (uses your price forecasts as base)
results = model.run_monte_carlo(
    simulations=5000,
    use_percentage_variation=False  # Default: growth-rate-based
)
```

### Using Percentage-Based Variation

```python
# Run Monte Carlo with percentage-based price variation
results = model.run_monte_carlo(
    simulations=5000,
    use_percentage_variation=True  # Apply multipliers directly to prices
)
```

---

## Parameters

### Price Parameters

- **`price_growth_base`** (float): Mean annual price growth rate
  - Used when `use_percentage_variation=False`
  - Example: `0.03` for 3% mean growth
  
- **`price_growth_std_dev`** (float): Standard deviation of price volatility
  - Controls uncertainty around your original forecasts
  - Example: `0.02` for 2% volatility
  - When `use_percentage_variation=True`, this becomes the std dev of price multipliers

### Volume Parameters

- **`volume_multiplier_base`** (float): Mean volume multiplier
  - Typically `1.0` to center on your original volumes
  - Example: `1.0` = 100% of base volume
  
- **`volume_std_dev`** (float): Standard deviation of volume multiplier
  - Controls delivery uncertainty
  - Example: `0.15` for 15% volatility

---

## Results

The simulation returns a dictionary with:

- **`irr_series`**: Array of 5,000 simulated IRRs
- **`npv_series`**: Array of 5,000 simulated NPVs
- **`mc_mean_irr`**: Mean IRR across all simulations
- **`mc_mean_npv`**: Mean NPV across all simulations
- **`mc_p10_irr`**: 10th percentile IRR (downside risk)
- **`mc_p90_irr`**: 90th percentile IRR (upside potential)
- **`mc_p10_npv`**: 10th percentile NPV (downside risk)
- **`mc_p90_npv`**: 90th percentile NPV (upside potential)
- **`mc_std_irr`**: Standard deviation of IRR
- **`mc_std_npv`**: Standard deviation of NPV
- **`simulations`**: Total number of simulations
- **`valid_simulations`**: Number of valid (non-NaN) results

---

## Key Features

### ✅ Respects Your Forecasts
- Uses your original price forecasts as the center of the distribution
- Does not replace your assumptions, but models uncertainty around them

### ✅ Dual-Variable Stochastic Modeling
- **Price Volatility**: Stochastic variations around your price curve
- **Volume Volatility**: Stochastic multipliers on credit delivery

### ✅ Flexible Variation Methods
- Growth-rate-based: Models uncertainty in price growth
- Percentage-based: Models uniform percentage uncertainty

### ✅ Comprehensive Results
- Full distribution of outcomes (5,000 simulations)
- Statistical summaries (mean, percentiles, std dev)
- Ready for Excel export with histogram

---

## Example Output

```
Running 5000 Monte Carlo simulations...
Monte Carlo simulation complete!
  Mean IRR: 17.73%
  P10 IRR: 16.76% (downside)
  P90 IRR: 18.74% (upside)
  Mean NPV: $31,232,545
  P10 NPV: $26,651,337 (downside)
  P90 NPV: $35,868,718 (upside)
  Valid simulations: 5000/5000
```

---

## Best Practices

1. **Use Your Original Forecasts**: Load data with your price forecasts - the simulation will use them as the base
2. **Set Appropriate Volatility**: 
   - Price volatility: 1-3% for stable markets, 3-5% for volatile markets
   - Volume volatility: 10-20% for delivery uncertainty
3. **Run Enough Simulations**: 5,000 simulations provide stable statistics
4. **Use Growth-Rate-Based (Default)**: Better for modeling price growth uncertainty
5. **Use Percentage-Based**: When you want uniform percentage uncertainty across all years

---

## Technical Details

### Growth-Rate-Based Method
- Calculates annual growth rates from your original price curve
- Applies stochastic deviations: `new_growth = original_growth + random_deviation`
- Random deviation drawn from: `N(0, price_growth_std_dev)`
- Preserves your original forecast pattern

### Percentage-Based Method
- Applies multipliers directly: `new_price = original_price × multiplier`
- Multiplier drawn from: `N(1.0, price_growth_std_dev)`
- Ensures multipliers are positive (minimum 0.01)

### Volume Method
- Applies multipliers: `new_volume = original_volume × multiplier`
- Multiplier drawn from: `N(volume_multiplier_base, volume_std_dev)`
- Ensures multipliers are positive (minimum 0.01)

---

## Questions?

For more information, see:
- `ARCHITECTURE.md` - System architecture
- `EXCEL_FORMULA_GUIDE.md` - Excel output guide
- `HOW_TO_USE.md` - General usage guide

