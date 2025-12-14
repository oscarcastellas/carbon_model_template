#!/usr/bin/env python3
"""
Create Fake Dataset for Stress Testing

Creates multiple fake datasets with different scenarios to test the model:
1. High-growth scenario (aggressive price increases, high volumes)
2. Low-growth scenario (moderate price increases, lower volumes)
3. Volatile scenario (high volatility in prices and volumes)
4. Conservative scenario (stable, predictable cash flows)
"""

import pandas as pd
import numpy as np
from pathlib import Path

def create_fake_dataset(scenario_name: str, years: int = 20, start_year: int = 2025):
    """
    Create a fake dataset for a specific scenario.
    
    Parameters:
    -----------
    scenario_name : str
        Name of scenario: 'high_growth', 'low_growth', 'volatile', 'conservative'
    years : int
        Number of years of data
    start_year : int
        Starting year
        
    Returns:
    --------
    pd.DataFrame
        Dataset with columns: Year, carbon_credits_gross, base_carbon_price, project_implementation_costs
    """
    np.random.seed(42)  # For reproducibility
    
    years_list = [start_year + i for i in range(years)]
    
    if scenario_name == 'high_growth':
        # Aggressive growth: high price increases, high volumes
        base_price = 50.0
        price_growth = 0.08  # 8% annual growth
        base_credits = 1_000_000
        credit_growth = 0.15  # 15% annual growth
        base_costs = 5_000_000
        cost_growth = 0.03  # 3% annual growth
        
    elif scenario_name == 'low_growth':
        # Conservative growth: moderate price increases, lower volumes
        base_price = 40.0
        price_growth = 0.03  # 3% annual growth
        base_credits = 500_000
        credit_growth = 0.05  # 5% annual growth
        base_costs = 3_000_000
        cost_growth = 0.02  # 2% annual growth
        
    elif scenario_name == 'volatile':
        # High volatility: erratic prices and volumes
        base_price = 45.0
        price_growth = 0.05
        price_volatility = 0.20  # High volatility
        base_credits = 750_000
        credit_growth = 0.08
        credit_volatility = 0.25  # High volatility
        base_costs = 4_000_000
        cost_growth = 0.04
        
    elif scenario_name == 'conservative':
        # Stable, predictable: low volatility, steady growth
        base_price = 42.0
        price_growth = 0.04  # 4% annual growth
        base_credits = 600_000
        credit_growth = 0.06  # 6% annual growth
        base_costs = 3_500_000
        cost_growth = 0.025  # 2.5% annual growth
        
    else:
        raise ValueError(f"Unknown scenario: {scenario_name}")
    
    # Generate data
    data = []
    for i, year in enumerate(years_list):
        year_num = i + 1
        
        # Price with growth
        if scenario_name == 'volatile':
            price = base_price * (1 + price_growth) ** i * (1 + np.random.normal(0, price_volatility))
        else:
            price = base_price * (1 + price_growth) ** i
        
        # Credits with growth
        if scenario_name == 'volatile':
            credits = base_credits * (1 + credit_growth) ** i * (1 + np.random.normal(0, credit_volatility))
        else:
            credits = base_credits * (1 + credit_growth) ** i
        
        # Costs with growth
        costs = base_costs * (1 + cost_growth) ** i
        
        # Ensure non-negative values
        price = max(price, 10.0)
        credits = max(credits, 100_000)
        costs = max(costs, 1_000_000)
        
        data.append({
            'Year': year,
            'carbon_credits_gross': credits,
            'base_carbon_price': price,
            'project_implementation_costs': costs
        })
    
    df = pd.DataFrame(data)
    df = df.set_index('Year')
    
    # Ensure column names match expected format
    df.columns = ['carbon_credits_gross', 'base_carbon_price', 'project_implementation_costs']
    
    return df

def create_all_scenarios():
    """Create all scenario datasets and save them."""
    scenarios = ['high_growth', 'low_growth', 'volatile', 'conservative']
    data_dir = Path(__file__).parent.parent / 'data'
    data_dir.mkdir(exist_ok=True)
    
    print("=" * 70)
    print("CREATING FAKE DATASETS FOR STRESS TESTING")
    print("=" * 70)
    print()
    
    for scenario in scenarios:
        print(f"Creating {scenario} scenario...")
        df = create_fake_dataset(scenario)
        
        # Save as Excel (matching the format expected by DataLoader)
        filename = data_dir / f"fake_dataset_{scenario}.xlsx"
        # DataLoader reads with header=None and processes the data
        # Save with Year as index and proper column names (DataLoader will process it)
        # Reset index to make Year a column for Excel display
        df_excel = df.reset_index()
        # Reorder: Year, carbon_credits_gross, base_carbon_price, project_implementation_costs
        cols = ['Year', 'carbon_credits_gross', 'base_carbon_price', 'project_implementation_costs']
        df_excel = df_excel[cols]
        # Save to 'Data' sheet (DataLoader looks for this)
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df_excel.to_excel(writer, sheet_name='Data', index=False)
        
        print(f"  ✓ Saved: {filename}")
        print(f"    Years: {len(df)} ({df.index.min()}-{df.index.max()})")
        print(f"    Price range: ${df['base_carbon_price'].min():.2f} - ${df['base_carbon_price'].max():.2f}/ton")
        print(f"    Credits range: {df['carbon_credits_gross'].min():,.0f} - {df['carbon_credits_gross'].max():,.0f}")
        print(f"    Costs range: ${df['project_implementation_costs'].min():,.0f} - ${df['project_implementation_costs'].max():,.0f}")
        print()
    
    print("=" * 70)
    print("ALL SCENARIOS CREATED")
    print("=" * 70)
    print()
    print("Files created:")
    for scenario in scenarios:
        filename = data_dir / f"fake_dataset_{scenario}.xlsx"
        print(f"  - {filename}")
    print()
    print("Use these files to stress test the model template!")

if __name__ == '__main__':
    create_all_scenarios()

