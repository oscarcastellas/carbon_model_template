#!/usr/bin/env python3
"""
Generate Full Volatility Analysis Charts

This script runs GBM analysis and generates comprehensive charts showing
carbon price volatility and its impact on investment returns.
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from analysis_config import AnalysisConfig
from analysis.volatility_visualizer import VolatilityVisualizer
from analysis.gbm_simulator import GBMPriceSimulator
import pandas as pd
import numpy as np


def generate_gbm_paths_for_visualization(
    base_prices: pd.Series,
    gbm_drift: float,
    gbm_volatility: float,
    num_paths: int = 1000
) -> list:
    """
    Generate multiple GBM paths for visualization.
    
    Parameters:
    -----------
    base_prices : pd.Series
        Base price forecast
    gbm_drift : float
        GBM drift parameter
    gbm_volatility : float
        GBM volatility parameter
    num_paths : int
        Number of paths to generate
        
    Returns:
    --------
    list
        List of GBM price paths
    """
    gbm_sim = GBMPriceSimulator()
    paths = []
    
    for i in range(num_paths):
        path = gbm_sim.generate_gbm_path_from_base(
            base_prices=base_prices,
            drift=gbm_drift,
            volatility=gbm_volatility,
            random_seed=None  # Random for each path
        )
        paths.append(path)
    
    return paths


def main():
    """Generate full volatility analysis with charts."""
    print()
    print("="*70)
    print("CARBON PRICE VOLATILITY ANALYSIS - CHART GENERATION")
    print("="*70)
    print()
    
    # Configure analysis
    config = AnalysisConfig()
    
    # GBM Configuration
    config.use_gbm = True
    config.gbm_drift = 0.03      # 3% expected return
    config.gbm_volatility = 0.15  # 15% volatility
    config.simulations = 5000
    config.random_seed = 42
    config.data_file = "Analyst_Model_Test_OCC.xlsx"
    
    print("Configuration:")
    print(f"  GBM Drift: {config.gbm_drift:.2%}")
    print(f"  GBM Volatility: {config.gbm_volatility:.2%}")
    print(f"  Simulations: {config.simulations:,}")
    print()
    
    # Initialize model and load data
    print("1. Loading data...")
    from data.loader import DataLoader
    loader = DataLoader()
    data = loader.load_data(config.data_file)
    base_prices = data['base_carbon_price']
    print(f"   ✓ Loaded {len(data)} years of data")
    print(f"   ✓ Base prices range: ${base_prices.min():.2f} - ${base_prices.max():.2f}/ton")
    print()
    
    # Generate GBM paths for visualization
    print("2. Generating GBM price paths for visualization...")
    print("   (Generating 1,000 paths for charts, 5,000 for Monte Carlo)")
    gbm_paths = generate_gbm_paths_for_visualization(
        base_prices=base_prices,
        gbm_drift=config.gbm_drift,
        gbm_volatility=config.gbm_volatility,
        num_paths=1000
    )
    print(f"   ✓ Generated {len(gbm_paths)} price paths")
    print()
    
    # Run Monte Carlo analysis
    print("3. Running Monte Carlo analysis...")
    from core.dcf import DCFCalculator
    from core.irr import IRRCalculator
    from analysis.monte_carlo import MonteCarloSimulator
    
    irr_calc = IRRCalculator()
    dcf_calc = DCFCalculator(
        wacc=config.wacc,
        rubicon_investment_total=config.rubicon_investment_total,
        investment_tenor=config.investment_tenor,
        irr_calculator=irr_calc
    )
    mc_sim = MonteCarloSimulator(dcf_calc, irr_calc)
    
    mc_results = mc_sim.run_monte_carlo(
        base_data=data,
        streaming_percentage=config.streaming_percentage_initial,
        price_growth_base=config.price_growth_base,
        price_growth_std_dev=config.price_growth_std_dev,
        volume_multiplier_base=config.volume_multiplier_base,
        volume_std_dev=config.volume_std_dev,
        simulations=config.simulations,
        random_seed=config.random_seed,
        use_percentage_variation=config.use_percentage_variation,
        use_gbm=True,
        gbm_drift=config.gbm_drift,
        gbm_volatility=config.gbm_volatility
    )
    
    print(f"   ✓ Mean IRR: {mc_results['mc_mean_irr']:.2%}")
    print(f"   ✓ P10 IRR: {mc_results['mc_p10_irr']:.2%}")
    print(f"   ✓ P90 IRR: {mc_results['mc_p90_irr']:.2%}")
    print()
    
    # Generate charts
    print("4. Generating volatility analysis charts...")
    visualizer = VolatilityVisualizer(output_dir="volatility_charts")
    
    saved_files = visualizer.generate_full_report(
        base_prices=base_prices,
        gbm_paths=gbm_paths,
        monte_carlo_results=mc_results,
        output_prefix="carbon_price_volatility"
    )
    
    print()
    print("="*70)
    print("ANALYSIS COMPLETE!")
    print("="*70)
    print()
    print("📊 Generated Charts:")
    print()
    print("1. Price Paths Chart")
    print("   Shows multiple GBM price paths with base forecast and percentiles")
    print(f"   → {saved_files['price_paths']}")
    print()
    print("2. Price Distribution Over Time")
    print("   Shows price distribution at Years 5, 10, 15, 20")
    print(f"   → {saved_files['price_distribution']}")
    print()
    print("3. Returns Distribution (IRR & NPV)")
    print("   Shows distribution of IRR and NPV from price volatility")
    print(f"   → {saved_files['returns_distribution']}")
    print()
    print("4. Volatility Heatmap")
    print("   Shows price percentiles (P10, P25, P50, P75, P90) over time")
    print(f"   → {saved_files['volatility_heatmap']}")
    print()
    print("5. Correlation Analysis")
    print("   Shows correlation between price volatility and returns")
    print(f"   → {saved_files['correlation_analysis']}")
    print()
    print("="*70)
    print("All charts saved in 'volatility_charts/' directory")
    print("="*70)
    print()


if __name__ == "__main__":
    main()

