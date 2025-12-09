#!/usr/bin/env python3
"""
Simple GBM Analysis Example

This script demonstrates how to easily configure and run GBM analysis.
Just modify the settings below and run!
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from analysis_config import AnalysisConfig


def main():
    """Run GBM analysis with easy configuration."""
    print()
    print("="*70)
    print("GBM ANALYSIS - EASY CONFIGURATION")
    print("="*70)
    print()
    
    # Create configuration
    config = AnalysisConfig()
    
    # ============================================
    # CONFIGURE YOUR ANALYSIS HERE
    # ============================================
    
    # Base financial assumptions
    config.wacc = 0.08
    config.rubicon_investment_total = 20_000_000
    config.investment_tenor = 5
    config.streaming_percentage_initial = 0.48
    
    # Monte Carlo settings
    config.simulations = 5000  # Number of simulations
    
    # ===== GBM CONFIGURATION =====
    # Set use_gbm = True to use Geometric Brownian Motion
    config.use_gbm = True
    
    # GBM Parameters
    config.gbm_drift = 0.03      # 3% expected annual return
    config.gbm_volatility = 0.15  # 15% annual volatility
    
    # Alternative: Use Growth-Rate method instead
    # config.use_gbm = False
    # config.price_growth_base = 0.03
    # config.price_growth_std_dev = 0.02
    
    # Volume assumptions (same for both methods)
    config.volume_multiplier_base = 1.0
    config.volume_std_dev = 0.15  # 15% volume volatility
    
    # Other settings
    config.random_seed = 42  # For reproducibility (set to None for random)
    config.data_file = "Analyst_Model_Test_OCC.xlsx"
    config.output_file = "gbm_analysis_results.xlsx"
    
    # ============================================
    # END CONFIGURATION - RUN ANALYSIS
    # ============================================
    
    # Print configuration
    config.print_config()
    
    # Run analysis
    print("Starting analysis...")
    print()
    config.run_analysis()
    
    print()
    print("="*70)
    print("WHERE TO FIND YOUR RESULTS")
    print("="*70)
    print()
    print(f"📊 Excel File: {config.output_file}")
    print()
    print("Sheet 1: Inputs & Assumptions")
    print("  • All your GBM parameters are shown here")
    print("  • Use GBM Method: Yes/No")
    print("  • GBM Drift (μ): Your drift setting")
    print("  • GBM Volatility (σ): Your volatility setting")
    print()
    print("Sheet 3: Summary & Results")
    print("  • Risk Assessment (with detailed flags)")
    print("  • Risk Score breakdown")
    print("  • Breakeven Analysis")
    print("  • Monte Carlo Summary (Mean, P10, P90)")
    print()
    print("Sheet 5: Monte Carlo Results")
    print("  • Simulation Method: Shows 'GBM (Geometric Brownian Motion)'")
    print("  • GBM Drift and Volatility parameters")
    print("  • Full 5,000 simulation results (IRR and NPV)")
    print("  • Histogram charts showing distribution")
    print()


if __name__ == "__main__":
    main()

