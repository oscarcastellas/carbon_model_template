#!/usr/bin/env python3
"""
Monte Carlo & GBM Analysis Configuration Tool

This script provides an easy way to configure and run Monte Carlo and GBM analysis.
You can set all assumptions in one place and run the analysis with a single command.
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Import using the same pattern as test scripts
from data.loader import DataLoader
from core.dcf import DCFCalculator
from core.irr import IRRCalculator
from analysis.goal_seeker import GoalSeeker
from analysis.sensitivity import SensitivityAnalyzer
from core.payback import PaybackCalculator
from analysis.monte_carlo import MonteCarloSimulator
from risk.flagger import RiskFlagger
from valuation.breakeven import BreakevenCalculator
from risk.scorer import RiskScoreCalculator
from export.excel import ExcelExporter
import pandas as pd

# Import main class - need to handle relative imports
try:
    from carbon_model_generator import CarbonModelGenerator
except ImportError:
    # If direct import fails, we'll use components directly
    CarbonModelGenerator = None


class AnalysisConfig:
    """
    Configuration class for Monte Carlo and GBM analysis.
    
    Makes it easy to set all assumptions and run analysis.
    """
    
    def __init__(self):
        """Initialize with default configuration."""
        # Base financial assumptions
        self.wacc = 0.08
        self.rubicon_investment_total = 20_000_000
        self.investment_tenor = 5
        self.streaming_percentage_initial = 0.48
        
        # Monte Carlo assumptions
        self.simulations = 5000
        self.price_growth_base = 0.03  # 3% mean growth
        self.price_growth_std_dev = 0.02  # 2% volatility
        self.volume_multiplier_base = 1.0
        self.volume_std_dev = 0.15  # 15% volatility
        
        # GBM-specific assumptions
        self.use_gbm = False  # Set to True to use GBM
        self.gbm_drift = 0.03  # 3% expected return
        self.gbm_volatility = 0.15  # 15% volatility
        
        # Analysis options
        self.use_percentage_variation = False
        self.random_seed = None  # Set to a number for reproducibility
        
        # Data file
        self.data_file = "Analyst_Model_Test_OCC.xlsx"
        
        # Output file
        self.output_file = "analysis_results.xlsx"
    
    def print_config(self):
        """Print current configuration."""
        print("="*70)
        print("CURRENT ANALYSIS CONFIGURATION")
        print("="*70)
        print()
        print("Base Financial Assumptions:")
        print(f"  WACC: {self.wacc:.2%}")
        print(f"  Investment: ${self.rubicon_investment_total:,.0f}")
        print(f"  Tenor: {self.investment_tenor} years")
        print(f"  Initial Streaming: {self.streaming_percentage_initial:.2%}")
        print()
        print("Monte Carlo Settings:")
        print(f"  Simulations: {self.simulations:,}")
        print(f"  Price Growth (Mean): {self.price_growth_base:.2%}")
        print(f"  Price Growth (Std Dev): {self.price_growth_std_dev:.2%}")
        print(f"  Volume Multiplier (Mean): {self.volume_multiplier_base:.2f}")
        print(f"  Volume (Std Dev): {self.volume_std_dev:.2%}")
        print()
        print("GBM Settings:")
        print(f"  Use GBM: {self.use_gbm}")
        if self.use_gbm:
            print(f"  GBM Drift (μ): {self.gbm_drift:.2%}")
            print(f"  GBM Volatility (σ): {self.gbm_volatility:.2%}")
        print()
        print("Other Settings:")
        print(f"  Use Percentage Variation: {self.use_percentage_variation}")
        print(f"  Random Seed: {self.random_seed if self.random_seed else 'None (random)'}")
        print(f"  Data File: {self.data_file}")
        print(f"  Output File: {self.output_file}")
        print()
    
    def run_analysis(self):
        """Run the complete analysis with current configuration."""
        print("="*70)
        print("RUNNING ANALYSIS")
        print("="*70)
        print()
        
        # Initialize model
        print("1. Initializing model...")
        if CarbonModelGenerator is None:
            # Use components directly if main class import fails
            return self._run_analysis_components()
        
        model = CarbonModelGenerator(
            wacc=self.wacc,
            rubicon_investment_total=self.rubicon_investment_total,
            investment_tenor=self.investment_tenor,
            streaming_percentage_initial=self.streaming_percentage_initial,
            price_growth_base=self.price_growth_base,
            price_growth_std_dev=self.price_growth_std_dev,
            volume_multiplier_base=self.volume_multiplier_base,
            volume_std_dev=self.volume_std_dev
        )
        print("   ✓ Model initialized")
        print()
        
        # Load data
        print("2. Loading data...")
        if not os.path.exists(self.data_file):
            print(f"   ✗ ERROR: {self.data_file} not found!")
            return None
        
        model.load_data(self.data_file)
        print(f"   ✓ Data loaded: {len(model.data)} years")
        print()
        
        # Run DCF
        print("3. Running DCF analysis...")
        dcf_results = model.run_dcf()
        print(f"   ✓ NPV: ${dcf_results['npv']:,.2f}")
        print(f"   ✓ IRR: {dcf_results['irr']:.2%}")
        print()
        
        # Goal-seeking
        print("4. Running goal-seeking (target IRR = 20%)...")
        try:
            goal_results = model.find_target_irr_stream(target_irr=0.20)
            print(f"   ✓ Target Streaming: {goal_results['streaming_percentage']:.2%}")
        except ValueError as e:
            print(f"   ⚠ {e}")
            print("   Using initial streaming percentage...")
        print()
        
        # Monte Carlo
        print("5. Running Monte Carlo simulation...")
        if self.use_gbm:
            print(f"   Method: GBM (Geometric Brownian Motion)")
            print(f"   Drift: {self.gbm_drift:.2%}, Volatility: {self.gbm_volatility:.2%}")
        else:
            print(f"   Method: Growth-Rate Based")
        
        print(f"   Simulations: {self.simulations:,}")
        print("   (This may take 1-3 minutes...)")
        print()
        
        mc_results = model.run_monte_carlo(
            simulations=self.simulations,
            price_growth_base=self.price_growth_base,
            price_growth_std_dev=self.price_growth_std_dev,
            volume_multiplier_base=self.volume_multiplier_base,
            volume_std_dev=self.volume_std_dev,
            random_seed=self.random_seed,
            use_percentage_variation=self.use_percentage_variation,
            use_gbm=self.use_gbm,
            gbm_drift=self.gbm_drift,
            gbm_volatility=self.gbm_volatility
        )
        
        print()
        print("   ✓ Monte Carlo complete!")
        print(f"   Mean IRR: {mc_results['mc_mean_irr']:.2%}")
        print(f"   P10 IRR: {mc_results['mc_p10_irr']:.2%}")
        print(f"   P90 IRR: {mc_results['mc_p90_irr']:.2%}")
        print()
        
        # Export to Excel
        print("6. Exporting to Excel...")
        model.export_model_to_excel(self.output_file)
        print(f"   ✓ Excel file created: {self.output_file}")
        print()
        
        print("="*70)
        print("ANALYSIS COMPLETE!")
        print("="*70)
        print()
        print(f"📊 Open '{self.output_file}' to view results")
        print()
        
        return model, mc_results
    
    def _run_analysis_components(self):
        """Run analysis using components directly (fallback if main class unavailable)."""
        print("   Using component-based approach...")
        
        # Initialize components
        loader = DataLoader()
        irr_calc = IRRCalculator()
        dcf_calc = DCFCalculator(
            wacc=self.wacc,
            rubicon_investment_total=self.rubicon_investment_total,
            investment_tenor=self.investment_tenor,
            irr_calculator=irr_calc
        )
        mc_sim = MonteCarloSimulator(dcf_calc, irr_calc)
        risk_flagger = RiskFlagger()
        risk_scorer = RiskScoreCalculator()
        breakeven_calc = BreakevenCalculator(dcf_calc, irr_calc)
        payback_calc = PaybackCalculator()
        exporter = ExcelExporter()
        
        # Load data
        print("2. Loading data...")
        data = loader.load_data(self.data_file)
        print(f"   ✓ Data loaded: {len(data)} years")
        print()
        
        # Run DCF
        print("3. Running DCF...")
        streaming = self.streaming_percentage_initial
        dcf_results = dcf_calc.run_dcf(data, streaming)
        npv = dcf_results['npv']
        irr = dcf_results['irr']
        payback = payback_calc.calculate_payback_period(dcf_results['cash_flows'])
        print(f"   ✓ NPV: ${npv:,.2f}, IRR: {irr:.2%}, Payback: {payback:.2f} years")
        print()
        
        # Monte Carlo
        print("4. Running Monte Carlo...")
        if self.use_gbm:
            print(f"   Method: GBM (Drift: {self.gbm_drift:.2%}, Volatility: {self.gbm_volatility:.2%})")
        mc_results = mc_sim.run_monte_carlo(
            base_data=data,
            streaming_percentage=streaming,
            price_growth_base=self.price_growth_base,
            price_growth_std_dev=self.price_growth_std_dev,
            volume_multiplier_base=self.volume_multiplier_base,
            volume_std_dev=self.volume_std_dev,
            simulations=self.simulations,
            random_seed=self.random_seed,
            use_percentage_variation=self.use_percentage_variation,
            use_gbm=self.use_gbm,
            gbm_drift=self.gbm_drift,
            gbm_volatility=self.gbm_volatility
        )
        print(f"   ✓ Mean IRR: {mc_results['mc_mean_irr']:.2%}")
        print()
        
        # Risk analysis
        print("5. Calculating risk metrics...")
        risk_flags = risk_flagger.flag_risks(irr, npv, payback, 
            credit_volumes=data['carbon_credits_gross'],
            project_costs=data['project_implementation_costs'])
        risk_score = risk_scorer.calculate_overall_risk_score(
            irr, npv, payback,
            credit_volumes=data['carbon_credits_gross'],
            base_prices=data['base_carbon_price'],
            project_costs=data['project_implementation_costs'],
            total_investment=self.rubicon_investment_total
        )
        print(f"   ✓ Risk Level: {risk_flags['risk_level'].upper()}, Score: {risk_score['overall_risk_score']}/100")
        print()
        
        # Breakeven
        print("6. Calculating breakeven...")
        breakeven = breakeven_calc.calculate_all_breakevens(data, streaming, 0.0)
        print("   ✓ Breakeven calculated")
        print()
        
        # Export
        print("7. Exporting to Excel...")
        assumptions = {
            'wacc': self.wacc,
            'rubicon_investment_total': self.rubicon_investment_total,
            'investment_tenor': self.investment_tenor,
            'streaming_percentage_initial': self.streaming_percentage_initial,
            'price_growth_base': self.price_growth_base,
            'price_growth_std_dev': self.price_growth_std_dev,
            'volume_multiplier_base': self.volume_multiplier_base,
            'volume_std_dev': self.volume_std_dev,
            'use_gbm': self.use_gbm,
            'gbm_drift': self.gbm_drift if self.use_gbm else None,
            'gbm_volatility': self.gbm_volatility if self.use_gbm else None,
            'simulations': self.simulations
        }
        
        exporter.export_model_to_excel(
            filename=self.output_file,
            assumptions=assumptions,
            target_streaming_percentage=streaming,
            target_irr=0.20,
            actual_irr=irr,
            valuation_schedule=dcf_results['results_df'],
            sensitivity_table=None,
            payback_period=payback,
            monte_carlo_results=mc_results,
            risk_flags=risk_flags,
            risk_score=risk_score,
            breakeven_results=breakeven
        )
        print(f"   ✓ Excel file created: {self.output_file}")
        print()
        
        return None, mc_results


def main():
    """Main function - easy configuration example."""
    print()
    print("="*70)
    print("MONTE CARLO & GBM ANALYSIS CONFIGURATION")
    print("="*70)
    print()
    
    # Create configuration
    config = AnalysisConfig()
    
    # ============================================
    # CONFIGURE YOUR ANALYSIS HERE
    # ============================================
    
    # Base assumptions (already set above, modify if needed)
    # config.wacc = 0.08
    # config.rubicon_investment_total = 20_000_000
    # etc.
    
    # Monte Carlo settings
    config.simulations = 5000  # Number of simulations
    
    # Choose your method:
    # Option 1: Use GBM (Geometric Brownian Motion)
    config.use_gbm = True
    config.gbm_drift = 0.03  # 3% expected annual return
    config.gbm_volatility = 0.15  # 15% annual volatility
    
    # Option 2: Use Growth-Rate method (set use_gbm = False)
    # config.use_gbm = False
    # config.price_growth_base = 0.03
    # config.price_growth_std_dev = 0.02
    
    # Volume assumptions (same for both methods)
    config.volume_multiplier_base = 1.0
    config.volume_std_dev = 0.15  # 15% volume volatility
    
    # Other settings
    config.random_seed = 42  # Set to None for random, or a number for reproducibility
    config.data_file = "Analyst_Model_Test_OCC.xlsx"
    config.output_file = "gbm_analysis_results.xlsx"
    
    # ============================================
    # END CONFIGURATION
    # ============================================
    
    # Print configuration
    config.print_config()
    
    # Ask for confirmation
    response = input("Run analysis with these settings? (y/n): ")
    if response.lower() != 'y':
        print("Analysis cancelled.")
        return
    
    print()
    
    # Run analysis
    config.run_analysis()


if __name__ == "__main__":
    main()

