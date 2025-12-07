#!/usr/bin/env python3
"""
Generate Full Excel Export with All New Productivity Modules

This script runs the complete pipeline including:
- DCF Analysis
- Goal-Seeking
- Risk Flagging
- Risk Scoring
- Breakeven Analysis
- Monte Carlo (optional)
- Full Excel Export
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Import modules directly
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


def main():
    print("="*70)
    print("GENERATING FULL EXCEL EXPORT WITH ALL PRODUCTIVITY MODULES")
    print("="*70)
    print()
    
    # 1. Initialize components
    print("1. Initializing components...")
    loader = DataLoader()
    irr_calc = IRRCalculator()
    dcf_calc = DCFCalculator(
        wacc=0.08,
        rubicon_investment_total=20_000_000,
        investment_tenor=5,
        irr_calculator=irr_calc
    )
    goal_seeker = None
    sensitivity = SensitivityAnalyzer(dcf_calc)
    payback = PaybackCalculator()
    mc_sim = MonteCarloSimulator(dcf_calc, irr_calc)
    risk_flagger = RiskFlagger()
    risk_score_calc = RiskScoreCalculator()
    breakeven_calc = BreakevenCalculator(dcf_calc, irr_calc)
    exporter = ExcelExporter()
    print("   ✓ All components initialized")
    print()
    
    # 2. Load data
    print("2. Loading test data...")
    data_file = "Analyst_Model_Test_OCC.xlsx"
    if not os.path.exists(data_file):
        print(f"   ✗ ERROR: {data_file} not found!")
        return
    
    data = loader.load_data(data_file)
    print(f"   ✓ Data loaded: {len(data)} years")
    print()
    
    # 3. Run DCF with initial streaming
    print("3. Running DCF analysis...")
    streaming_initial = 0.48
    dcf_results = dcf_calc.run_dcf(data, streaming_initial)
    npv = dcf_results['npv']
    irr = dcf_results['irr']
    payback_period = payback.calculate_payback_period(dcf_results['cash_flows'])
    print(f"   ✓ NPV: ${npv:,.2f}")
    print(f"   ✓ IRR: {irr:.2%}")
    print(f"   ✓ Payback: {payback_period:.2f} years")
    print()
    
    # 4. Run goal-seeking to find target streaming
    print("4. Running goal-seeking (target IRR = 20%)...")
    goal_seeker = GoalSeeker(dcf_calc, data)
    try:
        goal_results = goal_seeker.find_target_irr_stream(0.20)
        target_streaming = goal_results['streaming_percentage']
        target_irr = goal_results['actual_irr']
        print(f"   ✓ Target Streaming: {target_streaming:.2%}")
        print(f"   ✓ Actual IRR: {target_irr:.2%}")
        print()
        
        # Re-run DCF with target streaming
        dcf_results_target = dcf_calc.run_dcf(data, target_streaming)
        npv_target = dcf_results_target['npv']
        irr_target = dcf_results_target['irr']
        payback_target = payback.calculate_payback_period(dcf_results_target['cash_flows'])
    except ValueError as e:
        print(f"   ⚠ Goal-seeking not feasible: {e}")
        print("   Using initial streaming percentage instead...")
        target_streaming = streaming_initial
        target_irr = irr
        dcf_results_target = dcf_results
        npv_target = npv
        irr_target = irr
        payback_target = payback_period
        print(f"   Using streaming: {target_streaming:.2%} (IRR: {irr_target:.2%})")
        print()
    
    # 5. Calculate Risk Flags
    print("5. Calculating risk flags...")
    risk_flags = risk_flagger.flag_risks(
        irr=irr_target,
        npv=npv_target,
        payback_period=payback_target,
        credit_volumes=data['carbon_credits_gross'],
        project_costs=data['project_implementation_costs']
    )
    print(f"   ✓ Risk Level: {risk_flags['risk_level'].upper()}")
    print(f"   ✓ Red Flags: {risk_flags['flag_count']['red']}")
    print(f"   ✓ Yellow Flags: {risk_flags['flag_count']['yellow']}")
    print()
    
    # 6. Calculate Risk Score
    print("6. Calculating risk score...")
    risk_score = risk_score_calc.calculate_overall_risk_score(
        irr=irr_target,
        npv=npv_target,
        payback_period=payback_target,
        credit_volumes=data['carbon_credits_gross'],
        base_prices=data['base_carbon_price'],
        project_costs=data['project_implementation_costs'],
        total_investment=20_000_000
    )
    print(f"   ✓ Overall Risk Score: {risk_score['overall_risk_score']}/100")
    print(f"   ✓ Risk Category: {risk_score['risk_category']}")
    print()
    
    # 7. Calculate Breakeven
    print("7. Calculating breakeven points...")
    breakeven_results = breakeven_calc.calculate_all_breakevens(
        data, target_streaming, target_npv=0.0
    )
    print("   ✓ Breakeven calculations complete")
    if 'breakeven_price' in breakeven_results:
        bp = breakeven_results['breakeven_price']
        if bp and not pd.isna(bp.get('breakeven_price')):
            print(f"     - Breakeven Price: ${bp['breakeven_price']:,.2f}/ton")
    if 'breakeven_volume' in breakeven_results:
        bv = breakeven_results['breakeven_volume']
        if bv and not pd.isna(bv.get('breakeven_volume_multiplier')):
            print(f"     - Breakeven Volume: {bv['breakeven_volume_multiplier']:.2%}")
    if 'breakeven_streaming' in breakeven_results:
        bs = breakeven_results['breakeven_streaming']
        if bs and not pd.isna(bs.get('breakeven_streaming')):
            print(f"     - Breakeven Streaming: {bs['breakeven_streaming']:.2%}")
    print()
    
    # 8. Run Sensitivity Analysis
    print("8. Running sensitivity analysis...")
    credit_range = [0.8, 0.9, 1.0, 1.1, 1.2]
    price_range = [0.7, 0.85, 1.0, 1.15, 1.3]
    sensitivity_table = sensitivity.run_sensitivity_table(
        data, target_streaming, credit_range, price_range
    )
    print("   ✓ Sensitivity table generated")
    print()
    
    # 9. Run Monte Carlo (optional - takes a bit longer)
    print("9. Running Monte Carlo simulation (5,000 simulations)...")
    print("   (This may take 1-2 minutes...)")
    mc_results = mc_sim.run_monte_carlo(
        base_data=data,
        streaming_percentage=target_streaming,
        price_growth_base=0.03,
        price_growth_std_dev=0.02,
        volume_multiplier_base=1.0,
        volume_std_dev=0.15,
        simulations=5000
    )
    print(f"   ✓ Monte Carlo complete!")
    print(f"     - Mean IRR: {mc_results['mc_mean_irr']:.2%}")
    print(f"     - P10 IRR: {mc_results['mc_p10_irr']:.2%}")
    print(f"     - P90 IRR: {mc_results['mc_p90_irr']:.2%}")
    print()
    
    # Update risk flags with Monte Carlo volatility
    print("10. Updating risk assessment with Monte Carlo volatility...")
    irr_series = mc_results.get('irr_series', [])
    irr_volatility = None
    if len(irr_series) > 0:
        import numpy as np
        irr_valid = [x for x in irr_series if pd.notna(x)]
        if len(irr_valid) > 0:
            irr_volatility = np.std(irr_valid)
    
    risk_flags_updated = risk_flagger.flag_risks(
        irr=irr_target,
        npv=npv_target,
        payback_period=payback_target,
        irr_volatility=irr_volatility,
        credit_volumes=data['carbon_credits_gross'],
        project_costs=data['project_implementation_costs']
    )
    
    risk_score_updated = risk_score_calc.calculate_overall_risk_score(
        irr=irr_target,
        npv=npv_target,
        payback_period=payback_target,
        credit_volumes=data['carbon_credits_gross'],
        base_prices=data['base_carbon_price'],
        project_costs=data['project_implementation_costs'],
        price_volatility=0.02,  # From Monte Carlo assumptions
        volume_volatility=0.15,  # From Monte Carlo assumptions
        total_investment=20_000_000
    )
    print("   ✓ Risk assessment updated")
    print()
    
    # 11. Export to Excel
    print("11. Exporting to Excel with all modules...")
    output_file = "Full_Model_Export_With_Productivity_Tools.xlsx"
    
    assumptions = {
        'wacc': 0.08,
        'rubicon_investment_total': 20_000_000,
        'investment_tenor': 5,
        'streaming_percentage_initial': streaming_initial
    }
    
    exporter.export_model_to_excel(
        filename=output_file,
        assumptions=assumptions,
        target_streaming_percentage=target_streaming,
        target_irr=0.20,
        actual_irr=irr_target,
        valuation_schedule=dcf_results_target['results_df'],
        sensitivity_table=sensitivity_table,
        payback_period=payback_target,
        monte_carlo_results=mc_results,
        risk_flags=risk_flags_updated,
        risk_score=risk_score_updated,
        breakeven_results=breakeven_results
    )
    
    print(f"   ✓ Excel file created: {output_file}")
    print()
    print("="*70)
    print("EXPORT COMPLETE!")
    print("="*70)
    print()
    print(f"📊 Open '{output_file}' to view:")
    print("   • Sheet 1: Inputs & Assumptions")
    print("   • Sheet 2: Valuation Schedule (with formulas)")
    print("   • Sheet 3: Summary & Results")
    print("      - Risk Assessment (NEW!)")
    print("      - Risk Score (NEW!)")
    print("      - Breakeven Analysis (NEW!)")
    print("      - Monte Carlo Summary")
    print("   • Sheet 4: Sensitivity Analysis")
    print("   • Sheet 5: Monte Carlo Results (with histograms)")
    print()
    print("All new productivity modules are included in Sheet 3!")


if __name__ == "__main__":
    main()

