#!/usr/bin/env python3
"""
Simple Test Script - Excel Extraction and Full Pipeline

Run: python test_excel.py
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
from export.excel import ExcelExporter
import pandas as pd


def main():
    print("="*70)
    print("CARBON MODEL - EXCEL EXTRACTION TEST")
    print("="*70)
    print()
    
    # 1. Initialize components
    print("1. Initializing...")
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
    exporter = ExcelExporter()
    print("   ✓ Components initialized")
    print()
    
    # 2. Load data
    print("2. Loading data...")
    try:
        data = loader.load_data("Analyst_Model_Test_OCC.xlsx")
        print(f"   ✓ Data loaded: {data.shape}")
        print(f"   Columns: {list(data.columns)}")
        print(f"\n   Sample data:")
        print(data[['carbon_credits_gross', 'project_implementation_costs', 'base_carbon_price']].head())
        print()
    except Exception as e:
        print(f"   ✗ Error: {e}")
        import traceback
        traceback.print_exc()
        return
    
    # 3. Run DCF
    print("3. Running DCF...")
    streaming = 0.48
    dcf_results = dcf_calc.run_dcf(data, streaming)
    print(f"   ✓ NPV: ${dcf_results['npv']:,.2f}")
    print(f"   ✓ IRR: {dcf_results['irr']:.2%}")
    print()
    
    # 4. Goal-seeking
    print("4. Goal-seeking for 20% IRR...")
    goal_seeker = GoalSeeker(dcf_calc, data)
    try:
        goal = goal_seeker.find_target_irr_stream(0.20)
        print(f"   ✓ Required streaming: {goal['streaming_percentage']:.2%}")
        streaming = goal['streaming_percentage']
        dcf_results = dcf_calc.run_dcf(data, streaming)
        print()
    except Exception as e:
        print(f"   ⚠ {e}")
        print()
    
    # 5. Sensitivity
    print("5. Sensitivity analysis...")
    try:
        sens = sensitivity.run_sensitivity_table(
            data, streaming, [0.9, 1.0, 1.1], [0.8, 1.0, 1.2]
        )
        print(f"   ✓ Complete: {sens.shape}")
        print()
    except Exception as e:
        print(f"   ⚠ {e}")
        sens = None
        print()
    
    # 6. Monte Carlo
    print("6. Monte Carlo (5000 runs - may take 2-5 minutes)...")
    try:
        mc = mc_sim.run_monte_carlo(
            data, streaming, 0.03, 0.02, 1.0, 0.15, 5000, 42
        )
        print(f"   ✓ Mean IRR: {mc['mc_mean_irr']:.2%}")
        print(f"   ✓ P10 IRR: {mc['mc_p10_irr']:.2%}")
        print()
    except Exception as e:
        print(f"   ⚠ {e}")
        mc = None
        print()
    
    # 7. Export
    print("7. Exporting to Excel...")
    try:
        assumptions = {
            'wacc': 0.08,
            'rubicon_investment_total': 20_000_000,
            'investment_tenor': 5,
            'streaming_percentage_initial': 0.48
        }
        payback_period = payback.calculate_payback_period(dcf_results['cash_flows'])
        
        exporter.export_model_to_excel(
            "carbon_model_results.xlsx",
            assumptions,
            streaming,
            0.20,
            dcf_results['irr'],
            dcf_results['results_df'],
            sens,
            payback_period,
            mc
        )
        print("   ✓ Export complete: carbon_model_results.xlsx")
        print()
        print("="*70)
        print("SUCCESS! Check carbon_model_results.xlsx")
        print("="*70)
        print()
        print("The Excel file contains (ALL CALCULATIONS AS FORMULAS):")
        print("  Sheet 1: Inputs & Assumptions (all model inputs)")
        print("  Sheet 2: Valuation Schedule (20-year cash flows with formulas)")
        print("  Sheet 3: Summary & Results (NPV, IRR, Payback - all formulas)")
        print("  Sheet 4: Sensitivity Analysis (IRR sensitivity table)")
        print("  Sheet 5: Monte Carlo Results (full simulation + histogram)")
        print()
        print("NOTE: All calculations are linked via Excel formulas for full")
        print("      transparency and auditability. Change inputs to see results update!")
        print("="*70)
    except Exception as e:
        print(f"   ✗ Error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()

