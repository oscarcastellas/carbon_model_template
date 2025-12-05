#!/usr/bin/env python3
"""
Quick Test - Basic DCF and Excel Export (No Monte Carlo)

Run: python quick_test.py
This is faster and tests the core functionality.
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from data.data_loader import DataLoader
from calculators.dcf_calculator import DCFCalculator
from calculators.irr_calculator import IRRCalculator
from calculators.goal_seeker import GoalSeeker
from calculators.sensitivity_analyzer import SensitivityAnalyzer
from calculators.payback_calculator import PaybackCalculator
from reporting.excel_exporter import ExcelExporter
import pandas as pd


def main():
    print("="*70)
    print("QUICK TEST - DCF & EXCEL EXPORT")
    print("="*70)
    print()
    
    # Initialize
    print("1. Initializing...")
    loader = DataLoader()
    irr_calc = IRRCalculator()
    dcf_calc = DCFCalculator(
        wacc=0.08,
        rubicon_investment_total=20_000_000,
        investment_tenor=5,
        irr_calculator=irr_calc
    )
    sensitivity = SensitivityAnalyzer(dcf_calc)
    payback = PaybackCalculator()
    exporter = ExcelExporter()
    print("   ✓ Initialized")
    print()
    
    # Load data
    print("2. Loading data...")
    try:
        data = loader.load_data("Analyst_Model_Test_OCC.xlsx")
        print(f"   ✓ Data loaded: {data.shape}")
        print(f"   Columns: {list(data.columns)}")
        
        # Check if we have the right columns
        required = ['carbon_credits_gross', 'project_implementation_costs', 'base_carbon_price']
        if all(col in data.columns for col in required):
            print(f"   ✓ All required columns found")
            print(f"\n   Sample data (first 3 years):")
            print(data[required].head(3))
        else:
            print(f"   ⚠ Missing columns. Found: {list(data.columns)}")
            return
        print()
    except Exception as e:
        print(f"   ✗ Error: {e}")
        import traceback
        traceback.print_exc()
        return
    
    # Run DCF
    print("3. Running DCF analysis...")
    streaming = 0.48
    try:
        dcf_results = dcf_calc.run_dcf(data, streaming)
        print(f"   ✓ NPV: ${dcf_results['npv']:,.2f}")
        print(f"   ✓ IRR: {dcf_results['irr']:.2%}")
        if pd.isna(dcf_results['irr']):
            print("   ⚠ Warning: IRR calculation returned NaN")
        print()
    except Exception as e:
        print(f"   ✗ Error: {e}")
        import traceback
        traceback.print_exc()
        return
    
    # Goal-seeking
    print("4. Goal-seeking for 20% IRR...")
    try:
        goal_seeker = GoalSeeker(dcf_calc, data)
        goal = goal_seeker.find_target_irr_stream(0.20)
        print(f"   ✓ Required streaming: {goal['streaming_percentage']:.2%}")
        print(f"   ✓ Achieved IRR: {goal['actual_irr']:.2%}")
        streaming = goal['streaming_percentage']
        # Re-run DCF with target streaming
        dcf_results = dcf_calc.run_dcf(data, streaming)
        print()
    except Exception as e:
        print(f"   ⚠ Goal-seeking failed: {e}")
        print("   Continuing with 48% streaming...")
        print()
    
    # Sensitivity (quick test)
    print("5. Running sensitivity analysis...")
    try:
        sens = sensitivity.run_sensitivity_table(
            data, streaming, [0.9, 1.0, 1.1], [0.8, 1.0, 1.2]
        )
        print(f"   ✓ Sensitivity table: {sens.shape}")
        print(f"   Sample (1.0x credit, 1.0x price): IRR = {sens.loc['1.00x', '1.00x']:.2%}")
        print()
    except Exception as e:
        print(f"   ⚠ Sensitivity failed: {e}")
        sens = None
        print()
    
    # Calculate payback
    print("6. Calculating payback period...")
    try:
        payback_period = payback.calculate_payback_period(dcf_results['cash_flows'])
        print(f"   ✓ Payback period: {payback_period:.2f} years")
        print()
    except Exception as e:
        print(f"   ⚠ Payback calculation failed: {e}")
        payback_period = None
        print()
    
    # Export to Excel
    print("7. Exporting to Excel...")
    output_file = "carbon_model_quick_test.xlsx"
    try:
        assumptions = {
            'wacc': 0.08,
            'rubicon_investment_total': 20_000_000,
            'investment_tenor': 5,
            'streaming_percentage_initial': 0.48
        }
        
        exporter.export_model_to_excel(
            output_file,
            assumptions,
            streaming,
            0.20,
            dcf_results['irr'],
            dcf_results['results_df'],
            sens,
            payback_period,
            None  # No Monte Carlo for quick test
        )
        
        print(f"   ✓ Excel export complete!")
        print()
        print("="*70)
        print("SUCCESS!")
        print("="*70)
        print(f"Results exported to: {output_file}")
        print()
        print("The Excel file contains (ALL CALCULATIONS AS FORMULAS):")
        print("  Sheet 1: Inputs & Assumptions (all model inputs)")
        print("  Sheet 2: Valuation Schedule (20-year cash flows with formulas)")
        print("  Sheet 3: Summary & Results (NPV, IRR, Payback - all formulas)")
        print("  Sheet 4: Sensitivity Analysis (IRR sensitivity table)")
        print()
        print("NOTE: All calculations are linked via Excel formulas for full")
        print("      transparency and auditability. Change inputs to see results update!")
        print()
        print("To run full test with Monte Carlo, use: python test_excel.py")
        print("="*70)
        
    except Exception as e:
        print(f"   ✗ Export failed: {e}")
        import traceback
        traceback.print_exc()
        return


if __name__ == "__main__":
    main()

