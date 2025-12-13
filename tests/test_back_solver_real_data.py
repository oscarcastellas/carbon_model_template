#!/usr/bin/env python3
"""
Test Back-Solver with Real Project Data

Tests the Streaming Deal Valuation Back-Solver with actual project data
and verifies Excel export includes the results.
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Import modules directly (same approach as generate_full_excel.py)
from data.loader import DataLoader
from core.dcf import DCFCalculator
from core.irr import IRRCalculator
from valuation.deal_valuation import DealValuationSolver
from export.excel import ExcelExporter
import pandas as pd


def main():
    print("=" * 70)
    print("TESTING BACK-SOLVER WITH REAL PROJECT DATA")
    print("=" * 70)
    print()
    
    # Initialize components
    print("1. Initializing components...")
    loader = DataLoader()
    irr_calc = IRRCalculator()
    dcf_calc = DCFCalculator(
        wacc=0.08,
        rubicon_investment_total=20_000_000,
        investment_tenor=5,
        irr_calculator=irr_calc
    )
    solver = None
    exporter = ExcelExporter()
    print("   ✓ Components initialized")
    print()
    
    # Load real data
    print("2. Loading real project data...")
    data_file = "Analyst_Model_Test_OCC.xlsx"
    if not os.path.exists(data_file):
        print(f"   ✗ ERROR: {data_file} not found!")
        print(f"   Looking in: {os.getcwd()}")
        return
    
    data = loader.load_data(data_file)
    print(f"   ✓ Data loaded: {len(data)} years")
    print()
    
    # Run DCF first
    print("3. Running initial DCF analysis...")
    streaming_initial = 0.48
    dcf_results = dcf_calc.run_dcf(data, streaming_initial)
    print(f"   ✓ NPV: ${dcf_results['npv']:,.2f}")
    print(f"   ✓ IRR: {dcf_results['irr']:.2%}")
    print()
    
    # Initialize solver
    solver = DealValuationSolver(
        dcf_calculator=dcf_calc,
        data=data,
        tolerance=1e-4
    )
    
    # Test 1: Solve for Purchase Price given Target IRR
    print("4. Testing: Solve for Purchase Price (Target IRR = 20%)...")
    deal_valuation_results = None
    try:
        price_results = solver.solve_for_purchase_price(
            target_irr=0.20,
            streaming_percentage=0.48
        )
        deal_valuation_results = price_results
        print(f"   ✓ Maximum Purchase Price: ${price_results['purchase_price']:,.2f}")
        print(f"   ✓ Actual IRR Achieved: {price_results['actual_irr']:.2%}")
        print(f"   ✓ Target IRR: {price_results['target_irr']:.2%}")
        print(f"   ✓ Difference: {price_results['difference']:.4%}")
        print(f"   ✓ NPV at Calculated Price: ${price_results['npv']:,.2f}")
        print()
    except Exception as e:
        print(f"   ✗ Error: {e}")
        import traceback
        traceback.print_exc()
        print()
    
    # Test 2: Calculate IRR from Purchase Price
    print("5. Testing: Calculate IRR from Purchase Price ($20M)...")
    try:
        irr_results = solver.solve_for_project_irr(
            purchase_price=20_000_000,
            streaming_percentage=0.48
        )
        print(f"   ✓ Purchase Price: ${irr_results['purchase_price']:,.2f}")
        print(f"   ✓ Project IRR: {irr_results['irr']:.2%}")
        print(f"   ✓ NPV: ${irr_results['npv']:,.2f}")
        print()
    except Exception as e:
        print(f"   ✗ Error: {e}")
        print()
    
    # Test 3: Solve for Streaming % given Price + IRR
    print("6. Testing: Solve for Streaming % (Price=$15M, Target IRR=20%)...")
    try:
        streaming_results = solver.solve_for_streaming_given_price(
            purchase_price=15_000_000,
            target_irr=0.20
        )
        print(f"   ✓ Purchase Price: ${streaming_results['purchase_price']:,.2f}")
        print(f"   ✓ Required Streaming %: {streaming_results['streaming_percentage']:.2%}")
        print(f"   ✓ Actual IRR: {streaming_results['actual_irr']:.2%}")
        print(f"   ✓ Target IRR: {streaming_results['target_irr']:.2%}")
        print(f"   ✓ NPV: ${streaming_results['npv']:,.2f}")
        print()
    except Exception as e:
        print(f"   ✗ Error: {e}")
        print()
    
    # Export to Excel with back-solver results
    print("7. Exporting to Excel with back-solver results...")
    output_file = "test_back_solver_output.xlsx"
    try:
        # Get assumptions
        assumptions = {
            'wacc': 0.08,
            'rubicon_investment_total': 20_000_000,
            'investment_tenor': 5,
            'streaming_percentage_initial': 0.48
        }
        
        exporter.export_model_to_excel(
            filename=output_file,
            assumptions=assumptions,
            target_streaming_percentage=streaming_initial,
            target_irr=0.20,
            actual_irr=dcf_results['irr'],
            valuation_schedule=dcf_results['results_df'],
            sensitivity_table=None,
            payback_period=None,
            monte_carlo_results=None,
            risk_flags=None,
            risk_score=None,
            breakeven_results=None,
            deal_valuation_results=deal_valuation_results
        )
        print(f"   ✓ Excel file created: {output_file}")
        print(f"   ✓ Deal Valuation sheet should be included")
        print()
        
        # Verify file exists
        if os.path.exists(output_file):
            print("   ✓ File verification: PASSED")
            file_size = os.path.getsize(output_file) / 1024  # KB
            print(f"   ✓ File size: {file_size:.2f} KB")
        else:
            print("   ✗ File verification: FAILED")
    except Exception as e:
        print(f"   ✗ Error exporting to Excel: {e}")
        import traceback
        traceback.print_exc()
        print()
    
    print("=" * 70)
    print("TEST COMPLETE")
    print("=" * 70)
    print()
    print("Next steps:")
    print("  1. Open test_back_solver_output.xlsx")
    print("  2. Check for 'Deal Valuation' sheet")
    print("  3. Verify results match console output")


if __name__ == '__main__':
    main()

