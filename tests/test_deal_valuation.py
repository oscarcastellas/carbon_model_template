"""
Unit tests for Deal Valuation Solver module.
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pandas as pd
import numpy as np
from core.dcf import DCFCalculator
from core.irr import IRRCalculator
from valuation.deal_valuation import DealValuationSolver


def create_test_data():
    """Create test data for DCF calculations."""
    years = range(1, 21)
    data = pd.DataFrame({
        'year': years,
        'carbon_credits_gross': [10000 * (1.05 ** (y-1)) for y in years],
        'base_carbon_price': [50.0] * 20,
        'project_implementation_costs': [100000] * 20
    })
    data.set_index('year', inplace=True)
    return data


def test_solve_for_purchase_price():
    """Test solving for purchase price given target IRR."""
    print("Testing solve_for_purchase_price...")
    
    # Setup
    data = create_test_data()
    dcf_calculator = DCFCalculator(
        wacc=0.08,
        rubicon_investment_total=20_000_000,
        investment_tenor=5
    )
    
    solver = DealValuationSolver(
        dcf_calculator=dcf_calculator,
        data=data,
        tolerance=1e-4
    )
    
    # Test
    results = solver.solve_for_purchase_price(
        target_irr=0.20,
        streaming_percentage=0.48,
        investment_tenor=5
    )
    
    # Assert
    assert 'purchase_price' in results
    assert 'actual_irr' in results
    assert 'target_irr' in results
    assert results['purchase_price'] > 0
    assert abs(results['actual_irr'] - 0.20) < 0.01  # Within 1% tolerance
    
    print(f"✓ Purchase price: ${results['purchase_price']:,.2f}")
    print(f"✓ Actual IRR: {results['actual_irr']:.2%}")
    print(f"✓ Target IRR: {results['target_irr']:.2%}")
    print("✓ Test passed!\n")


def test_solve_for_project_irr():
    """Test calculating IRR from purchase price."""
    print("Testing solve_for_project_irr...")
    
    # Setup
    data = create_test_data()
    dcf_calculator = DCFCalculator(
        wacc=0.08,
        rubicon_investment_total=20_000_000,
        investment_tenor=5
    )
    
    solver = DealValuationSolver(
        dcf_calculator=dcf_calculator,
        data=data,
        tolerance=1e-4
    )
    
    # Test
    results = solver.solve_for_project_irr(
        purchase_price=20_000_000,
        streaming_percentage=0.48,
        investment_tenor=5
    )
    
    # Assert
    assert 'purchase_price' in results
    assert 'irr' in results
    assert 'npv' in results
    assert results['purchase_price'] == 20_000_000
    assert isinstance(results['irr'], float)
    assert not np.isnan(results['irr'])
    
    print(f"✓ Purchase price: ${results['purchase_price']:,.2f}")
    print(f"✓ Project IRR: {results['irr']:.2%}")
    print(f"✓ NPV: ${results['npv']:,.2f}")
    print("✓ Test passed!\n")


def test_solve_for_streaming_given_price():
    """Test solving for streaming percentage given price and IRR."""
    print("Testing solve_for_streaming_given_price...")
    
    # Setup
    data = create_test_data()
    dcf_calculator = DCFCalculator(
        wacc=0.08,
        rubicon_investment_total=20_000_000,
        investment_tenor=5
    )
    
    solver = DealValuationSolver(
        dcf_calculator=dcf_calculator,
        data=data,
        tolerance=1e-4
    )
    
    # Test
    results = solver.solve_for_streaming_given_price(
        purchase_price=20_000_000,
        target_irr=0.20,
        investment_tenor=5
    )
    
    # Assert
    assert 'streaming_percentage' in results
    assert 'purchase_price' in results
    assert 'actual_irr' in results
    assert 'target_irr' in results
    assert 0 <= results['streaming_percentage'] <= 1
    assert abs(results['actual_irr'] - 0.20) < 0.01  # Within 1% tolerance
    
    print(f"✓ Purchase price: ${results['purchase_price']:,.2f}")
    print(f"✓ Required streaming %: {results['streaming_percentage']:.2%}")
    print(f"✓ Actual IRR: {results['actual_irr']:.2%}")
    print(f"✓ Target IRR: {results['target_irr']:.2%}")
    print("✓ Test passed!\n")


if __name__ == '__main__':
    print("=" * 60)
    print("Deal Valuation Solver - Unit Tests")
    print("=" * 60)
    print()
    
    try:
        test_solve_for_purchase_price()
        test_solve_for_project_irr()
        test_solve_for_streaming_given_price()
        
        print("=" * 60)
        print("All tests passed! ✓")
        print("=" * 60)
    except Exception as e:
        print(f"\n✗ Test failed with error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

