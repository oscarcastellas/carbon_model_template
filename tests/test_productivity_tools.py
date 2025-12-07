"""
Test script for the new productivity tools:
1. Risk Flagging System
2. Quick Breakeven Calculator
3. Risk Score Calculator
"""

#!/usr/bin/env python3
"""
Test script for the new productivity tools:
1. Risk Flagging System
2. Quick Breakeven Calculator
3. Risk Score Calculator
"""

import sys
import os
import pandas as pd

# Add parent directory to path (to import from root)
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Import modules directly (same pattern as test_excel.py)
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

def test_productivity_tools():
    """Test all three productivity tools."""
    print("=" * 60)
    print("Testing Productivity Tools")
    print("=" * 60)
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
    risk_flagger = RiskFlagger()
    risk_score_calc = RiskScoreCalculator()
    breakeven_calc = BreakevenCalculator(dcf_calc, irr_calc)
    payback_calc = PaybackCalculator()
    print("   ✓ Components initialized")
    print()
    
    # Load data
    print("2. Loading data...")
    data_file = "Analyst_Model_Test_OCC.xlsx"
    if not os.path.exists(data_file):
        print(f"   ERROR: {data_file} not found. Please ensure the file exists.")
        return
    
    data = loader.load_data(data_file)
    print(f"   ✓ Data loaded: {len(data)} years")
    print()
    
    # Run DCF
    print("3. Running DCF analysis...")
    streaming = 0.48
    dcf_results = dcf_calc.run_dcf(data, streaming)
    npv = dcf_results['npv']
    irr = dcf_results['irr']
    print(f"   ✓ NPV: ${npv:,.2f}")
    print(f"   ✓ IRR: {irr:.2%}")
    
    # Calculate payback
    payback = payback_calc.calculate_payback_period(dcf_results['cash_flows'])
    print(f"   ✓ Payback: {payback:.2f} years")
    print()
    
    # Test 1: Risk Flagging
    print("4. Testing Risk Flagging System...")
    risk_flags = risk_flagger.flag_risks(
        irr=irr,
        npv=npv,
        payback_period=payback,
        credit_volumes=data['carbon_credits_gross'],
        project_costs=data['project_implementation_costs']
    )
    print(f"   Risk Level: {risk_flags['risk_level'].upper()}")
    print(f"   Red Flags: {risk_flags['flag_count']['red']}")
    print(f"   Yellow Flags: {risk_flags['flag_count']['yellow']}")
    print(f"   Green Indicators: {risk_flags['flag_count']['green']}")
    print()
    print("   Risk Summary:")
    print(risk_flagger.get_risk_summary(risk_flags))
    print()
    
    # Test 2: Risk Score
    print("5. Testing Risk Score Calculator...")
    risk_score = risk_score_calc.calculate_overall_risk_score(
        irr=irr,
        npv=npv,
        payback_period=payback,
        credit_volumes=data['carbon_credits_gross'],
        base_prices=data['base_carbon_price'],
        project_costs=data['project_implementation_costs'],
        total_investment=20_000_000
    )
    print(f"   Overall Risk Score: {risk_score['overall_risk_score']}/100")
    print(f"   Risk Category: {risk_score['risk_category']}")
    print(f"   Financial Risk: {risk_score['financial_risk']}/100")
    print(f"   Volume Risk: {risk_score['volume_risk']}/100")
    print(f"   Price Risk: {risk_score['price_risk']}/100")
    print(f"   Operational Risk: {risk_score['operational_risk']}/100")
    print()
    
    # Test 3: Breakeven Calculator
    print("6. Testing Breakeven Calculator...")
    print("   Calculating breakeven price...")
    be_price = breakeven_calc.calculate_breakeven_price(data, streaming, target_npv=0.0)
    if be_price and 'breakeven_price' in be_price and not pd.isna(be_price.get('breakeven_price')):
        print(f"   ✓ Breakeven Price: ${be_price['breakeven_price']:,.2f}/ton")
        print(f"   ✓ Base Price: ${be_price['base_price']:,.2f}/ton")
        print(f"   ✓ Multiplier: {be_price['price_multiplier']:.2f}x")
    else:
        print("   ⚠ Could not calculate breakeven price")
    print()
    
    print("   Calculating breakeven volume...")
    be_volume = breakeven_calc.calculate_breakeven_volume(data, streaming, target_npv=0.0)
    if be_volume and 'breakeven_volume_multiplier' in be_volume and not pd.isna(be_volume.get('breakeven_volume_multiplier')):
        print(f"   ✓ Breakeven Volume Multiplier: {be_volume['breakeven_volume_multiplier']:.2%}")
    else:
        print("   ⚠ Could not calculate breakeven volume")
    print()
    
    print("   Calculating breakeven streaming...")
    be_streaming = breakeven_calc.calculate_breakeven_streaming(data, target_npv=0.0)
    if be_streaming and 'breakeven_streaming' in be_streaming and not pd.isna(be_streaming.get('breakeven_streaming')):
        print(f"   ✓ Breakeven Streaming: {be_streaming['breakeven_streaming']:.2%}")
    else:
        print("   ⚠ Could not calculate breakeven streaming")
    print()
    
    print("=" * 60)
    print("All tests completed!")
    print("=" * 60)

if __name__ == "__main__":
    import pandas as pd
    test_productivity_tools()

