#!/usr/bin/env python3
"""
Generate Excel Output with Embedded Charts

This script generates a complete Excel file with:
1. All volatility analysis charts embedded
2. Interactive analysis sheet for parameter adjustment
3. Full model results
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from analysis_config import AnalysisConfig
from export.excel import ExcelExporter
from export.chart_exporter import ChartExporter
from export.interactive_sheet import InteractiveSheetCreator
from data.loader import DataLoader
from core.dcf import DCFCalculator
from core.irr import IRRCalculator
from analysis.monte_carlo import MonteCarloSimulator
from analysis.gbm_simulator import GBMPriceSimulator
from risk.flagger import RiskFlagger
from risk.scorer import RiskScoreCalculator
from valuation.breakeven import BreakevenCalculator
from core.payback import PaybackCalculator
from analysis.volatility_visualizer import VolatilityVisualizer
import pandas as pd


def main():
    """Generate Excel with charts and interactive sheet."""
    print()
    print("="*70)
    print("GENERATING EXCEL WITH CHARTS & INTERACTIVE ANALYSIS")
    print("="*70)
    print()
    
    # Configure analysis
    config = AnalysisConfig()
    config.use_gbm = True
    config.gbm_drift = 0.03
    config.gbm_volatility = 0.15
    config.simulations = 5000
    config.random_seed = 42
    config.data_file = "Analyst_Model_Test_OCC.xlsx"
    config.output_file = "carbon_model_with_charts.xlsx"
    
    print("Configuration:")
    config.print_config()
    print()
    
    # Step 1: Generate charts first
    print("Step 1: Generating volatility charts...")
    visualizer = VolatilityVisualizer(output_dir="volatility_charts")
    
    # Load data
    loader = DataLoader()
    data = loader.load_data(config.data_file)
    base_prices = data['base_carbon_price']
    
    # Generate GBM paths for visualization
    gbm_sim = GBMPriceSimulator()
    gbm_paths = []
    for i in range(1000):
        path = gbm_sim.generate_gbm_path_from_base(
            base_prices=base_prices,
            drift=config.gbm_drift,
            volatility=config.gbm_volatility,
            random_seed=None
        )
        gbm_paths.append(path)
    
    print("   ✓ Generated GBM price paths")
    
    # Run Monte Carlo
    print("Step 2: Running Monte Carlo analysis...")
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
    print()
    
    # Generate all charts
    print("Step 3: Generating all charts...")
    saved_charts = visualizer.generate_full_report(
        base_prices=base_prices,
        gbm_paths=gbm_paths,
        monte_carlo_results=mc_results,
        output_prefix="carbon_price_volatility"
    )
    print(f"   ✓ Generated {len(saved_charts)} charts")
    print()
    
    # Step 4: Run DCF and other analyses
    print("Step 4: Running DCF and risk analysis...")
    dcf_results = dcf_calc.run_dcf(data, config.streaming_percentage_initial)
    payback_calc = PaybackCalculator()
    payback = payback_calc.calculate_payback_period(dcf_results['cash_flows'])
    
    risk_flagger = RiskFlagger()
    risk_flags = risk_flagger.flag_risks(
        dcf_results['irr'],
        dcf_results['npv'],
        payback,
        credit_volumes=data['carbon_credits_gross'],
        project_costs=data['project_implementation_costs']
    )
    
    risk_scorer = RiskScoreCalculator()
    risk_score = risk_scorer.calculate_overall_risk_score(
        dcf_results['irr'],
        dcf_results['npv'],
        payback,
        credit_volumes=data['carbon_credits_gross'],
        base_prices=data['base_carbon_price'],
        project_costs=data['project_implementation_costs'],
        total_investment=config.rubicon_investment_total
    )
    
    breakeven_calc = BreakevenCalculator(dcf_calc, irr_calc)
    breakeven = breakeven_calc.calculate_all_breakevens(
        data, config.streaming_percentage_initial, 0.0
    )
    
    print("   ✓ All analyses complete")
    print()
    
    # Step 5: Create Excel with charts
    print("Step 5: Creating Excel file with embedded charts...")
    
    # Prepare assumptions
    assumptions = {
        'wacc': config.wacc,
        'rubicon_investment_total': config.rubicon_investment_total,
        'investment_tenor': config.investment_tenor,
        'streaming_percentage_initial': config.streaming_percentage_initial,
        'price_growth_base': config.price_growth_base,
        'price_growth_std_dev': config.price_growth_std_dev,
        'volume_multiplier_base': config.volume_multiplier_base,
        'volume_std_dev': config.volume_std_dev,
        'use_gbm': config.use_gbm,
        'gbm_drift': config.gbm_drift,
        'gbm_volatility': config.gbm_volatility,
        'simulations': config.simulations
    }
    
    # Create Excel exporter
    excel_exporter = ExcelExporter()
    
    # Create workbook manually to add charts
    import xlsxwriter
    workbook = xlsxwriter.Workbook(config.output_file, {'nan_inf_to_errors': True})
    
    # Create standard sheets first
    formats = excel_exporter._create_formats(workbook)
    
    # Sheet 1: Inputs & Assumptions
    inputs_sheet = workbook.add_worksheet('Inputs & Assumptions')
    excel_exporter._write_inputs_sheet(
        workbook, inputs_sheet, formats, assumptions,
        config.streaming_percentage_initial, 0.20
    )
    
    # Sheet 2: Valuation Schedule
    valuation_sheet = workbook.add_worksheet('Valuation Schedule')
    excel_exporter._write_valuation_schedule_with_formulas(
        workbook, valuation_sheet, formats, dcf_results['results_df'], inputs_sheet
    )
    
    # Sheet 3: Summary & Results
    summary_sheet = workbook.add_worksheet('Summary & Results')
    excel_exporter._write_summary_results_sheet(
        workbook, summary_sheet, formats, dcf_results['results_df'],
        inputs_sheet, dcf_results['irr'], payback, mc_results,
        risk_flags, risk_score, breakeven
    )
    
    # Sheet 4: Monte Carlo Results
    mc_sheet = workbook.add_worksheet('Monte Carlo Results')
    excel_exporter._write_monte_carlo_sheet(
        workbook, mc_sheet, formats, mc_results
    )
    
    # Sheet 5: Interactive Analysis (NEW)
    print("   Creating interactive analysis sheet...")
    interactive_creator = InteractiveSheetCreator(workbook)
    interactive_sheet = interactive_creator.create_interactive_analysis_sheet(
        base_assumptions=assumptions,
        monte_carlo_results=mc_results,
        sheet_name="Interactive Analysis"
    )
    
    # Sheet 6: Volatility Charts (NEW)
    print("   Embedding volatility charts...")
    chart_exporter = ChartExporter(workbook)
    charts_sheet = chart_exporter.create_charts_sheet(
        charts=saved_charts,
        sheet_name="Volatility Charts"
    )
    
    workbook.close()
    
    print()
    print("="*70)
    print("EXCEL FILE GENERATED SUCCESSFULLY!")
    print("="*70)
    print()
    print(f"📊 File: {config.output_file}")
    print()
    print("Sheets included:")
    print("  1. Inputs & Assumptions")
    print("  2. Valuation Schedule")
    print("  3. Summary & Results")
    print("  4. Monte Carlo Results")
    print("  5. Interactive Analysis ⭐ NEW - Adjust parameters here!")
    print("  6. Volatility Charts ⭐ NEW - All charts embedded!")
    print()
    print("How to use Interactive Analysis sheet:")
    print("  1. Open the Excel file")
    print("  2. Go to 'Interactive Analysis' sheet")
    print("  3. Adjust parameters (GBM drift, volatility, etc.)")
    print("  4. For full simulation, run: python3 examples/run_interactive_analysis.py")
    print()
    print("="*70)


if __name__ == "__main__":
    main()

