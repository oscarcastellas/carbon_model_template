"""
Example usage of CarbonModelGenerator class.

This script demonstrates how to use the CarbonModelGenerator for DCF analysis
and goal-seeking for target IRR scenarios.

The package is now modular, allowing you to use individual components:
- DataLoader: For data ingestion
- DCFCalculator: For financial calculations
- IRRCalculator: For IRR calculations
- GoalSeeker: For optimization
- CarbonModelGenerator: Main orchestrator class
"""

from carbon_model_generator import CarbonModelGenerator
# You can also import individual modules if needed:
# from data.data_loader import DataLoader
# from calculators.dcf_calculator import DCFCalculator
# from calculators.irr_calculator import IRRCalculator
# from calculators.goal_seeker import GoalSeeker
import pandas as pd


def main():
    """
    Example workflow using CarbonModelGenerator.
    """
    # Initialize the model with assumptions
    assumptions = {
        'wacc': 0.08,  # 8.0% discount rate
        'rubicon_investment_total': 20_000_000,  # $20M initial investment
        'investment_tenor': 5,  # 5 years deployment
        'streaming_percentage_initial': 0.48  # 48% initial streaming
    }
    
    # Create model instance
    model = CarbonModelGenerator(**assumptions)
    
    # Load data from file
    # Note: Update this path to match your actual data file
    data_file = "Analyst_Model_Test_OCC.xlsx"
    
    try:
        print("Loading data...")
        data = model.load_data(data_file)
        print(f"Data loaded successfully. Shape: {data.shape}")
        print(f"\nFirst few rows:")
        print(data.head())
        print(f"\nData columns: {list(data.columns)}")
        
        # Run DCF analysis with initial streaming percentage
        print("\n" + "="*60)
        print("Running DCF Analysis...")
        print("="*60)
        dcf_results = model.run_dcf()
        
        print(f"\nNPV: ${dcf_results['npv']:,.2f}")
        print(f"IRR: {dcf_results['irr']:.2%}")
        print(f"Streaming Percentage: {dcf_results['streaming_percentage']:.2%}")
        
        # Demonstrate explicit NPV and IRR calculation methods
        print("\n" + "-"*60)
        print("Explicit NPV and IRR Calculations:")
        print("-"*60)
        npv_explicit = model.calculate_npv()
        irr_explicit = model.calculate_irr()
        print(f"NPV (explicit): ${npv_explicit:,.2f}")
        print(f"IRR (explicit): {irr_explicit:.2%}")
        print(f"Verification - NPV matches: {abs(npv_explicit - dcf_results['npv']) < 0.01}")
        print(f"Verification - IRR matches: {abs(irr_explicit - dcf_results['irr']) < 0.0001}")
        
        # Display summary of results
        print("\n" + "="*60)
        print("DCF Results Summary (First 5 Years):")
        print("="*60)
        summary_cols = [
            'carbon_credits_gross',
            'rubicon_share_credits',
            'base_carbon_price',
            'rubicon_revenue',
            'rubicon_investment_cf',
            'rubicon_net_cash_flow',
            'present_value'
        ]
        print(dcf_results['results_df'][summary_cols].head())
        
        # Goal-seeking: Find streaming percentage for 20% IRR
        print("\n" + "="*60)
        print("Goal-Seeking: Finding Streaming Percentage for 20% IRR")
        print("="*60)
        target_irr = 0.20  # 20%
        
        try:
            goal_seek_results = model.find_target_irr_stream(target_irr)
            
            print(f"\nTarget IRR: {goal_seek_results['target_irr']:.2%}")
            print(f"Required Streaming Percentage: {goal_seek_results['streaming_percentage']:.2%}")
            print(f"Actual IRR Achieved: {goal_seek_results['actual_irr']:.2%}")
            print(f"Difference: {goal_seek_results['difference']:.4%}")
            print(f"NPV at Target IRR: ${goal_seek_results['npv']:,.2f}")
            if 'payback_period' in goal_seek_results:
                print(f"Payback Period: {goal_seek_results['payback_period']:.2f} years")
            
            # Run sensitivity analysis
            print("\n" + "="*60)
            print("Running Sensitivity Analysis...")
            print("="*60)
            credit_range = [0.8, 0.9, 1.0, 1.1, 1.2]
            price_range = [0.7, 0.85, 1.0, 1.15, 1.3]
            
            try:
                sensitivity_table = model.run_sensitivity_table(
                    credit_range=credit_range,
                    price_range=price_range
                )
                print("\nSensitivity Analysis Table (IRR by Credit Volume and Price):")
                print(sensitivity_table.to_string())
            except Exception as e:
                print(f"\nSensitivity analysis failed: {e}")
            
            # Export to Excel
            print("\n" + "="*60)
            print("Exporting to Excel...")
            print("="*60)
            try:
                excel_filename = "carbon_model_output.xlsx"
                model.export_model_to_excel(excel_filename)
                print(f"\nModel exported successfully to: {excel_filename}")
                print("The Excel file contains:")
                print("  - Inputs & Summary sheet")
                print("  - Valuation Schedule sheet")
                print("  - Sensitivity Analysis sheet")
            except Exception as e:
                print(f"\nExcel export failed: {e}")
            
        except ValueError as e:
            print(f"\nGoal-seeking failed: {e}")
        
    except FileNotFoundError:
        print(f"\nError: Could not find data file '{data_file}'")
        print("Please ensure the file exists or update the path in this script.")
        print("\nTo use this class with your data:")
        print("1. Export your Excel file to CSV, or")
        print("2. Ensure your Excel file has a sheet named 'Inputs', or")
        print("3. Update the file path in this script")
        
        # Create a sample data structure for demonstration
        print("\n" + "="*60)
        print("Creating Sample Data Structure for Reference:")
        print("="*60)
        create_sample_data_structure()
    
    except Exception as e:
        print(f"\nError: {e}")
        import traceback
        traceback.print_exc()


def create_sample_data_structure():
    """
    Create a sample CSV file structure for reference.
    """
    # Create sample data
    years = range(1, 21)
    sample_data = {
        'Year': years,
        'Carbon Credits Issued (Gross)': [100000 + i*5000 for i in range(20)],
        'Project Implementation Costs': [500000 if i < 3 else 100000 for i in range(20)],
        'Base Carbon Price': [15.0 + i*0.5 for i in range(20)]
    }
    
    df = pd.DataFrame(sample_data)
    sample_file = "sample_input_data.csv"
    df.to_csv(sample_file, index=False)
    print(f"\nSample data structure saved to: {sample_file}")
    print("\nExpected CSV structure:")
    print(df.head(10).to_string())


if __name__ == "__main__":
    main()

