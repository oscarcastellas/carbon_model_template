"""
Example demonstrating flexible assumption handling in CarbonModelGenerator.

Shows three ways to provide assumptions:
1. Direct input in constructor
2. Extract from data file
3. Set after initialization
"""

from carbon_model_generator import CarbonModelGenerator
import pandas as pd


def example_1_direct_input():
    """Example 1: Provide assumptions directly in constructor."""
    print("="*60)
    print("Example 1: Direct Input Assumptions")
    print("="*60)
    
    model = CarbonModelGenerator(
        wacc=0.08,
        rubicon_investment_total=20_000_000,
        investment_tenor=5,
        streaming_percentage_initial=0.48
    )
    
    print("Model created with assumptions:")
    print(model.get_assumptions())
    print("\n✓ All assumptions set, ready to load data")


def example_2_extract_from_file():
    """Example 2: Extract assumptions from data file."""
    print("\n" + "="*60)
    print("Example 2: Extract Assumptions from File")
    print("="*60)
    
    # Create model without assumptions
    model = CarbonModelGenerator()
    
    # Load data and extract assumptions
    data_file = "Analyst_Model_Test_OCC.xlsx"
    try:
        data, extracted_assumptions = model.load_data_with_assumptions(
            data_file,
            use_extracted_assumptions=True
        )
        
        print(f"Extracted assumptions from file:")
        for key, value in extracted_assumptions.items():
            print(f"  {key}: {value}")
        
        if model._has_all_assumptions():
            print("\n✓ All assumptions extracted, ready to run analysis")
        else:
            print("\n⚠ Some assumptions missing. Please set them manually:")
            missing = [k for k, v in model.get_assumptions().items() if v is None]
            print(f"  Missing: {missing}")
            print("\nYou can set them using:")
            print("  model.set_assumptions(wacc=0.08, ...)")
            
    except FileNotFoundError:
        print(f"File {data_file} not found. Skipping this example.")


def example_3_set_after_init():
    """Example 3: Set assumptions after initialization."""
    print("\n" + "="*60)
    print("Example 3: Set Assumptions After Initialization")
    print("="*60)
    
    # Create model without assumptions
    model = CarbonModelGenerator()
    
    # Set assumptions using dictionary
    model.set_assumptions(assumptions={
        'wacc': 0.08,
        'rubicon_investment_total': 20_000_000,
        'investment_tenor': 5,
        'streaming_percentage_initial': 0.48
    })
    
    print("Assumptions set after initialization:")
    print(model.get_assumptions())
    print("\n✓ All assumptions set, ready to load data")
    
    # Or set individually
    print("\n" + "-"*60)
    print("Updating individual assumptions:")
    model.set_assumptions(wacc=0.10)  # Update just WACC
    print(f"Updated WACC: {model.wacc}")


def example_4_mixed_approach():
    """Example 4: Mixed approach - some from file, some from user."""
    print("\n" + "="*60)
    print("Example 4: Mixed Approach (File + Override)")
    print("="*60)
    
    model = CarbonModelGenerator()
    
    data_file = "Analyst_Model_Test_OCC.xlsx"
    try:
        # Extract from file, but override some values
        data, extracted = model.load_data_with_assumptions(
            data_file,
            use_extracted_assumptions=True,
            override_assumptions={
                'wacc': 0.10,  # Override extracted WACC
                'streaming_percentage_initial': 0.50  # Override extracted streaming
            }
        )
        
        print("Final assumptions (extracted + overrides):")
        print(model.get_assumptions())
        
    except FileNotFoundError:
        print(f"File {data_file} not found. Skipping this example.")


def example_5_assumption_validation():
    """Example 5: Show what happens when assumptions are missing."""
    print("\n" + "="*60)
    print("Example 5: Assumption Validation")
    print("="*60)
    
    model = CarbonModelGenerator()
    
    print("Model created without assumptions:")
    print(f"  Has all assumptions: {model._has_all_assumptions()}")
    print(f"  Current assumptions: {model.get_assumptions()}")
    
    # Try to load data without assumptions
    try:
        model.load_data("dummy_file.xlsx")
    except ValueError as e:
        print(f"\n✓ Correctly prevented loading data: {e}")
    
    # Set partial assumptions
    model.set_assumptions(wacc=0.08, rubicon_investment_total=20_000_000)
    print(f"\nAfter setting partial assumptions:")
    print(f"  Has all assumptions: {model._has_all_assumptions()}")
    print(f"  Missing: {[k for k, v in model.get_assumptions().items() if v is None]}")


if __name__ == "__main__":
    example_1_direct_input()
    example_2_extract_from_file()
    example_3_set_after_init()
    example_4_mixed_approach()
    example_5_assumption_validation()
    
    print("\n" + "="*60)
    print("Summary: Three Ways to Provide Assumptions")
    print("="*60)
    print("""
1. Direct Input (Constructor):
   model = CarbonModelGenerator(wacc=0.08, ...)

2. Extract from File:
   model = CarbonModelGenerator()
   data, assumptions = model.load_data_with_assumptions("file.xlsx")

3. Set After Initialization:
   model = CarbonModelGenerator()
   model.set_assumptions(wacc=0.08, ...)
   # or
   model.set_assumptions(assumptions={'wacc': 0.08, ...})
    """)

