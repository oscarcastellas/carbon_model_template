"""
Sensitivity Analysis Module: Performs sensitivity analysis on DCF model.

This module provides functionality to run sensitivity analysis by varying
key inputs and calculating resulting IRR for each scenario.
"""

import pandas as pd
import numpy as np
from typing import List, Dict, Optional
from .dcf_calculator import DCFCalculator


class SensitivityAnalyzer:
    """
    Performs sensitivity analysis on carbon credit streaming models.
    
    Varies credit volumes and carbon prices to see impact on IRR.
    """
    
    def __init__(self, dcf_calculator: DCFCalculator):
        """
        Initialize the Sensitivity Analyzer.
        
        Parameters:
        -----------
        dcf_calculator : DCFCalculator
            DCF calculator instance to use for calculations
        """
        self.dcf_calculator = dcf_calculator
    
    def run_sensitivity_table(
        self,
        data: pd.DataFrame,
        streaming_percentage: float,
        credit_range: List[float],
        price_range: List[float]
    ) -> pd.DataFrame:
        """
        Run sensitivity analysis by varying credit volumes and prices.
        
        Creates a 2D table showing IRR for each combination of:
        - Credit Volume Multipliers (rows)
        - Carbon Price Multipliers (columns)
        
        Parameters:
        -----------
        data : pd.DataFrame
            Base input data
        streaming_percentage : float
            Target streaming percentage to use for all scenarios
        credit_range : List[float]
            Range of credit volume multipliers (e.g., [0.9, 1.0, 1.1])
        price_range : List[float]
            Range of carbon price multipliers (e.g., [0.8, 1.0, 1.2])
            
        Returns:
        --------
        pd.DataFrame
            2D DataFrame with credit multipliers as index, price multipliers as columns,
            and IRR values as cells
        """
        # Initialize results matrix
        results = []
        
        # Iterate over credit multipliers (rows)
        for credit_mult in credit_range:
            row_results = []
            
            # Iterate over price multipliers (columns)
            for price_mult in price_range:
                # Create modified data
                modified_data = data.copy()
                modified_data['carbon_credits_gross'] = (
                    data['carbon_credits_gross'] * credit_mult
                )
                modified_data['base_carbon_price'] = (
                    data['base_carbon_price'] * price_mult
                )
                
                # Run DCF with modified data
                try:
                    dcf_results = self.dcf_calculator.run_dcf(
                        modified_data,
                        streaming_percentage
                    )
                    irr = dcf_results['irr']
                    
                    # Handle NaN or invalid IRR
                    if pd.isna(irr) or not np.isfinite(irr):
                        row_results.append(np.nan)
                    else:
                        row_results.append(irr)
                except Exception as e:
                    # If calculation fails, store NaN
                    row_results.append(np.nan)
            
            results.append(row_results)
        
        # Create DataFrame
        sensitivity_df = pd.DataFrame(
            results,
            index=[f"{mult:.2f}x" for mult in credit_range],
            columns=[f"{mult:.2f}x" for mult in price_range]
        )
        
        # Add descriptive index and column names
        sensitivity_df.index.name = 'Credit Volume Multiplier'
        sensitivity_df.columns.name = 'Carbon Price Multiplier'
        
        return sensitivity_df

