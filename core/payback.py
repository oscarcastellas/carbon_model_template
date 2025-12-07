"""
Payback Period Calculator Module: Calculates payback period for investments.

This module provides functionality to calculate the payback period based on
cumulative cash flows.
"""

import pandas as pd
import numpy as np
from typing import Optional


class PaybackCalculator:
    """
    Calculates payback period for investment cash flows.
    
    Payback period is the time it takes for cumulative cash flows to become positive.
    """
    
    def calculate_payback_period(
        self,
        cash_flows: pd.Series,
        method: str = 'simple'
    ) -> Optional[float]:
        """
        Calculate payback period from cash flow series.
        
        Parameters:
        -----------
        cash_flows : pd.Series
            Net cash flow series (indexed by Year)
        method : str
            'simple' for simple payback, 'discounted' for discounted payback
            
        Returns:
        --------
        float or None
            Payback period in years, or None if payback never occurs
        """
        if method == 'simple':
            return self._calculate_simple_payback(cash_flows)
        elif method == 'discounted':
            return self._calculate_discounted_payback(cash_flows)
        else:
            raise ValueError(f"Unknown method: {method}. Use 'simple' or 'discounted'")
    
    def _calculate_simple_payback(self, cash_flows: pd.Series) -> Optional[float]:
        """
        Calculate simple payback period.
        
        Parameters:
        -----------
        cash_flows : pd.Series
            Net cash flow series
            
        Returns:
        --------
        float or None
            Payback period in years
        """
        cumulative = cash_flows.cumsum()
        
        # Find first year where cumulative is positive
        positive_years = cumulative[cumulative > 0]
        
        if len(positive_years) == 0:
            return None  # Never pays back
        
        first_positive_year = positive_years.index[0]
        
        # If cumulative was negative in previous year, interpolate
        if first_positive_year > 1:
            prev_year = first_positive_year - 1
            prev_cumulative = cumulative.loc[prev_year]
            curr_cumulative = cumulative.loc[first_positive_year]
            year_cf = cash_flows.loc[first_positive_year]
            
            if year_cf != 0:
                # Interpolate to find exact payback point
                fraction = abs(prev_cumulative) / year_cf
                return prev_year + fraction
            else:
                return float(first_positive_year)
        else:
            return float(first_positive_year)
    
    def _calculate_discounted_payback(
        self,
        cash_flows: pd.Series,
        discount_rate: float = 0.08
    ) -> Optional[float]:
        """
        Calculate discounted payback period.
        
        Parameters:
        -----------
        cash_flows : pd.Series
            Net cash flow series
        discount_rate : float
            Discount rate for present value calculation
            
        Returns:
        --------
        float or None
            Discounted payback period in years
        """
        # Calculate present values
        discount_factors = 1 / ((1 + discount_rate) ** (cash_flows.index - 1))
        present_values = cash_flows * discount_factors
        cumulative_pv = present_values.cumsum()
        
        # Find first year where cumulative PV is positive
        positive_years = cumulative_pv[cumulative_pv > 0]
        
        if len(positive_years) == 0:
            return None  # Never pays back
        
        first_positive_year = positive_years.index[0]
        
        # If cumulative was negative in previous year, interpolate
        if first_positive_year > 1:
            prev_year = first_positive_year - 1
            prev_cumulative = cumulative_pv.loc[prev_year]
            curr_cumulative = cumulative_pv.loc[first_positive_year]
            year_pv = present_values.loc[first_positive_year]
            
            if year_pv != 0:
                # Interpolate to find exact payback point
                fraction = abs(prev_cumulative) / year_pv
                return prev_year + fraction
            else:
                return float(first_positive_year)
        else:
            return float(first_positive_year)

