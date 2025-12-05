"""
DCF Calculator Module: Handles Discounted Cash Flow calculations.

This module provides functionality to calculate cash flows, NPV, and other
financial metrics for carbon credit streaming agreements.
"""

import pandas as pd
import numpy as np
from typing import Dict, Optional
from .irr_calculator import IRRCalculator


class DCFCalculator:
    """
    Performs DCF (Discounted Cash Flow) calculations for carbon credit projects.
    
    Calculates revenue, cash flows, NPV, and other financial metrics.
    """
    
    def __init__(
        self,
        wacc: float,
        rubicon_investment_total: float,
        investment_tenor: int,
        irr_calculator: Optional[IRRCalculator] = None
    ):
        """
        Initialize the DCF Calculator.
        
        Parameters:
        -----------
        wacc : float
            Discount rate (e.g., 0.08 for 8.0%)
        rubicon_investment_total : float
            Initial investment amount in USD
        investment_tenor : int
            Period over which investment is deployed in years
        irr_calculator : IRRCalculator, optional
            IRR calculator instance. If None, creates a new one.
        """
        self.wacc = wacc
        self.rubicon_investment_total = rubicon_investment_total
        self.investment_tenor = investment_tenor
        self.irr_calculator = irr_calculator or IRRCalculator()
    
    def calculate_share_of_credits(
        self,
        data: pd.DataFrame,
        streaming_percentage: float
    ) -> pd.Series:
        """
        Calculate Rubicon's share of carbon credits.
        
        Parameters:
        -----------
        data : pd.DataFrame
            Input data with 'carbon_credits_gross' column
        streaming_percentage : float
            Percentage of credits Rubicon receives (0.0 to 1.0)
            
        Returns:
        --------
        pd.Series
            Rubicon's share of credits
        """
        return data['carbon_credits_gross'] * streaming_percentage
    
    def calculate_revenue(
        self,
        data: pd.DataFrame,
        share_of_credits: pd.Series
    ) -> pd.Series:
        """
        Calculate Rubicon's revenue from credit sales.
        
        Parameters:
        -----------
        data : pd.DataFrame
            Input data with 'base_carbon_price' column
        share_of_credits : pd.Series
            Rubicon's share of credits
            
        Returns:
        --------
        pd.Series
            Revenue (Credits × Price)
        """
        return share_of_credits * data['base_carbon_price']
    
    def calculate_investment_cash_flow(
        self,
        data: pd.DataFrame
    ) -> pd.Series:
        """
        Calculate investment deployment cash flow.
        
        Distributes investment evenly over investment_tenor years as
        negative cash flows.
        
        Parameters:
        -----------
        data : pd.DataFrame
            Input data (used for index alignment)
            
        Returns:
        --------
        pd.Series
            Investment cash flow (negative for first N years)
        """
        annual_investment = self.rubicon_investment_total / self.investment_tenor
        investment_cf = pd.Series(0.0, index=data.index)
        investment_cf.loc[investment_cf.index <= self.investment_tenor] = -annual_investment
        return investment_cf
    
    def calculate_net_cash_flow(
        self,
        revenue: pd.Series,
        investment_cf: pd.Series
    ) -> pd.Series:
        """
        Calculate net cash flow (Revenue - Investment).
        
        Parameters:
        -----------
        revenue : pd.Series
            Revenue stream
        investment_cf : pd.Series
            Investment cash flow
            
        Returns:
        --------
        pd.Series
            Net cash flow
        """
        return revenue + investment_cf
    
    def calculate_discount_factors(self, data: pd.DataFrame) -> pd.Series:
        """
        Calculate discount factors for each period.
        
        Parameters:
        -----------
        data : pd.DataFrame
            Input data (used for index alignment)
            
        Returns:
        --------
        pd.Series
            Discount factors
        """
        # Discount factor = 1 / (1 + WACC)^(Year - 1)
        # Year 1 is not discounted (Year - 1 = 0)
        return 1 / ((1 + self.wacc) ** (data.index - 1))
    
    def calculate_present_values(
        self,
        cash_flows: pd.Series,
        discount_factors: pd.Series
    ) -> pd.Series:
        """
        Calculate present value of each cash flow.
        
        Parameters:
        -----------
        cash_flows : pd.Series
            Cash flow stream
        discount_factors : pd.Series
            Discount factors
            
        Returns:
        --------
        pd.Series
            Present values
        """
        return cash_flows * discount_factors
    
    def calculate_npv(self, present_values: pd.Series) -> float:
        """
        Calculate Net Present Value.
        
        Parameters:
        -----------
        present_values : pd.Series
            Present values of cash flows
            
        Returns:
        --------
        float
            Net Present Value
        """
        return present_values.sum()
    
    def calculate_cumulative_metrics(
        self,
        cash_flows: pd.Series,
        present_values: pd.Series
    ) -> Dict[str, pd.Series]:
        """
        Calculate cumulative cash flow and cumulative present value.
        
        Parameters:
        -----------
        cash_flows : pd.Series
            Net cash flow stream
        present_values : pd.Series
            Present values
            
        Returns:
        --------
        Dict[str, pd.Series]
            Dictionary with 'cumulative_cash_flow' and 'cumulative_pv'
        """
        return {
            'cumulative_cash_flow': cash_flows.cumsum(),
            'cumulative_pv': present_values.cumsum()
        }
    
    def run_dcf(
        self,
        data: pd.DataFrame,
        streaming_percentage: float
    ) -> Dict:
        """
        Run complete DCF analysis.
        
        Parameters:
        -----------
        data : pd.DataFrame
            Input data with required columns
        streaming_percentage : float
            Percentage of credits Rubicon receives (0.0 to 1.0)
            
        Returns:
        --------
        dict
            Dictionary containing:
            - 'results_df': DataFrame with all calculated metrics
            - 'npv': Net Present Value
            - 'irr': Internal Rate of Return
            - 'cash_flows': Net cash flow series
        """
        # Validate streaming percentage
        if not (0 <= streaming_percentage <= 1):
            raise ValueError(
                f"streaming_percentage must be between 0 and 1, "
                f"got {streaming_percentage}"
            )
        
        # Initialize results DataFrame
        results = data.copy()
        
        # Calculate Rubicon's Share of Credits
        results['rubicon_share_credits'] = self.calculate_share_of_credits(
            data, streaming_percentage
        )
        
        # Calculate Rubicon's Revenue
        results['rubicon_revenue'] = self.calculate_revenue(
            data, results['rubicon_share_credits']
        )
        
        # Calculate Investment Cash Flow
        results['rubicon_investment_cf'] = self.calculate_investment_cash_flow(data)
        
        # Calculate Net Cash Flow
        results['rubicon_net_cash_flow'] = self.calculate_net_cash_flow(
            results['rubicon_revenue'],
            results['rubicon_investment_cf']
        )
        
        # Calculate Discount Factors
        results['discount_factor'] = self.calculate_discount_factors(data)
        
        # Calculate Present Values
        results['present_value'] = self.calculate_present_values(
            results['rubicon_net_cash_flow'],
            results['discount_factor']
        )
        
        # Calculate Cumulative Metrics
        cumulative = self.calculate_cumulative_metrics(
            results['rubicon_net_cash_flow'],
            results['present_value']
        )
        results['cumulative_cash_flow'] = cumulative['cumulative_cash_flow']
        results['cumulative_pv'] = cumulative['cumulative_pv']
        
        # Calculate NPV
        npv = self.calculate_npv(results['present_value'])
        
        # Calculate IRR
        cash_flows_array = results['rubicon_net_cash_flow'].values
        irr = self.irr_calculator.calculate_irr(cash_flows_array)
        
        return {
            'results_df': results,
            'npv': npv,
            'irr': irr,
            'cash_flows': results['rubicon_net_cash_flow']
        }

