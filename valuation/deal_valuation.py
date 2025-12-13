"""
Deal Valuation Solver Module: Solves for purchase price, IRR, or streaming percentage.

This module extends goal-seeking functionality to solve for purchase price given
target IRR, or solve for IRR given purchase price. This solves the core business
question: "What price should we pay for this deal?"
"""

import pandas as pd
import numpy as np
from scipy.optimize import brentq
from typing import Dict, Optional, Callable
try:
    from ..core.dcf import DCFCalculator
    from ..core.irr import IRRCalculator
except ImportError:
    from core.dcf import DCFCalculator
    from core.irr import IRRCalculator


class DealValuationSolver:
    """
    Solves for purchase price, IRR, or streaming percentage in streaming deals.
    
    Extends GoalSeeker functionality to handle price-based optimization.
    """
    
    def __init__(
        self,
        dcf_calculator: DCFCalculator,
        data: pd.DataFrame,
        tolerance: float = 1e-4
    ):
        """
        Initialize the Deal Valuation Solver.
        
        Parameters:
        -----------
        dcf_calculator : DCFCalculator
            DCF calculator instance (will be modified with new investment_total)
        data : pd.DataFrame
            Input data for calculations
        tolerance : float
            Tolerance for convergence (default: 1e-4)
        """
        self.dcf_calculator = dcf_calculator
        self.data = data
        self.tolerance = tolerance
        self.original_investment_total = dcf_calculator.rubicon_investment_total
        self.original_investment_tenor = dcf_calculator.investment_tenor
        self.original_wacc = dcf_calculator.wacc
        self.original_irr_calculator = dcf_calculator.irr_calculator
    
    def create_price_error_function(
        self,
        target_irr: float,
        streaming_percentage: float,
        investment_tenor: int
    ) -> Callable[[float], float]:
        """
        Create error function for price optimization.
        
        The error function returns the difference between actual IRR and target IRR.
        We want to find where this equals zero.
        
        Parameters:
        -----------
        target_irr : float
            Target IRR as decimal
        streaming_percentage : float
            Fixed streaming percentage
        investment_tenor : int
            Investment tenor
            
        Returns:
        --------
        Callable
            Error function that takes purchase_price and returns error
        """
        def price_error(purchase_price: float) -> float:
            """
            Calculate error between actual IRR and target IRR.
            
            Parameters:
            -----------
            purchase_price : float
                Purchase price to test
                
            Returns:
            --------
            float
                Error (actual_irr - target_irr)
            """
            # Validate purchase price
            if purchase_price <= 0:
                return 1e10  # Large error for invalid values
            
            # Create temporary DCF calculator with new investment total
            temp_dcf = DCFCalculator(
                wacc=self.original_wacc,
                rubicon_investment_total=purchase_price,
                investment_tenor=investment_tenor,
                irr_calculator=self.original_irr_calculator
            )
            
            # Run DCF with this purchase price
            result = temp_dcf.run_dcf(self.data, streaming_percentage)
            actual_irr = result['irr']
            
            # Handle NaN IRR
            if np.isnan(actual_irr):
                return 1e10
            
            return actual_irr - target_irr
        
        return price_error
    
    def validate_price_feasibility(
        self,
        error_function: Callable[[float], float],
        min_price: float = 1_000,
        max_price: float = 1_000_000_000
    ) -> None:
        """
        Validate that a solution is feasible within price bounds.
        
        Parameters:
        -----------
        error_function : Callable
            Error function to test
        min_price : float
            Minimum price to test (default: $1,000)
        max_price : float
            Maximum price to test (default: $1B)
            
        Raises:
        -------
        ValueError
            If target IRR is not achievable
        """
        error_min = error_function(min_price)
        error_max = error_function(max_price)
        
        # If both bounds have same sign, solution may not exist
        if error_min * error_max > 0:
            if error_min > 0:  # IRR too high even at min price
                raise ValueError(
                    f"Target IRR cannot be achieved. "
                    f"Even at ${min_price:,.0f}, IRR is too high."
                )
            else:  # IRR too low even at max price
                raise ValueError(
                    f"Target IRR cannot be achieved. "
                    f"Even at ${max_price:,.0f}, IRR is too low."
                )
    
    def find_optimal_price(
        self,
        error_function: Callable[[float], float],
        min_price: float = 1_000,
        max_price: float = 1_000_000_000
    ) -> float:
        """
        Find optimal purchase price using Brent's method.
        
        Parameters:
        -----------
        error_function : Callable
            Error function to minimize
        min_price : float
            Minimum price bound
        max_price : float
            Maximum price bound
            
        Returns:
        --------
        float
            Optimal purchase price
            
        Raises:
        -------
        RuntimeError
            If optimization fails
        """
        try:
            optimal_price = brentq(
                error_function,
                min_price,
                max_price,
                xtol=self.tolerance,
                maxiter=100
            )
            return optimal_price
        except (ValueError, RuntimeError) as e:
            raise RuntimeError(
                f"Could not find optimal purchase price. "
                f"Optimization failed: {e}"
            )
    
    def solve_for_purchase_price(
        self,
        target_irr: float,
        streaming_percentage: float,
        investment_tenor: Optional[int] = None
    ) -> Dict:
        """
        Solve for maximum purchase price given target IRR.
        
        This is the inverse of goal-seeking:
        - Instead of: streaming % → IRR
        - We solve: price → IRR (where price affects investment_total)
        
        Parameters:
        -----------
        target_irr : float
            Target IRR as decimal (e.g., 0.20 for 20%)
        streaming_percentage : float
            Fixed streaming percentage
        investment_tenor : int, optional
            Investment tenor (uses original if not provided)
            
        Returns:
        --------
        dict
            Dictionary containing:
            - 'purchase_price': Maximum purchase price
            - 'actual_irr': Actual IRR achieved
            - 'target_irr': Target IRR
            - 'streaming_percentage': Streaming percentage used
            - 'npv': NPV at calculated price
            - 'results_df': Full DCF results
        """
        if investment_tenor is None:
            investment_tenor = self.original_investment_tenor
        
        # Create error function
        error_function = self.create_price_error_function(
            target_irr=target_irr,
            streaming_percentage=streaming_percentage,
            investment_tenor=investment_tenor
        )
        
        # Validate feasibility
        self.validate_price_feasibility(error_function)
        
        # Find optimal purchase price
        optimal_price = self.find_optimal_price(error_function)
        
        # Run final DCF with optimal purchase price
        temp_dcf = DCFCalculator(
            wacc=self.original_wacc,
            rubicon_investment_total=optimal_price,
            investment_tenor=investment_tenor,
            irr_calculator=self.original_irr_calculator
        )
        final_results = temp_dcf.run_dcf(self.data, streaming_percentage)
        
        return {
            'purchase_price': optimal_price,
            'actual_irr': final_results['irr'],
            'target_irr': target_irr,
            'difference': abs(final_results['irr'] - target_irr),
            'streaming_percentage': streaming_percentage,
            'investment_tenor': investment_tenor,
            'npv': final_results['npv'],
            'results_df': final_results['results_df']
        }
    
    def solve_for_project_irr(
        self,
        purchase_price: float,
        streaming_percentage: float,
        investment_tenor: Optional[int] = None
    ) -> Dict:
        """
        Calculate project IRR given a specific purchase price.
        
        Useful for: "If we pay $X, what IRR do we get?"
        
        Parameters:
        -----------
        purchase_price : float
            Purchase price in USD
        streaming_percentage : float
            Streaming percentage
        investment_tenor : int, optional
            Investment tenor (uses original if not provided)
            
        Returns:
        --------
        dict
            Dictionary containing:
            - 'purchase_price': Purchase price
            - 'irr': Project IRR
            - 'npv': NPV
            - 'results_df': Full DCF results
        """
        if investment_tenor is None:
            investment_tenor = self.original_investment_tenor
        
        # Validate purchase price
        if purchase_price <= 0:
            raise ValueError("Purchase price must be positive")
        
        # Create temporary DCF calculator with specified purchase price
        temp_dcf = DCFCalculator(
            wacc=self.original_wacc,
            rubicon_investment_total=purchase_price,
            investment_tenor=investment_tenor,
            irr_calculator=self.original_irr_calculator
        )
        
        # Run DCF
        results = temp_dcf.run_dcf(self.data, streaming_percentage)
        
        return {
            'purchase_price': purchase_price,
            'irr': results['irr'],
            'npv': results['npv'],
            'streaming_percentage': streaming_percentage,
            'investment_tenor': investment_tenor,
            'results_df': results['results_df']
        }
    
    def solve_for_streaming_given_price(
        self,
        purchase_price: float,
        target_irr: float,
        investment_tenor: Optional[int] = None
    ) -> Dict:
        """
        Solve for streaming percentage given purchase price and target IRR.
        
        Useful for: "If we pay $X and want Y% IRR, what streaming % do we need?"
        
        This reuses the existing GoalSeeker logic but with a modified investment_total.
        
        Parameters:
        -----------
        purchase_price : float
            Purchase price in USD
        target_irr : float
            Target IRR as decimal
        investment_tenor : int, optional
            Investment tenor (uses original if not provided)
            
        Returns:
        --------
        dict
            Dictionary containing:
            - 'streaming_percentage': Required streaming percentage
            - 'purchase_price': Purchase price
            - 'actual_irr': Actual IRR achieved
            - 'target_irr': Target IRR
            - 'npv': NPV
            - 'results_df': Full DCF results
        """
        if investment_tenor is None:
            investment_tenor = self.original_investment_tenor
        
        # Validate purchase price
        if purchase_price <= 0:
            raise ValueError("Purchase price must be positive")
        
        # Create temporary DCF calculator with specified purchase price
        temp_dcf = DCFCalculator(
            wacc=self.original_wacc,
            rubicon_investment_total=purchase_price,
            investment_tenor=investment_tenor,
            irr_calculator=self.original_irr_calculator
        )
        
        # Use GoalSeeker logic to find streaming percentage
        try:
            from ..analysis.goal_seeker import GoalSeeker
        except ImportError:
            from analysis.goal_seeker import GoalSeeker
        goal_seeker = GoalSeeker(
            dcf_calculator=temp_dcf,
            data=self.data,
            tolerance=self.tolerance
        )
        
        # Find optimal streaming percentage
        goal_results = goal_seeker.find_target_irr_stream(target_irr)
        
        return {
            'streaming_percentage': goal_results['streaming_percentage'],
            'purchase_price': purchase_price,
            'actual_irr': goal_results['actual_irr'],
            'target_irr': target_irr,
            'difference': goal_results['difference'],
            'investment_tenor': investment_tenor,
            'npv': goal_results['npv'],
            'results_df': goal_results['results_df']
        }

