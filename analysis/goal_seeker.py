"""
Goal Seeker Module: Handles optimization for target IRR scenarios.

This module provides functionality to find the streaming percentage required
to achieve a target IRR using optimization techniques.
"""

import numpy as np
import pandas as pd
from scipy.optimize import brentq
from typing import Dict, Callable, Optional
try:
    from ..core.dcf import DCFCalculator
except ImportError:
    from core.dcf import DCFCalculator


class GoalSeeker:
    """
    Finds optimal streaming percentage to achieve target IRR.
    
    Uses optimization methods to iteratively find the streaming percentage
    that results in the desired target IRR.
    """
    
    def __init__(
        self,
        dcf_calculator: DCFCalculator,
        data: pd.DataFrame,
        tolerance: float = 1e-4
    ):
        """
        Initialize the Goal Seeker.
        
        Parameters:
        -----------
        dcf_calculator : DCFCalculator
            DCF calculator instance to use for calculations
        data : pd.DataFrame
            Input data for calculations
        tolerance : float
            Tolerance for convergence (default: 1e-4)
        """
        self.dcf_calculator = dcf_calculator
        self.data = data
        self.tolerance = tolerance
    
    def create_irr_error_function(
        self,
        target_irr: float
    ) -> Callable[[float], float]:
        """
        Create an error function for IRR optimization.
        
        The error function returns the difference between actual IRR and target IRR.
        We want to find where this equals zero.
        
        Parameters:
        -----------
        target_irr : float
            Target IRR as decimal (e.g., 0.20 for 20%)
            
        Returns:
        --------
        Callable
            Error function that takes streaming_percentage and returns error
        """
        def irr_error(streaming_pct: float) -> float:
            """
            Calculate error between actual IRR and target IRR.
            
            Parameters:
            -----------
            streaming_pct : float
                Streaming percentage to test
                
            Returns:
            --------
            float
                Error (actual_irr - target_irr)
            """
            # Validate streaming percentage
            if not (0 <= streaming_pct <= 1):
                return 1e10  # Large error for invalid values
            
            # Run DCF with this streaming percentage
            result = self.dcf_calculator.run_dcf(self.data, streaming_pct)
            actual_irr = result['irr']
            
            # Handle NaN IRR
            if np.isnan(actual_irr):
                return 1e10
            
            return actual_irr - target_irr
        
        return irr_error
    
    def validate_feasibility(
        self,
        error_function: Callable[[float], float]
    ) -> None:
        """
        Validate that a solution is feasible within [0, 1] bounds.
        
        Parameters:
        -----------
        error_function : Callable
            Error function to test
            
        Raises:
        -------
        ValueError
            If target IRR is not achievable
        """
        lower_bound = 0.0
        upper_bound = 1.0
        
        error_lower = error_function(lower_bound)
        error_upper = error_function(upper_bound)
        
        # If both bounds have same sign, solution may not exist in [0, 1]
        if error_lower * error_upper > 0:
            if error_lower > 0:  # IRR too high even at 0%
                raise ValueError(
                    f"Target IRR cannot be achieved. "
                    f"Even at 0% streaming, IRR is too high."
                )
            else:  # IRR too low even at 100%
                raise ValueError(
                    f"Target IRR cannot be achieved. "
                    f"Even at 100% streaming, IRR is too low."
                )
    
    def find_optimal_streaming(
        self,
        error_function: Callable[[float], float]
    ) -> float:
        """
        Find optimal streaming percentage using Brent's method.
        
        Parameters:
        -----------
        error_function : Callable
            Error function to minimize
            
        Returns:
        --------
        float
            Optimal streaming percentage
            
        Raises:
        -------
        RuntimeError
            If optimization fails
        """
        lower_bound = 0.0
        upper_bound = 1.0
        
        try:
            optimal_streaming = brentq(
                error_function,
                lower_bound,
                upper_bound,
                xtol=self.tolerance,
                maxiter=100
            )
            return optimal_streaming
        except (ValueError, RuntimeError) as e:
            raise RuntimeError(
                f"Could not find optimal streaming percentage. "
                f"Optimization failed: {e}"
            )
    
    def find_target_irr_stream(
        self,
        target_irr: float
    ) -> Dict:
        """
        Find streaming percentage that achieves target IRR.
        
        Parameters:
        -----------
        target_irr : float
            Target IRR as decimal (e.g., 0.20 for 20%)
            
        Returns:
        --------
        dict
            Dictionary containing:
            - 'streaming_percentage': The calculated streaming percentage
            - 'actual_irr': The actual IRR achieved
            - 'target_irr': The target IRR
            - 'difference': The difference between actual and target
            - 'results_df': Full DCF results at the calculated streaming percentage
            - 'npv': NPV at the calculated streaming percentage
        """
        # Create error function
        error_function = self.create_irr_error_function(target_irr)
        
        # Validate feasibility
        self.validate_feasibility(error_function)
        
        # Find optimal streaming percentage
        optimal_streaming = self.find_optimal_streaming(error_function)
        
        # Run final DCF with optimal streaming percentage
        final_results = self.dcf_calculator.run_dcf(self.data, optimal_streaming)
        
        return {
            'streaming_percentage': optimal_streaming,
            'actual_irr': final_results['irr'],
            'target_irr': target_irr,
            'difference': abs(final_results['irr'] - target_irr),
            'results_df': final_results['results_df'],
            'npv': final_results['npv']
        }

