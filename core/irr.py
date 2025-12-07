"""
IRR Calculator Module: Handles Internal Rate of Return calculations.

This module provides robust IRR calculation using various optimization methods
with fallback strategies for edge cases.
"""

import numpy as np
from scipy.optimize import brentq, fsolve
import warnings
from typing import Optional


class IRRCalculator:
    """
    Calculates Internal Rate of Return (IRR) for cash flow streams.
    
    Uses Brent's method as the primary approach with fallback strategies
    for edge cases.
    """
    
    def __init__(self, default_guess: float = 0.1, tolerance: float = 1e-6):
        """
        Initialize the IRR Calculator.
        
        Parameters:
        -----------
        default_guess : float
            Default initial guess for IRR (default: 0.1 = 10%)
        tolerance : float
            Tolerance for convergence (default: 1e-6)
        """
        self.default_guess = default_guess
        self.tolerance = tolerance
    
    def npv_function(self, cash_flows: np.ndarray, rate: float) -> float:
        """
        Calculate NPV for a given discount rate.
        
        Parameters:
        -----------
        cash_flows : np.ndarray
            Array of cash flows
        rate : float
            Discount rate
            
        Returns:
        --------
        float
            Net Present Value
        """
        periods = np.arange(len(cash_flows))
        return np.sum(cash_flows / ((1 + rate) ** periods))
    
    def find_bounds(self, cash_flows: np.ndarray) -> tuple:
        """
        Find reasonable bounds for IRR calculation.
        
        Parameters:
        -----------
        cash_flows : np.ndarray
            Array of cash flows
            
        Returns:
        --------
        tuple
            (lower_bound, upper_bound) for IRR search
        """
        lower_bound = -0.99  # Can't have -100% or less
        upper_bound = 10.0   # 1000% as initial upper limit
        
        # Check if NPV is still positive at upper bound
        if self.npv_function(cash_flows, upper_bound) > 0:
            upper_bound = 100.0  # Try even higher
        
        return lower_bound, upper_bound
    
    def calculate_irr_brentq(self, cash_flows: np.ndarray) -> Optional[float]:
        """
        Calculate IRR using Brent's method (bracketing).
        
        Parameters:
        -----------
        cash_flows : np.ndarray
            Array of cash flows (negative for investment, positive for returns)
            
        Returns:
        --------
        float or None
            Internal Rate of Return (as decimal) or None if calculation fails
        """
        def npv_func(rate):
            return self.npv_function(cash_flows, rate)
        
        try:
            lower_bound, upper_bound = self.find_bounds(cash_flows)
            irr = brentq(
                npv_func,
                lower_bound,
                upper_bound,
                xtol=self.tolerance,
                maxiter=100
            )
            return irr
        except (ValueError, RuntimeError):
            return None
    
    def calculate_irr_fsolve(self, cash_flows: np.ndarray) -> Optional[float]:
        """
        Calculate IRR using fsolve (alternative method).
        
        Parameters:
        -----------
        cash_flows : np.ndarray
            Array of cash flows
            
        Returns:
        --------
        float or None
            Internal Rate of Return (as decimal) or None if calculation fails
        """
        def npv_func(rate):
            return self.npv_function(cash_flows, rate[0])
        
        try:
            irr = fsolve(npv_func, [self.default_guess])[0]
            # Validate the result
            if abs(self.npv_function(cash_flows, irr)) > 1e-3:
                return None
            return irr
        except:
            return None
    
    def calculate_irr(self, cash_flows: np.ndarray) -> float:
        """
        Calculate Internal Rate of Return with fallback strategies.
        
        Tries multiple methods in order:
        1. Brent's method (brentq) - most reliable
        2. fsolve - alternative optimization
        3. Returns NaN if all methods fail
        
        Parameters:
        -----------
        cash_flows : np.ndarray
            Array of cash flows (negative for investment, positive for returns)
            
        Returns:
        --------
        float
            Internal Rate of Return (as decimal, e.g., 0.20 for 20%)
            Returns np.nan if calculation fails
        """
        # Validate input
        if len(cash_flows) == 0:
            return np.nan
        
        # Try Brent's method first (most reliable)
        irr = self.calculate_irr_brentq(cash_flows)
        if irr is not None:
            return irr
        
        # Fallback to fsolve
        warnings.warn(
            "IRR calculation using brentq failed. Trying alternative method (fsolve)."
        )
        irr = self.calculate_irr_fsolve(cash_flows)
        if irr is not None:
            return irr
        
        # Last resort: return NaN
        warnings.warn("Could not calculate IRR. All methods failed. Returning NaN.")
        return np.nan

