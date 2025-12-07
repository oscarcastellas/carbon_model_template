"""
Breakeven Calculator Module: Finds breakeven points for key variables.

This module calculates breakeven points for price, volume, and streaming
percentage to support deal assessment and negotiation.
"""

import pandas as pd
import numpy as np
from scipy.optimize import brentq, fsolve
from typing import Dict, Optional, Tuple
try:
    from ..core.dcf import DCFCalculator
    from ..core.irr import IRRCalculator
except ImportError:
    from core.dcf import DCFCalculator
    from core.irr import IRRCalculator


class BreakevenCalculator:
    """
    Calculates breakeven points for financial model variables.
    
    Finds the value of a variable (price, volume, streaming %) needed
    to achieve a target metric (NPV = 0, IRR = target, etc.).
    """
    
    def __init__(
        self,
        dcf_calculator: DCFCalculator,
        irr_calculator: IRRCalculator
    ):
        """
        Initialize Breakeven Calculator.
        
        Parameters:
        -----------
        dcf_calculator : DCFCalculator
            DCF calculator instance
        irr_calculator : IRRCalculator
            IRR calculator instance
        """
        self.dcf_calculator = dcf_calculator
        self.irr_calculator = irr_calculator
    
    def calculate_breakeven_price(
        self,
        data: pd.DataFrame,
        streaming_percentage: float,
        target_npv: float = 0.0,
        tolerance: float = 1e-4
    ) -> Dict:
        """
        Calculate breakeven carbon price (price needed for target NPV).
        
        Parameters:
        -----------
        data : pd.DataFrame
            Base project data
        streaming_percentage : float
            Streaming percentage to use
        target_npv : float
            Target NPV (default: 0 for true breakeven)
        tolerance : float
            Optimization tolerance
            
        Returns:
        --------
        Dict
            Dictionary containing:
            - 'breakeven_price': Price per ton needed
            - 'base_price': Average base price from data
            - 'price_multiplier': Multiplier needed (breakeven/base)
            - 'target_npv': Target NPV used
        """
        # Get average base price
        base_prices = data['base_carbon_price']
        avg_base_price = base_prices[base_prices > 0].mean()
        
        if pd.isna(avg_base_price) or avg_base_price <= 0:
            return {
                'breakeven_price': np.nan,
                'base_price': 0.0,
                'price_multiplier': np.nan,
                'target_npv': target_npv,
                'error': 'No valid base prices found'
            }
        
        # Create error function for optimization
        def npv_error(price_multiplier: float) -> float:
            modified_data = data.copy()
            modified_data['base_carbon_price'] = base_prices * price_multiplier
            
            try:
                results = self.dcf_calculator.run_dcf(modified_data, streaming_percentage)
                npv = results['npv']
                if pd.isna(npv):
                    return 1e6  # Large error if NPV is NaN
                return npv - target_npv
            except:
                return 1e6
        
        # Find breakeven price multiplier
        try:
            # Try brentq first (more reliable)
            multiplier = brentq(
                npv_error,
                a=0.1,  # 10% of base price
                b=5.0,  # 500% of base price
                xtol=tolerance
            )
        except:
            # Fallback to fsolve
            try:
                multiplier = fsolve(npv_error, [1.0])[0]
            except:
                return {
                    'breakeven_price': np.nan,
                    'base_price': avg_base_price,
                    'price_multiplier': np.nan,
                    'target_npv': target_npv,
                    'error': 'Could not solve for breakeven price'
                }
        
        breakeven_price = avg_base_price * multiplier
        
        return {
            'breakeven_price': breakeven_price,
            'base_price': avg_base_price,
            'price_multiplier': multiplier,
            'target_npv': target_npv
        }
    
    def calculate_breakeven_volume(
        self,
        data: pd.DataFrame,
        streaming_percentage: float,
        target_npv: float = 0.0,
        tolerance: float = 1e-4
    ) -> Dict:
        """
        Calculate breakeven credit volume (volume needed for target NPV).
        
        Parameters:
        -----------
        data : pd.DataFrame
            Base project data
        streaming_percentage : float
            Streaming percentage to use
        target_npv : float
            Target NPV (default: 0 for true breakeven)
        tolerance : float
            Optimization tolerance
            
        Returns:
        --------
        Dict
            Dictionary containing:
            - 'breakeven_volume_multiplier': Volume multiplier needed
            - 'base_volume': Average base volume from data
            - 'target_npv': Target NPV used
        """
        # Get average base volume
        base_volumes = data['carbon_credits_gross']
        avg_base_volume = base_volumes[base_volumes > 0].mean()
        
        if pd.isna(avg_base_volume) or avg_base_volume <= 0:
            return {
                'breakeven_volume_multiplier': np.nan,
                'base_volume': 0.0,
                'target_npv': target_npv,
                'error': 'No valid base volumes found'
            }
        
        # Create error function
        def npv_error(volume_multiplier: float) -> float:
            modified_data = data.copy()
            modified_data['carbon_credits_gross'] = base_volumes * volume_multiplier
            
            try:
                results = self.dcf_calculator.run_dcf(modified_data, streaming_percentage)
                npv = results['npv']
                if pd.isna(npv):
                    return 1e6
                return npv - target_npv
            except:
                return 1e6
        
        # Find breakeven volume multiplier
        try:
            multiplier = brentq(
                npv_error,
                a=0.1,  # 10% of base volume
                b=5.0,  # 500% of base volume
                xtol=tolerance
            )
        except:
            try:
                multiplier = fsolve(npv_error, [1.0])[0]
            except:
                return {
                    'breakeven_volume_multiplier': np.nan,
                    'base_volume': avg_base_volume,
                    'target_npv': target_npv,
                    'error': 'Could not solve for breakeven volume'
                }
        
        return {
            'breakeven_volume_multiplier': multiplier,
            'base_volume': avg_base_volume,
            'target_npv': target_npv
        }
    
    def calculate_breakeven_streaming(
        self,
        data: pd.DataFrame,
        target_npv: float = 0.0,
        tolerance: float = 1e-4
    ) -> Dict:
        """
        Calculate breakeven streaming percentage (streaming % needed for target NPV).
        
        Parameters:
        -----------
        data : pd.DataFrame
            Base project data
        target_npv : float
            Target NPV (default: 0 for true breakeven)
        tolerance : float
            Optimization tolerance
            
        Returns:
        --------
        Dict
            Dictionary containing:
            - 'breakeven_streaming': Streaming percentage needed
            - 'target_npv': Target NPV used
        """
        # Create error function
        def npv_error(streaming_pct: float) -> float:
            try:
                results = self.dcf_calculator.run_dcf(data, streaming_pct)
                npv = results['npv']
                if pd.isna(npv):
                    return 1e6
                return npv - target_npv
            except:
                return 1e6
        
        # Find breakeven streaming percentage
        try:
            streaming = brentq(
                npv_error,
                a=0.01,  # 1% minimum
                b=1.0,   # 100% maximum
                xtol=tolerance
            )
        except:
            try:
                streaming = fsolve(npv_error, [0.5])[0]
                streaming = max(0.01, min(1.0, streaming))  # Clamp to valid range
            except:
                return {
                    'breakeven_streaming': np.nan,
                    'target_npv': target_npv,
                    'error': 'Could not solve for breakeven streaming'
                }
        
        return {
            'breakeven_streaming': streaming,
            'target_npv': target_npv
        }
    
    def calculate_all_breakevens(
        self,
        data: pd.DataFrame,
        streaming_percentage: float,
        target_npv: float = 0.0
    ) -> Dict:
        """
        Calculate all breakeven points at once.
        
        Parameters:
        -----------
        data : pd.DataFrame
            Base project data
        streaming_percentage : float
            Streaming percentage to use
        target_npv : float
            Target NPV (default: 0 for true breakeven)
            
        Returns:
        --------
        Dict
            Dictionary with all breakeven calculations
        """
        return {
            'breakeven_price': self.calculate_breakeven_price(
                data, streaming_percentage, target_npv
            ),
            'breakeven_volume': self.calculate_breakeven_volume(
                data, streaming_percentage, target_npv
            ),
            'breakeven_streaming': self.calculate_breakeven_streaming(
                data, target_npv
            ),
            'target_npv': target_npv
        }

