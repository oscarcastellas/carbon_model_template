"""
Monte Carlo Simulation Module: Dual-variable stochastic modeling.

This module provides Monte Carlo simulation functionality to assess
probabilistic risk on IRR and NPV using stochastic price and volume paths.
"""

import pandas as pd
import numpy as np
from typing import Dict, List, Tuple, Optional
from .dcf_calculator import DCFCalculator
from .irr_calculator import IRRCalculator


class MonteCarloSimulator:
    """
    Performs Monte Carlo simulation for carbon credit streaming models.
    
    Simulates stochastic price growth and volume delivery to assess
    probabilistic risk on IRR and NPV.
    """
    
    def __init__(
        self,
        dcf_calculator: DCFCalculator,
        irr_calculator: IRRCalculator
    ):
        """
        Initialize the Monte Carlo Simulator.
        
        Parameters:
        -----------
        dcf_calculator : DCFCalculator
            DCF calculator instance
        irr_calculator : IRRCalculator
            IRR calculator instance
        """
        self.dcf_calculator = dcf_calculator
        self.irr_calculator = irr_calculator
    
    def generate_price_path(
        self,
        base_prices: pd.Series,
        price_growth_base: float,
        price_growth_std_dev: float,
        num_years: int = 20,
        use_percentage_variation: bool = False
    ) -> pd.Series:
        """
        Generate a stochastic 20-year price path based on original price forecasts.
        
        This method uses YOUR original carbon price forecasts from the input data as
        the base, then applies stochastic variations around those forecasts. This ensures
        that the Monte Carlo simulation respects your price forecast assumptions while
        modeling uncertainty.
        
        Two modes available:
        1. Growth-rate-based (default): Applies stochastic deviations to the annual
           growth rates implied by your original price curve
        2. Percentage-based: Applies stochastic percentage multipliers directly to prices
        
        Parameters:
        -----------
        base_prices : pd.Series
            YOUR original carbon price forecasts (Year 1 to Year 20) from input data.
            These are used as the base/center of the stochastic distribution.
        price_growth_base : float
            Mean annual growth rate deviation (e.g., 0.03 for 3%).
            When use_percentage_variation=False, this is the mean deviation added to
            your original growth rates. When True, this parameter is not used.
        price_growth_std_dev : float
            Standard deviation of annual growth rate deviation (e.g., 0.02 for 2%).
            Controls the volatility/uncertainty around your original forecasts.
        num_years : int
            Number of years (default: 20)
        use_percentage_variation : bool
            If True, applies percentage multipliers directly to prices.
            If False (default), applies stochastic deviations to growth rates.
            
        Returns:
        --------
        pd.Series
            Stochastic price path indexed by Year, centered around your original forecasts
        """
        if use_percentage_variation:
            # Mode 2: Percentage-based variation
            # Apply stochastic percentage multipliers directly to original prices
            price_path = base_prices.copy()
            
            # Generate percentage multipliers (e.g., 0.98 to 1.02 for ±2% variation)
            # Convert std_dev from growth rate to percentage multiplier
            # If std_dev = 0.02 (2% growth volatility), we want ~2% price variation
            percentage_std = price_growth_std_dev
            
            # Generate multipliers: 1.0 ± random variation
            multipliers = np.random.normal(
                loc=1.0,  # Mean multiplier is 1.0 (no bias from original)
                scale=percentage_std,
                size=num_years
            )
            
            # Ensure multipliers are positive
            multipliers = np.maximum(multipliers, 0.01)
            
            # Apply multipliers to original prices
            price_path = base_prices * multipliers
            
            return price_path
        
        else:
            # Mode 1: Growth-rate-based variation (default)
            # Uses your original price forecasts and applies stochastic deviations
            # to the growth rates implied by your curve
            price_path = base_prices.copy()
            
            # Generate stochastic growth rate deviations from normal distribution
            # These will be ADDED to your original curve's growth rates
            growth_deviations = np.random.normal(
                loc=0.0,  # Mean deviation is 0 (centered on your original forecasts)
                scale=price_growth_std_dev,
                size=num_years - 1  # One less than years (growth between years)
            )
            
            # Apply stochastic deviations to YOUR original price curve
            # Start from Year 2 and apply growth deviations
            for i in range(1, num_years):
                base_prev = base_prices.iloc[i - 1]
                base_curr = base_prices.iloc[i]
                
                # Calculate YOUR original growth rate from your forecast curve
                if base_prev > 0:
                    base_growth = (base_curr / base_prev) - 1
                else:
                    # If previous price is 0, use the mean growth rate
                    base_growth = price_growth_base
                
                # Apply stochastic deviation: YOUR growth + random deviation
                # This adds variability around YOUR original forecast
                stochastic_growth = base_growth + growth_deviations[i - 1]
                
                # Calculate new price
                prev_price = price_path.iloc[i - 1]
                if prev_price > 0:
                    price_path.iloc[i] = prev_price * (1 + stochastic_growth)
                else:
                    # If previous price is 0, use base price
                    price_path.iloc[i] = base_curr
            
            return price_path
    
    def generate_volume_path(
        self,
        base_volumes: pd.Series,
        volume_multiplier_base: float,
        volume_std_dev: float,
        num_years: int = 20
    ) -> pd.Series:
        """
        Generate a stochastic 20-year volume path using yearly multipliers.
        
        Each year's multiplier is drawn from a normal distribution.
        
        Parameters:
        -----------
        base_volumes : pd.Series
            Base carbon credit volumes (Year 1 to Year 20)
        volume_multiplier_base : float
            Mean volume multiplier (e.g., 1.0)
        volume_std_dev : float
            Standard deviation of volume multiplier (e.g., 0.15 for 15%)
        num_years : int
            Number of years (default: 20)
            
        Returns:
        --------
        pd.Series
            Stochastic volume path indexed by Year
        """
        # Generate yearly multipliers from normal distribution
        multipliers = np.random.normal(
            loc=volume_multiplier_base,
            scale=volume_std_dev,
            size=num_years
        )
        
        # Ensure multipliers are positive (truncate at 0)
        multipliers = np.maximum(multipliers, 0.01)  # Minimum 1% to avoid zero/negative
        
        # Apply multipliers to base volumes
        volume_path = base_volumes * multipliers
        
        return volume_path
    
    def run_single_simulation(
        self,
        base_data: pd.DataFrame,
        streaming_percentage: float,
        price_growth_base: float,
        price_growth_std_dev: float,
        volume_multiplier_base: float,
        volume_std_dev: float,
        use_percentage_variation: bool = False
    ) -> Tuple[float, float]:
        """
        Run a single Monte Carlo simulation.
        
        Uses YOUR original price forecasts from base_data as the center of the
        stochastic distribution, then applies variability around those forecasts.
        
        Parameters:
        -----------
        base_data : pd.DataFrame
            Base input data with YOUR original price forecasts and credit volumes.
            The simulation builds sensitivities around these original values.
        streaming_percentage : float
            Target streaming percentage to use
        price_growth_base : float
            Mean annual price growth rate (used when use_percentage_variation=False)
        price_growth_std_dev : float
            Standard deviation controlling price volatility around YOUR forecasts
        volume_multiplier_base : float
            Mean volume multiplier (typically 1.0 to center on original volumes)
        volume_std_dev : float
            Standard deviation of volume multiplier (volatility around original volumes)
        use_percentage_variation : bool
            If True, applies percentage multipliers directly to prices.
            If False (default), applies stochastic deviations to growth rates.
            
        Returns:
        --------
        Tuple[float, float]
            (IRR, NPV) for this simulation
        """
        # Create modified data for this simulation
        sim_data = base_data.copy()
        
        # Generate stochastic price path based on YOUR original price forecasts
        sim_data['base_carbon_price'] = self.generate_price_path(
            base_prices=base_data['base_carbon_price'],
            price_growth_base=price_growth_base,
            price_growth_std_dev=price_growth_std_dev,
            use_percentage_variation=use_percentage_variation
        )
        
        # Generate stochastic volume path
        sim_data['carbon_credits_gross'] = self.generate_volume_path(
            base_volumes=base_data['carbon_credits_gross'],
            volume_multiplier_base=volume_multiplier_base,
            volume_std_dev=volume_std_dev
        )
        
        # Calculate DCF with stochastic data
        try:
            dcf_results = self.dcf_calculator.run_dcf(
                data=sim_data,
                streaming_percentage=streaming_percentage
            )
            
            irr = dcf_results['irr']
            npv = dcf_results['npv']
            
            # Handle NaN values
            if pd.isna(irr) or not np.isfinite(irr):
                irr = np.nan
            if pd.isna(npv) or not np.isfinite(npv):
                npv = np.nan
            
            return float(irr), float(npv)
            
        except Exception:
            # If calculation fails, return NaN
            return np.nan, np.nan
    
    def run_monte_carlo(
        self,
        base_data: pd.DataFrame,
        streaming_percentage: float,
        price_growth_base: float,
        price_growth_std_dev: float,
        volume_multiplier_base: float,
        volume_std_dev: float,
        simulations: int = 5000,
        random_seed: Optional[int] = None,
        use_percentage_variation: bool = False
    ) -> Dict:
        """
        Run Monte Carlo simulation with dual-variable stochastic modeling.
        
        This simulation uses YOUR original price forecasts and credit volumes from
        base_data as the center of the stochastic distribution. It then applies
        variability around those forecasts to assess probabilistic risk.
        
        Parameters:
        -----------
        base_data : pd.DataFrame
            Base input data with YOUR original price forecasts and credit volumes.
            The simulation builds sensitivities around these original values.
        streaming_percentage : float
            Target streaming percentage (typically from goal-seeking)
        price_growth_base : float
            Mean annual price growth rate (e.g., 0.03 for 3%).
            Used when use_percentage_variation=False.
        price_growth_std_dev : float
            Standard deviation controlling price volatility around YOUR forecasts
            (e.g., 0.02 for 2% volatility)
        volume_multiplier_base : float
            Mean volume multiplier (e.g., 1.0 to center on original volumes)
        volume_std_dev : float
            Standard deviation of volume multiplier (e.g., 0.15 for 15% volatility)
        simulations : int
            Number of simulations to run (default: 5000)
        random_seed : int, optional
            Random seed for reproducibility
        use_percentage_variation : bool
            If True, applies percentage multipliers directly to prices.
            If False (default), applies stochastic deviations to growth rates.
            
        Returns:
        --------
        Dict
            Dictionary containing:
            - 'irr_series': Array of simulated IRRs
            - 'npv_series': Array of simulated NPVs
            - 'mc_mean_irr': Mean IRR across simulations
            - 'mc_mean_npv': Mean NPV across simulations
            - 'mc_p10_irr': 10th percentile IRR
            - 'mc_p90_irr': 90th percentile IRR
            - 'mc_p10_npv': 10th percentile NPV
            - 'mc_p90_npv': 90th percentile NPV
            - 'mc_std_irr': Standard deviation of IRR
            - 'mc_std_npv': Standard deviation of NPV
        """
        if random_seed is not None:
            np.random.seed(random_seed)
        
        irr_results = []
        npv_results = []
        
        # Run simulations
        for i in range(simulations):
            if (i + 1) % 1000 == 0:
                print(f"Running simulation {i + 1}/{simulations}...")
            
            irr, npv = self.run_single_simulation(
                base_data=base_data,
                streaming_percentage=streaming_percentage,
                price_growth_base=price_growth_base,
                price_growth_std_dev=price_growth_std_dev,
                volume_multiplier_base=volume_multiplier_base,
                volume_std_dev=volume_std_dev,
                use_percentage_variation=use_percentage_variation
            )
            
            irr_results.append(irr)
            npv_results.append(npv)
        
        # Convert to numpy arrays
        irr_array = np.array(irr_results)
        npv_array = np.array(npv_results)
        
        # Remove NaN values for statistics
        irr_valid = irr_array[~np.isnan(irr_array)]
        npv_valid = npv_array[~np.isnan(npv_array)]
        
        # Calculate statistics
        results = {
            'irr_series': irr_array,
            'npv_series': npv_array,
            'mc_mean_irr': float(np.mean(irr_valid)) if len(irr_valid) > 0 else np.nan,
            'mc_mean_npv': float(np.mean(npv_valid)) if len(npv_valid) > 0 else np.nan,
            'mc_p10_irr': float(np.percentile(irr_valid, 10)) if len(irr_valid) > 0 else np.nan,
            'mc_p90_irr': float(np.percentile(irr_valid, 90)) if len(irr_valid) > 0 else np.nan,
            'mc_p10_npv': float(np.percentile(npv_valid, 10)) if len(npv_valid) > 0 else np.nan,
            'mc_p90_npv': float(np.percentile(npv_valid, 90)) if len(npv_valid) > 0 else np.nan,
            'mc_std_irr': float(np.std(irr_valid)) if len(irr_valid) > 0 else np.nan,
            'mc_std_npv': float(np.std(npv_valid)) if len(npv_valid) > 0 else np.nan,
            'simulations': simulations,
            'valid_simulations': len(irr_valid)
        }
        
        return results

