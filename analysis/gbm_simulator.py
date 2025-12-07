"""
Geometric Brownian Motion (GBM) Price Simulator Module.

This module provides sophisticated price volatility modeling using Geometric
Brownian Motion, the industry-standard approach for modeling asset price volatility.

GBM models price as: dS = μS dt + σS dW

Where:
- S = price
- μ = drift (expected return)
- σ = volatility (standard deviation)
- dW = Wiener process (random walk)
"""

import pandas as pd
import numpy as np
from typing import Optional


class GBMPriceSimulator:
    """
    Geometric Brownian Motion price simulator for carbon credits.
    
    Uses Euler-Maruyama discretization to generate stochastic price paths
    that follow a geometric Brownian motion process.
    """
    
    def __init__(self):
        """Initialize the GBM Price Simulator."""
        pass
    
    def generate_gbm_path(
        self,
        initial_price: float,
        drift: float,
        volatility: float,
        num_years: int = 20,
        time_steps: int = 20,
        random_seed: Optional[int] = None
    ) -> pd.Series:
        """
        Generate price path using Geometric Brownian Motion.
        
        Uses Euler-Maruyama discretization:
        S(t+Δt) = S(t) * exp((μ - σ²/2)Δt + σ√Δt * Z)
        
        Where:
        - S(t) = price at time t
        - μ = drift (expected annual return)
        - σ = volatility (annual standard deviation)
        - Δt = time step (1 year)
        - Z ~ N(0,1) = standard normal random variable
        
        Parameters:
        -----------
        initial_price : float
            Starting price (S₀)
        drift : float
            Annual expected return (μ), e.g., 0.03 for 3%
        volatility : float
            Annual volatility (σ), e.g., 0.15 for 15%
        num_years : int
            Number of years to simulate (default: 20)
        time_steps : int
            Number of time steps (default: 20, one per year)
        random_seed : int, optional
            Random seed for reproducibility
            
        Returns:
        --------
        pd.Series
            Price path indexed by year (1 to num_years)
        """
        if random_seed is not None:
            np.random.seed(random_seed)
        
        # Time step size (in years)
        dt = num_years / time_steps
        
        # Initialize price array
        prices = np.zeros(time_steps + 1)
        prices[0] = initial_price
        
        # Generate random shocks (standard normal)
        random_shocks = np.random.normal(0, 1, time_steps)
        
        # Euler-Maruyama discretization
        # S(t+Δt) = S(t) * exp((μ - σ²/2)Δt + σ√Δt * Z)
        for i in range(time_steps):
            drift_term = (drift - 0.5 * volatility ** 2) * dt
            diffusion_term = volatility * np.sqrt(dt) * random_shocks[i]
            prices[i + 1] = prices[i] * np.exp(drift_term + diffusion_term)
        
        # Create Series indexed by year
        years = range(1, num_years + 1)
        # If time_steps != num_years, interpolate or take every nth value
        if time_steps == num_years:
            price_series = pd.Series(prices[1:], index=years)
        else:
            # Interpolate to match num_years
            indices = np.linspace(1, time_steps, num_years, dtype=int)
            price_series = pd.Series(prices[1:][indices], index=years)
        
        return price_series
    
    def generate_gbm_path_from_base(
        self,
        base_prices: pd.Series,
        drift: float,
        volatility: float,
        random_seed: Optional[int] = None
    ) -> pd.Series:
        """
        Generate GBM path starting from base price series.
        
        Uses the first non-zero price as initial price, then applies
        GBM process to generate stochastic variations.
        
        Parameters:
        -----------
        base_prices : pd.Series
            Base price forecast (used for initial price)
        drift : float
            Annual expected return (μ)
        volatility : float
            Annual volatility (σ)
        random_seed : int, optional
            Random seed for reproducibility
            
        Returns:
        --------
        pd.Series
            Stochastic price path with same index as base_prices
        """
        # Find first non-zero price as initial price
        initial_price = base_prices[base_prices > 0].iloc[0] if (base_prices > 0).any() else base_prices.iloc[0]
        
        num_years = len(base_prices)
        
        # Generate GBM path
        gbm_path = self.generate_gbm_path(
            initial_price=initial_price,
            drift=drift,
            volatility=volatility,
            num_years=num_years,
            time_steps=num_years,
            random_seed=random_seed
        )
        
        # Match index to base_prices
        gbm_path.index = base_prices.index
        
        return gbm_path
    
    def calculate_implied_volatility(
        self,
        price_series: pd.Series
    ) -> float:
        """
        Calculate implied volatility from a price series.
        
        Estimates volatility from historical price returns.
        
        Parameters:
        -----------
        price_series : pd.Series
            Historical price series
            
        Returns:
        --------
        float
            Estimated annual volatility
        """
        # Calculate returns
        returns = price_series.pct_change().dropna()
        
        # Annualize volatility (assuming daily/period returns)
        # If prices are annual, volatility is already annual
        # If prices are more frequent, need to annualize
        volatility = returns.std()
        
        # Annualize if needed (assuming 252 trading days or adjust as needed)
        # For annual data, no adjustment needed
        if len(price_series) > 1:
            periods_per_year = len(price_series) / ((price_series.index[-1] - price_series.index[0]) + 1)
            if periods_per_year > 1:
                volatility = volatility * np.sqrt(periods_per_year)
        
        return volatility
    
    def calculate_implied_drift(
        self,
        price_series: pd.Series
    ) -> float:
        """
        Calculate implied drift from a price series.
        
        Estimates expected return from historical price growth.
        
        Parameters:
        -----------
        price_series : pd.Series
            Historical price series
            
        Returns:
        --------
        float
            Estimated annual drift (expected return)
        """
        # Calculate total return
        if len(price_series) < 2:
            return 0.0
        
        initial_price = price_series.iloc[0]
        final_price = price_series.iloc[-1]
        num_years = len(price_series) - 1
        
        if initial_price > 0 and num_years > 0:
            # Annualized return
            total_return = (final_price / initial_price) ** (1.0 / num_years) - 1
            return total_return
        
        return 0.0

