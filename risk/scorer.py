"""
Risk Score Calculator Module: Calculates overall risk score for projects.

This module provides a simple 0-100 risk score combining multiple risk factors
for quick project ranking and prioritization.
"""

from typing import Dict, Optional
import pandas as pd
import numpy as np


class RiskScoreCalculator:
    """
    Calculates overall risk score for carbon credit projects.
    
    Combines financial, volume, price, and operational risk factors
    into a single 0-100 score (lower = lower risk, higher = higher risk).
    """
    
    # Default weights (sum to 1.0)
    DEFAULT_WEIGHTS = {
        'financial_risk': 0.40,    # IRR, NPV, Payback
        'volume_risk': 0.25,        # Credit delivery uncertainty
        'price_risk': 0.20,         # Price volatility
        'operational_risk': 0.15     # Project complexity, costs
    }
    
    def __init__(self, weights: Dict = None):
        """
        Initialize Risk Score Calculator.
        
        Parameters:
        -----------
        weights : Dict, optional
            Custom weights for risk factors. If None, uses defaults.
            Weights should sum to 1.0.
        """
        if weights:
            # Normalize weights to sum to 1.0
            total = sum(weights.values())
            self.weights = {k: v/total for k, v in weights.items()}
        else:
            self.weights = self.DEFAULT_WEIGHTS.copy()
    
    def calculate_financial_risk(
        self,
        irr: float,
        npv: float,
        payback_period: Optional[float] = None
    ) -> float:
        """
        Calculate financial risk score (0-100).
        
        Lower IRR, negative NPV, long payback = higher risk.
        
        Parameters:
        -----------
        irr : float
            Internal Rate of Return
        npv : float
            Net Present Value
        payback_period : float, optional
            Payback period in years
            
        Returns:
        --------
        float
            Financial risk score (0-100)
        """
        risk_score = 0.0
        
        # IRR risk (0-40 points)
        if pd.isna(irr):
            risk_score += 40  # Maximum risk if no IRR
        elif irr < 0.10:
            risk_score += 40
        elif irr < 0.15:
            risk_score += 30
        elif irr < 0.20:
            risk_score += 15
        elif irr < 0.25:
            risk_score += 5
        # IRR >= 25% = 0 risk points
        
        # NPV risk (0-35 points)
        if pd.isna(npv):
            risk_score += 35
        elif npv < 0:
            risk_score += 35
        elif npv < 5_000_000:
            risk_score += 25
        elif npv < 10_000_000:
            risk_score += 15
        elif npv < 20_000_000:
            risk_score += 5
        # NPV >= $20M = 0 risk points
        
        # Payback risk (0-25 points)
        if payback_period is not None and not pd.isna(payback_period):
            if payback_period > 15:
                risk_score += 25
            elif payback_period > 12:
                risk_score += 20
            elif payback_period > 10:
                risk_score += 10
            elif payback_period > 8:
                risk_score += 5
            # Payback <= 8 years = 0 risk points
        
        # Normalize to 0-100 scale
        return min(100.0, risk_score)
    
    def calculate_volume_risk(
        self,
        credit_volumes: pd.Series,
        volume_volatility: Optional[float] = None
    ) -> float:
        """
        Calculate volume risk score (0-100).
        
        Low volumes, high volatility, many zero years = higher risk.
        
        Parameters:
        -----------
        credit_volumes : pd.Series
            Annual credit volumes
        volume_volatility : float, optional
            Standard deviation of volume multiplier (from Monte Carlo)
            
        Returns:
        --------
        float
            Volume risk score (0-100)
        """
        risk_score = 0.0
        
        # Total volume risk (0-40 points)
        total_volume = credit_volumes.sum()
        if total_volume < 1_000_000:
            risk_score += 40
        elif total_volume < 5_000_000:
            risk_score += 30
        elif total_volume < 10_000_000:
            risk_score += 20
        elif total_volume < 20_000_000:
            risk_score += 10
        # Total >= 20M = 0 risk points
        
        # Zero years risk (0-30 points)
        zero_years = (credit_volumes == 0).sum()
        if zero_years > 10:
            risk_score += 30
        elif zero_years > 5:
            risk_score += 20
        elif zero_years > 2:
            risk_score += 10
        # Zero years <= 2 = 0 risk points
        
        # Volatility risk (0-30 points)
        if volume_volatility is not None:
            if volume_volatility > 0.25:  # 25% volatility
                risk_score += 30
            elif volume_volatility > 0.15:
                risk_score += 20
            elif volume_volatility > 0.10:
                risk_score += 10
            # Volatility <= 10% = 0 risk points
        
        return min(100.0, risk_score)
    
    def calculate_price_risk(
        self,
        base_prices: pd.Series,
        price_volatility: Optional[float] = None
    ) -> float:
        """
        Calculate price risk score (0-100).
        
        Low prices, high volatility = higher risk.
        
        Parameters:
        -----------
        base_prices : pd.Series
            Base carbon prices
        price_volatility : float, optional
            Standard deviation of price growth (from Monte Carlo)
            
        Returns:
        --------
        float
            Price risk score (0-100)
        """
        risk_score = 0.0
        
        # Average price risk (0-50 points)
        avg_price = base_prices[base_prices > 0].mean()
        if pd.isna(avg_price) or avg_price < 20:
            risk_score += 50
        elif avg_price < 30:
            risk_score += 40
        elif avg_price < 40:
            risk_score += 25
        elif avg_price < 50:
            risk_score += 10
        # Price >= $50 = 0 risk points
        
        # Volatility risk (0-50 points)
        if price_volatility is not None:
            if price_volatility > 0.05:  # 5% volatility
                risk_score += 50
            elif price_volatility > 0.03:
                risk_score += 30
            elif price_volatility > 0.02:
                risk_score += 15
            # Volatility <= 2% = 0 risk points
        
        return min(100.0, risk_score)
    
    def calculate_operational_risk(
        self,
        project_costs: pd.Series,
        total_investment: Optional[float] = None
    ) -> float:
        """
        Calculate operational risk score (0-100).
        
        High costs, large investment = higher risk.
        
        Parameters:
        -----------
        project_costs : pd.Series
            Annual project implementation costs
        total_investment : float, optional
            Total Rubicon investment
            
        Returns:
        --------
        float
            Operational risk score (0-100)
        """
        risk_score = 0.0
        
        # Total costs risk (0-60 points)
        total_costs = abs(project_costs.sum())
        if total_costs > 200_000_000:
            risk_score += 60
        elif total_costs > 100_000_000:
            risk_score += 40
        elif total_costs > 50_000_000:
            risk_score += 25
        elif total_costs > 25_000_000:
            risk_score += 10
        # Costs <= $25M = 0 risk points
        
        # Investment size risk (0-40 points)
        if total_investment is not None:
            if total_investment > 50_000_000:
                risk_score += 40
            elif total_investment > 30_000_000:
                risk_score += 25
            elif total_investment > 20_000_000:
                risk_score += 15
            elif total_investment > 10_000_000:
                risk_score += 5
            # Investment <= $10M = 0 risk points
        
        return min(100.0, risk_score)
    
    def calculate_overall_risk_score(
        self,
        irr: float,
        npv: float,
        payback_period: Optional[float] = None,
        credit_volumes: Optional[pd.Series] = None,
        base_prices: Optional[pd.Series] = None,
        project_costs: Optional[pd.Series] = None,
        volume_volatility: Optional[float] = None,
        price_volatility: Optional[float] = None,
        total_investment: Optional[float] = None
    ) -> Dict:
        """
        Calculate overall risk score combining all factors.
        
        Parameters:
        -----------
        irr : float
            Internal Rate of Return
        npv : float
            Net Present Value
        payback_period : float, optional
            Payback period in years
        credit_volumes : pd.Series, optional
            Annual credit volumes
        base_prices : pd.Series, optional
            Base carbon prices
        project_costs : pd.Series, optional
            Annual project costs
        volume_volatility : float, optional
            Volume volatility (from Monte Carlo)
        price_volatility : float, optional
            Price volatility (from Monte Carlo)
        total_investment : float, optional
            Total Rubicon investment
            
        Returns:
        --------
        Dict
            Dictionary containing:
            - 'overall_risk_score': Overall score (0-100)
            - 'financial_risk': Financial risk score
            - 'volume_risk': Volume risk score
            - 'price_risk': Price risk score
            - 'operational_risk': Operational risk score
            - 'risk_category': 'Low', 'Medium', or 'High'
        """
        # Calculate individual risk scores
        financial_risk = self.calculate_financial_risk(irr, npv, payback_period)
        
        volume_risk = 0.0
        if credit_volumes is not None:
            volume_risk = self.calculate_volume_risk(credit_volumes, volume_volatility)
        
        price_risk = 0.0
        if base_prices is not None:
            price_risk = self.calculate_price_risk(base_prices, price_volatility)
        
        operational_risk = 0.0
        if project_costs is not None:
            operational_risk = self.calculate_operational_risk(project_costs, total_investment)
        
        # Calculate weighted overall score
        overall_score = (
            financial_risk * self.weights['financial_risk'] +
            volume_risk * self.weights['volume_risk'] +
            price_risk * self.weights['price_risk'] +
            operational_risk * self.weights['operational_risk']
        )
        
        # Categorize risk
        if overall_score < 30:
            risk_category = 'Low'
        elif overall_score < 60:
            risk_category = 'Medium'
        else:
            risk_category = 'High'
        
        return {
            'overall_risk_score': round(overall_score, 1),
            'financial_risk': round(financial_risk, 1),
            'volume_risk': round(volume_risk, 1),
            'price_risk': round(price_risk, 1),
            'operational_risk': round(operational_risk, 1),
            'risk_category': risk_category
        }

