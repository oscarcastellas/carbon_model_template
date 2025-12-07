"""
Risk Flagging Module: Automatically flags projects with risk indicators.

This module provides simple risk flagging based on financial metrics
to quickly identify projects that need attention.
"""

from typing import Dict, List, Tuple
import pandas as pd


class RiskFlagger:
    """
    Flags projects with risk indicators based on financial metrics.
    
    Provides simple red/yellow/green risk scoring for quick project assessment.
    """
    
    # Risk thresholds (configurable)
    RED_FLAG_THRESHOLDS = {
        'irr_min': 0.15,           # IRR below 15%
        'npv_min': 0,              # Negative NPV
        'payback_max': 15.0,        # Payback > 15 years
        'irr_volatility_high': 0.05,  # IRR std dev > 5%
    }
    
    YELLOW_FLAG_THRESHOLDS = {
        'irr_min': 0.18,           # IRR below 18%
        'npv_min': 5_000_000,      # NPV < $5M
        'payback_max': 12.0,        # Payback > 12 years
        'irr_volatility_high': 0.03,  # IRR std dev > 3%
    }
    
    def __init__(
        self,
        red_thresholds: Dict = None,
        yellow_thresholds: Dict = None
    ):
        """
        Initialize Risk Flagger.
        
        Parameters:
        -----------
        red_thresholds : Dict, optional
            Custom red flag thresholds
        yellow_thresholds : Dict, optional
            Custom yellow flag thresholds
        """
        if red_thresholds:
            self.RED_FLAG_THRESHOLDS.update(red_thresholds)
        if yellow_thresholds:
            self.YELLOW_FLAG_THRESHOLDS.update(yellow_thresholds)
    
    def flag_risks(
        self,
        irr: float,
        npv: float,
        payback_period: float = None,
        irr_volatility: float = None,
        credit_volumes: pd.Series = None,
        project_costs: pd.Series = None
    ) -> Dict:
        """
        Flag risks for a project based on financial metrics.
        
        Parameters:
        -----------
        irr : float
            Internal Rate of Return
        npv : float
            Net Present Value
        payback_period : float, optional
            Payback period in years
        irr_volatility : float, optional
            Standard deviation of IRR (from Monte Carlo)
        credit_volumes : pd.Series, optional
            Annual credit volumes for volume risk assessment
        project_costs : pd.Series, optional
            Annual project costs for cost risk assessment
            
        Returns:
        --------
        Dict
            Dictionary containing:
            - 'risk_level': 'red', 'yellow', or 'green'
            - 'flags': List of flag descriptions
            - 'red_flags': List of red flag issues
            - 'yellow_flags': List of yellow flag warnings
            - 'green_flags': List of positive indicators
        """
        red_flags = []
        yellow_flags = []
        green_flags = []
        
        # Check IRR
        if pd.isna(irr) or irr < self.RED_FLAG_THRESHOLDS['irr_min']:
            red_flags.append(f"Low IRR: {irr:.2%} (below {self.RED_FLAG_THRESHOLDS['irr_min']:.0%})")
        elif irr < self.YELLOW_FLAG_THRESHOLDS['irr_min']:
            yellow_flags.append(f"IRR below target: {irr:.2%} (target: {self.YELLOW_FLAG_THRESHOLDS['irr_min']:.0%})")
        else:
            green_flags.append(f"Strong IRR: {irr:.2%}")
        
        # Check NPV
        if pd.isna(npv) or npv < self.RED_FLAG_THRESHOLDS['npv_min']:
            red_flags.append(f"Negative or low NPV: ${npv:,.0f}")
        elif npv < self.YELLOW_FLAG_THRESHOLDS['npv_min']:
            yellow_flags.append(f"Moderate NPV: ${npv:,.0f} (below ${self.YELLOW_FLAG_THRESHOLDS['npv_min']:,.0f})")
        else:
            green_flags.append(f"Strong NPV: ${npv:,.0f}")
        
        # Check Payback Period
        if payback_period is not None:
            if pd.isna(payback_period) or payback_period > self.RED_FLAG_THRESHOLDS['payback_max']:
                red_flags.append(f"Long payback: {payback_period:.1f} years (exceeds {self.RED_FLAG_THRESHOLDS['payback_max']:.0f} years)")
            elif payback_period > self.YELLOW_FLAG_THRESHOLDS['payback_max']:
                yellow_flags.append(f"Extended payback: {payback_period:.1f} years")
            else:
                green_flags.append(f"Reasonable payback: {payback_period:.1f} years")
        
        # Check IRR Volatility (from Monte Carlo)
        if irr_volatility is not None:
            if irr_volatility > self.RED_FLAG_THRESHOLDS['irr_volatility_high']:
                red_flags.append(f"High IRR volatility: {irr_volatility:.2%} (std dev)")
            elif irr_volatility > self.YELLOW_FLAG_THRESHOLDS['irr_volatility_high']:
                yellow_flags.append(f"Moderate IRR volatility: {irr_volatility:.2%}")
            else:
                green_flags.append(f"Low IRR volatility: {irr_volatility:.2%}")
        
        # Check Credit Volumes (if provided)
        if credit_volumes is not None:
            total_credits = credit_volumes.sum()
            if total_credits < 1_000_000:  # Less than 1M credits
                yellow_flags.append(f"Low total credits: {total_credits:,.0f}")
            elif total_credits > 50_000_000:  # Very high
                green_flags.append(f"High credit volume: {total_credits:,.0f}")
            
            # Check for zero years
            zero_years = (credit_volumes == 0).sum()
            if zero_years > 5:
                yellow_flags.append(f"Many zero-credit years: {zero_years} years")
        
        # Check Project Costs (if provided)
        if project_costs is not None:
            total_costs = abs(project_costs.sum())
            if total_costs > 200_000_000:  # Very high costs
                yellow_flags.append(f"High total costs: ${total_costs:,.0f}")
        
        # Determine overall risk level
        if len(red_flags) > 0:
            risk_level = 'red'
        elif len(yellow_flags) > 0:
            risk_level = 'yellow'
        else:
            risk_level = 'green'
        
        # Combine all flags
        all_flags = red_flags + yellow_flags
        
        return {
            'risk_level': risk_level,
            'flags': all_flags,
            'red_flags': red_flags,
            'yellow_flags': yellow_flags,
            'green_flags': green_flags,
            'flag_count': {
                'red': len(red_flags),
                'yellow': len(yellow_flags),
                'green': len(green_flags)
            }
        }
    
    def get_risk_summary(self, risk_flags: Dict) -> str:
        """
        Get a human-readable risk summary.
        
        Parameters:
        -----------
        risk_flags : Dict
            Output from flag_risks() method
            
        Returns:
        --------
        str
            Formatted risk summary
        """
        risk_level = risk_flags['risk_level']
        red_count = risk_flags['flag_count']['red']
        yellow_count = risk_flags['flag_count']['yellow']
        green_count = risk_flags['flag_count']['green']
        
        summary = f"Risk Level: {risk_level.upper()}\n"
        summary += f"Red Flags: {red_count}, Yellow Flags: {yellow_count}, Green Indicators: {green_count}\n\n"
        
        if risk_flags['red_flags']:
            summary += "🚨 RED FLAGS:\n"
            for flag in risk_flags['red_flags']:
                summary += f"  • {flag}\n"
            summary += "\n"
        
        if risk_flags['yellow_flags']:
            summary += "⚠️  YELLOW FLAGS:\n"
            for flag in risk_flags['yellow_flags']:
                summary += f"  • {flag}\n"
            summary += "\n"
        
        if risk_flags['green_flags']:
            summary += "✅ POSITIVE INDICATORS:\n"
            for flag in risk_flags['green_flags']:
                summary += f"  • {flag}\n"
        
        return summary

