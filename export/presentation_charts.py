"""
Presentation Charts Module: Generates charts for investor committee presentations.

This module creates professional charts for the presentation sheets:
- Inputs & Assumptions
- Valuation Schedule
- Summary & Results

Charts are generated as PNG images and embedded into Excel using openpyxl.
"""

import os
import tempfile
from pathlib import Path
from typing import Dict, Optional, List
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')  # Non-interactive backend
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
import seaborn as sns

# Set style for professional charts
sns.set_style("whitegrid")
plt.rcParams['figure.facecolor'] = 'white'
plt.rcParams['axes.facecolor'] = 'white'
plt.rcParams['font.size'] = 10
plt.rcParams['axes.titlesize'] = 12
plt.rcParams['axes.labelsize'] = 10
plt.rcParams['xtick.labelsize'] = 9
plt.rcParams['ytick.labelsize'] = 9


class PresentationChartGenerator:
    """
    Generates professional charts for presentation sheets.
    """
    
    def __init__(self, temp_dir: Optional[str] = None):
        """
        Initialize chart generator.
        
        Parameters:
        -----------
        temp_dir : str, optional
            Directory to save temporary chart files. If None, uses system temp.
        """
        if temp_dir is None:
            self.temp_dir = tempfile.gettempdir()
        else:
            self.temp_dir = temp_dir
            os.makedirs(temp_dir, exist_ok=True)
        
        # Color scheme for professional presentations
        self.colors = {
            'primary': '#366092',      # Blue
            'secondary': '#70AD47',     # Green
            'accent': '#FFC000',        # Yellow/Orange
            'negative': '#C00000',      # Red
            'neutral': '#808080',       # Gray
            'light_blue': '#D9E1F2',
            'light_green': '#E2EFDA',
            'light_yellow': '#FFF2CC'
        }
    
    def create_assumptions_summary_chart(
        self,
        assumptions: Dict,
        streaming_pct: float,
        output_path: Optional[str] = None
    ) -> str:
        """
        Create key assumptions summary chart (pie or bar chart).
        
        Parameters:
        -----------
        assumptions : Dict
            Dictionary of assumptions
        streaming_pct : float
            Streaming percentage
        output_path : str, optional
            Path to save chart. If None, saves to temp file.
            
        Returns:
        --------
        str
            Path to saved chart image
        """
        if output_path is None:
            output_path = os.path.join(self.temp_dir, 'assumptions_summary.png')
        
        fig, ax = plt.subplots(figsize=(8, 6))
        
        # Prepare data
        labels = []
        values = []
        colors_list = []
        
        if 'rubicon_investment_total' in assumptions:
            labels.append('Investment Total')
            values.append(assumptions['rubicon_investment_total'] / 1e6)  # Convert to millions
            colors_list.append(self.colors['primary'])
        
        if 'wacc' in assumptions:
            labels.append('WACC')
            values.append(assumptions['wacc'] * 100)  # Convert to percentage
            colors_list.append(self.colors['secondary'])
        
        if 'investment_tenor' in assumptions:
            labels.append('Tenor (Years)')
            values.append(assumptions['investment_tenor'])
            colors_list.append(self.colors['accent'])
        
        labels.append('Streaming %')
        values.append(streaming_pct * 100)
        colors_list.append(self.colors['light_blue'])
        
        # Create horizontal bar chart
        y_pos = np.arange(len(labels))
        bars = ax.barh(y_pos, values, color=colors_list, edgecolor='black', linewidth=1.5)
        
        # Add value labels on bars
        for i, (bar, val) in enumerate(zip(bars, values)):
            if i == 0:  # Investment Total
                label_text = f'${val:.1f}M'
            elif i == 1:  # WACC
                label_text = f'{val:.1f}%'
            elif i == 2:  # Tenor
                label_text = f'{int(val)} years'
            else:  # Streaming
                label_text = f'{val:.1f}%'
            
            ax.text(bar.get_width() * 0.5, bar.get_y() + bar.get_height()/2,
                   label_text, ha='center', va='center', fontweight='bold', color='white')
        
        ax.set_yticks(y_pos)
        ax.set_yticklabels(labels)
        ax.set_xlabel('Value', fontweight='bold')
        ax.set_title('Key Assumptions Summary', fontsize=14, fontweight='bold', pad=20)
        ax.grid(axis='x', alpha=0.3)
        
        plt.tight_layout()
        plt.savefig(output_path, dpi=150, bbox_inches='tight')
        plt.close()
        
        return output_path
    
    def create_price_projection_chart(
        self,
        assumptions: Dict,
        years: int = 20,
        output_path: Optional[str] = None
    ) -> str:
        """
        Create price growth projection chart (line chart).
        
        Parameters:
        -----------
        assumptions : Dict
            Dictionary of assumptions (should include price_growth_base)
        years : int
            Number of years to project
        output_path : str, optional
            Path to save chart. If None, saves to temp file.
            
        Returns:
        --------
        str
            Path to saved chart image
        """
        if output_path is None:
            output_path = os.path.join(self.temp_dir, 'price_projection.png')
        
        fig, ax = plt.subplots(figsize=(10, 6))
        
        # Generate price projection
        base_price = assumptions.get('base_price', 50.0)  # Default $50/ton
        growth_rate = assumptions.get('price_growth_base', 0.03)  # Default 3%
        
        years_list = list(range(1, years + 1))
        prices = [base_price * (1 + growth_rate) ** (y - 1) for y in years_list]
        
        # Plot line
        ax.plot(years_list, prices, color=self.colors['primary'], linewidth=2.5, marker='o', markersize=4)
        ax.fill_between(years_list, prices, alpha=0.2, color=self.colors['primary'])
        
        ax.set_xlabel('Year', fontweight='bold')
        ax.set_ylabel('Price ($/ton)', fontweight='bold')
        ax.set_title('Carbon Price Growth Projection', fontsize=14, fontweight='bold', pad=20)
        ax.grid(True, alpha=0.3)
        
        # Format y-axis as currency
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'${x:.0f}'))
        
        plt.tight_layout()
        plt.savefig(output_path, dpi=150, bbox_inches='tight')
        plt.close()
        
        return output_path
    
    def create_volume_projection_chart(
        self,
        assumptions: Dict,
        years: int = 20,
        output_path: Optional[str] = None
    ) -> str:
        """
        Create volume projection chart (line chart).
        
        Parameters:
        -----------
        assumptions : Dict
            Dictionary of assumptions
        years : int
            Number of years to project
        output_path : str, optional
            Path to save chart. If None, saves to temp file.
            
        Returns:
        --------
        str
            Path to saved chart image
        """
        if output_path is None:
            output_path = os.path.join(self.temp_dir, 'volume_projection.png')
        
        fig, ax = plt.subplots(figsize=(10, 6))
        
        # Generate volume projection (simplified - would use actual data if available)
        base_volume = assumptions.get('base_volume', 100000)  # Default 100k credits
        volume_multiplier = assumptions.get('volume_multiplier_base', 1.0)
        
        years_list = list(range(1, years + 1))
        volumes = [base_volume * volume_multiplier * (1.02 ** (y - 1)) for y in years_list]  # 2% growth
        
        # Plot line
        ax.plot(years_list, volumes, color=self.colors['secondary'], linewidth=2.5, marker='s', markersize=4)
        ax.fill_between(years_list, volumes, alpha=0.2, color=self.colors['secondary'])
        
        ax.set_xlabel('Year', fontweight='bold')
        ax.set_ylabel('Credit Volume', fontweight='bold')
        ax.set_title('Credit Volume Projection', fontsize=14, fontweight='bold', pad=20)
        ax.grid(True, alpha=0.3)
        
        # Format y-axis
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x/1000:.0f}K'))
        
        plt.tight_layout()
        plt.savefig(output_path, dpi=150, bbox_inches='tight')
        plt.close()
        
        return output_path
    
    def create_cash_flow_waterfall(
        self,
        valuation_schedule: pd.DataFrame,
        years: int = 20,
        output_path: Optional[str] = None
    ) -> str:
        """
        Create cash flow waterfall chart (stacked bar chart).
        
        Parameters:
        -----------
        valuation_schedule : pd.DataFrame
            Valuation schedule with cash flow data
        years : int
            Number of years to show
        output_path : str, optional
            Path to save chart. If None, saves to temp file.
            
        Returns:
        --------
        str
            Path to saved chart image
        """
        if output_path is None:
            output_path = os.path.join(self.temp_dir, 'cash_flow_waterfall.png')
        
        fig, ax = plt.subplots(figsize=(14, 6))
        
        # Extract data
        if 'cash_flow' in valuation_schedule.columns:
            cash_flows = valuation_schedule['cash_flow'].head(years).values
        else:
            # Fallback: use index if cash_flow column doesn't exist
            cash_flows = np.zeros(years)
        
        years_list = list(range(1, min(len(cash_flows), years) + 1))
        cash_flows = cash_flows[:len(years_list)]
        
        # Create stacked bars (positive and negative)
        positive_flows = np.where(cash_flows > 0, cash_flows, 0)
        negative_flows = np.where(cash_flows < 0, cash_flows, 0)
        
        # Plot bars
        x_pos = np.arange(len(years_list))
        width = 0.6
        
        bars_pos = ax.bar(x_pos, positive_flows, width, color=self.colors['secondary'], 
                         label='Positive Cash Flow', edgecolor='black', linewidth=0.5)
        bars_neg = ax.bar(x_pos, negative_flows, width, color=self.colors['negative'], 
                         label='Negative Cash Flow', edgecolor='black', linewidth=0.5)
        
        ax.set_xlabel('Year', fontweight='bold')
        ax.set_ylabel('Cash Flow ($)', fontweight='bold')
        ax.set_title('Annual Cash Flow Waterfall', fontsize=14, fontweight='bold', pad=20)
        ax.set_xticks(x_pos)
        ax.set_xticklabels(years_list)
        ax.legend(loc='upper left')
        ax.grid(axis='y', alpha=0.3)
        ax.axhline(y=0, color='black', linewidth=0.8)
        
        # Format y-axis as currency
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'${x/1e6:.1f}M'))
        
        plt.tight_layout()
        plt.savefig(output_path, dpi=150, bbox_inches='tight')
        plt.close()
        
        return output_path
    
    def create_cumulative_cash_flow(
        self,
        valuation_schedule: pd.DataFrame,
        years: int = 20,
        output_path: Optional[str] = None
    ) -> str:
        """
        Create cumulative cash flow chart (line chart).
        
        Parameters:
        -----------
        valuation_schedule : pd.DataFrame
            Valuation schedule with cash flow data
        years : int
            Number of years to show
        output_path : str, optional
            Path to save chart. If None, saves to temp file.
            
        Returns:
        --------
        str
            Path to saved chart image
        """
        if output_path is None:
            output_path = os.path.join(self.temp_dir, 'cumulative_cash_flow.png')
        
        fig, ax = plt.subplots(figsize=(10, 6))
        
        # Extract and calculate cumulative cash flow
        if 'cash_flow' in valuation_schedule.columns:
            cash_flows = valuation_schedule['cash_flow'].head(years).values
        else:
            cash_flows = np.zeros(years)
        
        years_list = list(range(1, min(len(cash_flows), years) + 1))
        cash_flows = cash_flows[:len(years_list)]
        cumulative = np.cumsum(cash_flows)
        
        # Plot line
        ax.plot(years_list, cumulative, color=self.colors['primary'], linewidth=2.5, marker='o', markersize=5)
        ax.fill_between(years_list, cumulative, 0, alpha=0.2, color=self.colors['primary'], 
                       where=(cumulative >= 0))
        ax.fill_between(years_list, cumulative, 0, alpha=0.2, color=self.colors['negative'], 
                       where=(cumulative < 0))
        
        ax.axhline(y=0, color='black', linewidth=1, linestyle='--')
        ax.set_xlabel('Year', fontweight='bold')
        ax.set_ylabel('Cumulative Cash Flow ($)', fontweight='bold')
        ax.set_title('Cumulative Cash Flow Over Time', fontsize=14, fontweight='bold', pad=20)
        ax.grid(True, alpha=0.3)
        
        # Format y-axis as currency
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'${x/1e6:.1f}M'))
        
        plt.tight_layout()
        plt.savefig(output_path, dpi=150, bbox_inches='tight')
        plt.close()
        
        return output_path
    
    def create_npv_trend(
        self,
        valuation_schedule: pd.DataFrame,
        years: int = 20,
        output_path: Optional[str] = None
    ) -> str:
        """
        Create NPV trend chart (line chart showing NPV progression).
        
        Parameters:
        -----------
        valuation_schedule : pd.DataFrame
            Valuation schedule with present value data
        years : int
            Number of years to show
        output_path : str, optional
            Path to save chart. If None, saves to temp file.
            
        Returns:
        --------
        str
            Path to saved chart image
        """
        if output_path is None:
            output_path = os.path.join(self.temp_dir, 'npv_trend.png')
        
        fig, ax = plt.subplots(figsize=(10, 6))
        
        # Extract and calculate cumulative NPV
        if 'present_value' in valuation_schedule.columns:
            pv_values = valuation_schedule['present_value'].head(years).values
        else:
            pv_values = np.zeros(years)
        
        years_list = list(range(1, min(len(pv_values), years) + 1))
        pv_values = pv_values[:len(years_list)]
        cumulative_npv = np.cumsum(pv_values)
        
        # Plot line
        ax.plot(years_list, cumulative_npv, color=self.colors['accent'], linewidth=2.5, marker='s', markersize=5)
        ax.fill_between(years_list, cumulative_npv, 0, alpha=0.2, color=self.colors['accent'], 
                       where=(cumulative_npv >= 0))
        ax.fill_between(years_list, cumulative_npv, 0, alpha=0.2, color=self.colors['negative'], 
                       where=(cumulative_npv < 0))
        
        ax.axhline(y=0, color='black', linewidth=1, linestyle='--')
        ax.set_xlabel('Year', fontweight='bold')
        ax.set_ylabel('Cumulative NPV ($)', fontweight='bold')
        ax.set_title('NPV Build-Up Over Time', fontsize=14, fontweight='bold', pad=20)
        ax.grid(True, alpha=0.3)
        
        # Format y-axis as currency
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'${x/1e6:.1f}M'))
        
        plt.tight_layout()
        plt.savefig(output_path, dpi=150, bbox_inches='tight')
        plt.close()
        
        return output_path
    
    def create_risk_breakdown(
        self,
        risk_score: Dict,
        output_path: Optional[str] = None
    ) -> str:
        """
        Create risk score breakdown chart (pie or bar chart).
        
        Parameters:
        -----------
        risk_score : Dict
            Dictionary with risk score components
        output_path : str, optional
            Path to save chart. If None, saves to temp file.
            
        Returns:
        --------
        str
            Path to saved chart image
        """
        if output_path is None:
            output_path = os.path.join(self.temp_dir, 'risk_breakdown.png')
        
        fig, ax = plt.subplots(figsize=(8, 6))
        
        # Extract risk components
        labels = []
        values = []
        colors_list = []
        
        if 'financial_risk' in risk_score:
            labels.append('Financial Risk')
            values.append(risk_score['financial_risk'])
            colors_list.append(self.colors['negative'])
        
        if 'volume_risk' in risk_score:
            labels.append('Volume Risk')
            values.append(risk_score['volume_risk'])
            colors_list.append(self.colors['accent'])
        
        if 'price_risk' in risk_score:
            labels.append('Price Risk')
            values.append(risk_score['price_risk'])
            colors_list.append(self.colors['neutral'])
        
        if not labels:
            # Fallback: create empty chart
            ax.text(0.5, 0.5, 'No risk data available', ha='center', va='center', 
                   fontsize=12, transform=ax.transAxes)
            ax.set_title('Risk Score Breakdown', fontsize=14, fontweight='bold', pad=20)
            plt.tight_layout()
            plt.savefig(output_path, dpi=150, bbox_inches='tight')
            plt.close()
            return output_path
        
        # Create horizontal bar chart
        y_pos = np.arange(len(labels))
        bars = ax.barh(y_pos, values, color=colors_list, edgecolor='black', linewidth=1.5)
        
        # Add value labels
        for bar, val in zip(bars, values):
            ax.text(bar.get_width() * 0.5, bar.get_y() + bar.get_height()/2,
                   f'{val:.0f}', ha='center', va='center', fontweight='bold', color='white')
        
        ax.set_yticks(y_pos)
        ax.set_yticklabels(labels)
        ax.set_xlabel('Risk Score', fontweight='bold')
        ax.set_title('Risk Score Breakdown', fontsize=14, fontweight='bold', pad=20)
        ax.set_xlim(0, 100)
        ax.grid(axis='x', alpha=0.3)
        
        plt.tight_layout()
        plt.savefig(output_path, dpi=150, bbox_inches='tight')
        plt.close()
        
        return output_path
    
    def create_return_summary(
        self,
        target_irr: float,
        actual_irr: float,
        output_path: Optional[str] = None
    ) -> str:
        """
        Create investment return summary chart (bar chart comparing Target vs Actual IRR).
        
        Parameters:
        -----------
        target_irr : float
            Target IRR
        actual_irr : float
            Actual IRR achieved
        output_path : str, optional
            Path to save chart. If None, saves to temp file.
            
        Returns:
        --------
        str
            Path to saved chart image
        """
        if output_path is None:
            output_path = os.path.join(self.temp_dir, 'return_summary.png')
        
        fig, ax = plt.subplots(figsize=(8, 6))
        
        # Prepare data
        labels = ['Target IRR', 'Actual IRR']
        values = [target_irr * 100, actual_irr * 100]  # Convert to percentage
        colors_list = [self.colors['neutral'], self.colors['secondary'] if actual_irr >= target_irr else self.colors['negative']]
        
        # Create bar chart
        bars = ax.bar(labels, values, color=colors_list, edgecolor='black', linewidth=1.5, width=0.6)
        
        # Add value labels on bars
        for bar, val in zip(bars, values):
            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
                   f'{val:.1f}%', ha='center', va='bottom', fontweight='bold', fontsize=12)
        
        ax.set_ylabel('IRR (%)', fontweight='bold')
        ax.set_title('Investment Return Summary', fontsize=14, fontweight='bold', pad=20)
        ax.set_ylim(0, max(values) * 1.2)
        ax.grid(axis='y', alpha=0.3)
        
        # Add target line
        ax.axhline(y=target_irr * 100, color=self.colors['accent'], linewidth=1.5, 
                  linestyle='--', label='Target')
        
        plt.tight_layout()
        plt.savefig(output_path, dpi=150, bbox_inches='tight')
        plt.close()
        
        return output_path
    
    def embed_chart_in_excel(
        self,
        chart_path: str,
        worksheet,
        cell_ref: str,
        width: int = 400,
        height: int = 300
    ) -> None:
        """
        Embed chart image into Excel worksheet using openpyxl.
        
        Parameters:
        -----------
        chart_path : str
            Path to chart image file
        worksheet : openpyxl.worksheet.worksheet.Worksheet
            Excel worksheet to add chart to
        cell_ref : str
            Cell reference where to place chart (e.g., 'E2')
        width : int
            Chart width in pixels (default: 400)
        height : int
            Chart height in pixels (default: 300)
        """
        try:
            from openpyxl.drawing.image import Image
            
            if not os.path.exists(chart_path):
                print(f"Warning: Chart file not found: {chart_path}")
                return
            
            img = Image(chart_path)
            img.width = width
            img.height = height
            
            worksheet.add_image(img, cell_ref)
            
        except ImportError:
            print("Warning: openpyxl.drawing.image not available. Chart cannot be embedded.")
        except Exception as e:
            print(f"Warning: Could not embed chart: {e}")

