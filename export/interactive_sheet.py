"""
Interactive Analysis Sheet Module: Creates Excel sheet for interactive parameter adjustment.
"""

import xlsxwriter
from typing import Dict, Optional
import pandas as pd
import numpy as np


class InteractiveSheetCreator:
    """
    Creates an interactive Excel sheet for adjusting GBM/Monte Carlo parameters.
    """
    
    def __init__(self, workbook: xlsxwriter.Workbook):
        """
        Initialize interactive sheet creator.
        
        Parameters:
        -----------
        workbook : xlsxwriter.Workbook
            Excel workbook
        """
        self.workbook = workbook
    
    def create_interactive_analysis_sheet(
        self,
        base_assumptions: Dict,
        monte_carlo_results: Optional[Dict] = None,
        sheet_name: str = "Interactive Analysis"
    ) -> xlsxwriter.Workbook.worksheet_class:
        """
        Create interactive sheet for parameter adjustment.
        
        Parameters:
        -----------
        base_assumptions : Dict
            Base assumptions dictionary
        monte_carlo_results : Dict, optional
            Monte Carlo results
        sheet_name : str
            Name of the sheet
            
        Returns:
        --------
        xlsxwriter worksheet
            Created worksheet
        """
        worksheet = self.workbook.add_worksheet(sheet_name)
        
        # Define formats
        formats = {
            'title': self.workbook.add_format({
                'bold': True, 'font_size': 16, 'bg_color': '#366092',
                'font_color': 'white', 'align': 'center', 'valign': 'vcenter'
            }),
            'subtitle': self.workbook.add_format({
                'bold': True, 'font_size': 12, 'bg_color': '#E7E6E6',
                'align': 'left', 'valign': 'vcenter'
            }),
            'input_label': self.workbook.add_format({
                'bold': True, 'bg_color': '#D9E1F2', 'border': 1,
                'align': 'right', 'valign': 'vcenter'
            }),
            'input_cell': self.workbook.add_format({
                'bg_color': '#FFF2CC', 'border': 1, 'num_format': 'General',
                'valign': 'vcenter'
            }),
            'formula_cell': self.workbook.add_format({
                'bg_color': '#E2EFDA', 'border': 1, 'valign': 'vcenter'
            }),
            'percent': self.workbook.add_format({
                'num_format': '0.00%', 'border': 1, 'valign': 'vcenter'
            }),
            'currency': self.workbook.add_format({
                'num_format': '$#,##0', 'border': 1, 'valign': 'vcenter'
            }),
            'number': self.workbook.add_format({
                'num_format': '#,##0', 'border': 1, 'valign': 'vcenter'
            }),
            'note': self.workbook.add_format({
                'italic': True, 'font_color': '#666666', 'font_size': 9
            }),
            'warning': self.workbook.add_format({
                'bold': True, 'bg_color': '#FFEB9C', 'border': 1
            })
        }
        
        row = 0
        
        # Title
        worksheet.merge_range(row, 0, row, 4, 
                            'Interactive Monte Carlo & GBM Analysis', formats['title'])
        worksheet.set_row(row, 35)
        row += 2
        
        # Instructions
        instructions = [
            "INSTRUCTIONS:",
            "1. Adjust parameters in the 'Input Parameters' section below",
            "2. Results will update automatically based on your inputs",
            "3. For full Monte Carlo simulation, run: python3 examples/run_interactive_analysis.py",
            "4. This sheet provides formula-based approximations for quick analysis"
        ]
        
        for i, instruction in enumerate(instructions):
            if i == 0:
                worksheet.write(row, 0, instruction, formats['subtitle'])
            else:
                worksheet.write(row, 0, instruction, formats['note'])
            row += 1
        
        row += 1
        
        # Section 1: Base Financial Assumptions
        worksheet.write(row, 0, 'Base Financial Assumptions', formats['subtitle'])
        worksheet.merge_range(row, 0, row, 4, '', formats['subtitle'])
        row += 1
        
        base_params = [
            ('WACC', 'wacc', 'percent', 0.08),
            ('Rubicon Investment Total', 'rubicon_investment_total', 'currency', 20000000),
            ('Investment Tenor (Years)', 'investment_tenor', 'number', 5),
            ('Initial Streaming Percentage', 'streaming_percentage_initial', 'percent', 0.48)
        ]
        
        for label, key, fmt_type, default in base_params:
            worksheet.write(row, 0, label, formats['input_label'])
            value = base_assumptions.get(key, default)
            if fmt_type == 'percent':
                worksheet.write(row, 1, value, formats['input_cell'])
                worksheet.write(row, 1, value, formats['percent'])
            elif fmt_type == 'currency':
                worksheet.write(row, 1, value, formats['input_cell'])
                worksheet.write(row, 1, value, formats['currency'])
            else:
                worksheet.write(row, 1, value, formats['input_cell'])
                worksheet.write(row, 1, value, formats['number'])
            
            # Create named range for formula references
            cell_ref = xlsxwriter.utility.xl_rowcol_to_cell(row, 1)
            self.workbook.define_name(
                f'Input_{key.upper()}',
                f"'{worksheet.name}'!{cell_ref}"
            )
            row += 1
        
        row += 1
        
        # Section 2: GBM Parameters
        worksheet.write(row, 0, 'GBM (Geometric Brownian Motion) Parameters', formats['subtitle'])
        worksheet.merge_range(row, 0, row, 4, '', formats['subtitle'])
        row += 1
        
        gbm_params = [
            ('Use GBM Method', 'use_gbm', 'text', False),
            ('GBM Drift (μ) - Expected Return', 'gbm_drift', 'percent', 0.03),
            ('GBM Volatility (σ) - Price Volatility', 'gbm_volatility', 'percent', 0.15)
        ]
        
        for label, key, fmt_type, default in gbm_params:
            worksheet.write(row, 0, label, formats['input_label'])
            value = base_assumptions.get(key, default)
            
            if fmt_type == 'text':
                worksheet.write(row, 1, 'Yes' if value else 'No', formats['input_cell'])
            elif fmt_type == 'percent':
                worksheet.write(row, 1, value, formats['input_cell'])
                worksheet.write(row, 1, value, formats['percent'])
            else:
                worksheet.write(row, 1, value, formats['input_cell'])
            
            cell_ref = xlsxwriter.utility.xl_rowcol_to_cell(row, 1)
            self.workbook.define_name(
                f'Input_{key.upper()}',
                f"'{worksheet.name}'!{cell_ref}"
            )
            row += 1
        
        row += 1
        
        # Section 3: Monte Carlo Parameters
        worksheet.write(row, 0, 'Monte Carlo Simulation Parameters', formats['subtitle'])
        worksheet.merge_range(row, 0, row, 4, '', formats['subtitle'])
        row += 1
        
        mc_params = [
            ('Number of Simulations', 'simulations', 'number', 5000),
            ('Price Growth Base (Mean)', 'price_growth_base', 'percent', 0.03),
            ('Price Growth Std Dev', 'price_growth_std_dev', 'percent', 0.02),
            ('Volume Multiplier Base (Mean)', 'volume_multiplier_base', 'number', 1.0),
            ('Volume Std Dev', 'volume_std_dev', 'percent', 0.15)
        ]
        
        for label, key, fmt_type, default in mc_params:
            worksheet.write(row, 0, label, formats['input_label'])
            value = base_assumptions.get(key, default)
            
            if fmt_type == 'percent':
                worksheet.write(row, 1, value, formats['input_cell'])
                worksheet.write(row, 1, value, formats['percent'])
            elif fmt_type == 'number':
                if key == 'simulations':
                    worksheet.write(row, 1, int(value), formats['input_cell'])
                    worksheet.write(row, 1, int(value), formats['number'])
                else:
                    worksheet.write(row, 1, value, formats['input_cell'])
                    worksheet.write(row, 1, value, formats['number'])
            else:
                worksheet.write(row, 1, value, formats['input_cell'])
            
            cell_ref = xlsxwriter.utility.xl_rowcol_to_cell(row, 1)
            self.workbook.define_name(
                f'Input_{key.upper()}',
                f"'{worksheet.name}'!{cell_ref}"
            )
            row += 1
        
        row += 2
        
        # Section 4: Quick Results (Formula-based approximations)
        worksheet.write(row, 0, 'Quick Results (Formula-Based)', formats['subtitle'])
        worksheet.merge_range(row, 0, row, 4, '', formats['subtitle'])
        row += 1
        
        worksheet.write(row, 0, 'Note:', formats['note'])
        worksheet.write(row, 1, 
                        'These are approximate results. For full Monte Carlo simulation, run Python script.',
                        formats['note'])
        row += 1
        
        # Show current results if available
        if monte_carlo_results:
            results_data = [
                ('Mean IRR', 'mc_mean_irr', 'percent'),
                ('P10 IRR (10th Percentile)', 'mc_p10_irr', 'percent'),
                ('P90 IRR (90th Percentile)', 'mc_p90_irr', 'percent'),
                ('Mean NPV', 'mc_mean_npv', 'currency'),
                ('P10 NPV', 'mc_p10_npv', 'currency'),
                ('P90 NPV', 'mc_p90_npv', 'currency')
            ]
            
            worksheet.write(row, 0, 'Current Results:', formats['input_label'])
            row += 1
            
            for label, key, fmt_type in results_data:
                worksheet.write(row, 0, label, formats['input_label'])
                value = monte_carlo_results.get(key, 0)
                
                if pd.isna(value) or not np.isfinite(value):
                    worksheet.write(row, 1, 'N/A', formats['formula_cell'])
                elif fmt_type == 'percent':
                    worksheet.write(row, 1, value, formats['formula_cell'])
                    worksheet.write(row, 1, value, formats['percent'])
                else:
                    worksheet.write(row, 1, value, formats['formula_cell'])
                    worksheet.write(row, 1, value, formats['currency'])
                row += 1
        
        row += 2
        
        # Section 5: Instructions for Python Script
        worksheet.write(row, 0, 'Run Full Analysis', formats['subtitle'])
        worksheet.merge_range(row, 0, row, 4, '', formats['subtitle'])
        row += 1
        
        instructions_text = [
            "To run full Monte Carlo analysis with your adjusted parameters:",
            "",
            "1. Save this Excel file",
            "2. Run: python3 examples/run_interactive_analysis.py",
            "3. The script will read parameters from this sheet",
            "4. Results will be updated in the 'Summary & Results' sheet",
            "",
            "Or use the configuration tool:",
            "  python3 analysis_config.py"
        ]
        
        for text in instructions_text:
            if text:
                worksheet.write(row, 0, text, formats['note'])
            row += 1
        
        # Set column widths
        worksheet.set_column(0, 0, 45)
        worksheet.set_column(1, 1, 25)
        worksheet.set_column(2, 4, 15)
        
        return worksheet

