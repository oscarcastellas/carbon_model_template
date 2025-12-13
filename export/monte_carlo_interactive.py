"""
Interactive Monte Carlo Simulation Sheet Creator

Creates an Excel sheet with input cells and results area for running
Monte Carlo simulation from within Excel.
"""

import xlsxwriter
from typing import Dict, Optional
import pandas as pd
import numpy as np


class InteractiveMonteCarloSheet:
    """
    Creates an interactive Excel sheet for Monte Carlo simulation.
    
    Users can:
    - Set simulation parameters (number of sims, GBM settings, etc.)
    - Set volatility assumptions
    - Run Monte Carlo simulation
    - See statistical results
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
    
    def create_interactive_sheet(
        self,
        base_assumptions: Dict,
        sheet_name: str = "Monte Carlo Results"
    ):
        """
        Create interactive Monte Carlo simulation sheet.
        
        Parameters:
        -----------
        base_assumptions : Dict
            Base assumptions dictionary
        sheet_name : str
            Name of the sheet
        """
        worksheet = self.workbook.add_worksheet(sheet_name)
        
        # Define formats
        formats = {
            'title': self.workbook.add_format({
                'bold': True, 'font_size': 16, 'bg_color': '#366092',
                'font_color': 'white', 'align': 'center', 'valign': 'vcenter'
            }),
            'section_header': self.workbook.add_format({
                'bold': True, 'font_size': 12, 'bg_color': '#E7E6E6',
                'align': 'left', 'valign': 'vcenter'
            }),
            'input_label': self.workbook.add_format({
                'bold': True, 'bg_color': '#D9E1F2', 'border': 1,
                'align': 'right', 'valign': 'vcenter'
            }),
            'input_cell': self.workbook.add_format({
                'bg_color': '#FFF2CC', 'border': 1,
                'valign': 'vcenter'
            }),
            'result_label': self.workbook.add_format({
                'bold': True, 'bg_color': '#E2EFDA', 'border': 1,
                'align': 'right', 'valign': 'vcenter'
            }),
            'result_cell': self.workbook.add_format({
                'bg_color': '#E2EFDA', 'border': 1,
                'valign': 'vcenter'
            }),
            'percent': self.workbook.add_format({
                'num_format': '0.00%', 'border': 1, 'valign': 'vcenter'
            }),
            'currency': self.workbook.add_format({
                'num_format': '$#,##0', 'border': 1, 'valign': 'vcenter'
            }),
            'currency_2dec': self.workbook.add_format({
                'num_format': '$#,##0.00', 'border': 1, 'valign': 'vcenter'
            }),
            'number': self.workbook.add_format({
                'num_format': '#,##0', 'border': 1, 'valign': 'vcenter'
            }),
            'number_2dec': self.workbook.add_format({
                'num_format': '#,##0.00', 'border': 1, 'valign': 'vcenter'
            }),
            'note': self.workbook.add_format({
                'italic': True, 'font_color': '#666666', 'font_size': 9
            }),
            'button': self.workbook.add_format({
                'bold': True, 'bg_color': '#70AD47', 'font_color': 'white',
                'align': 'center', 'valign': 'vcenter', 'border': 1
            })
        }
        
        row = 0
        
        # Title
        worksheet.write(row, 0, 'Interactive Monte Carlo Simulation', formats['title'])
        worksheet.merge_range(row, 0, row, 4, 
                            'Interactive Monte Carlo Simulation', formats['title'])
        worksheet.set_row(row, 35)
        row += 2
        
        # Instructions
        instructions = [
            "INSTRUCTIONS:",
            "1. Set your simulation parameters in the sections below",
            "2. Choose GBM or Growth Rate method for price volatility",
            "3. Run: python3 scripts/run_monte_carlo_from_excel.py [this_file.xlsx]",
            "4. Results will appear in the 'Results' section"
        ]
        
        for i, instruction in enumerate(instructions):
            if i == 0:
                worksheet.write(row, 0, instruction, formats['section_header'])
            else:
                worksheet.write(row, 0, instruction, formats['note'])
            row += 1
        
        row += 1
        
        # Section 1: Basic Parameters
        worksheet.write(row, 0, 'Basic Parameters', formats['section_header'])
        row += 1
        
        inputs_basic = [
            ('Number of Simulations', 'B8', 'number', 'More sims = more accurate but slower'),
            ('Streaming Percentage', 'B9', 'percent', 'Percentage of credits streamed'),
            ('Random Seed (optional)', 'B10', 'number', 'Leave blank for random, or set for reproducibility')
        ]
        
        for label, cell_ref, fmt_type, note in inputs_basic:
            worksheet.write(row, 0, label, formats['input_label'])
            # Empty cells - user fills
            if fmt_type == 'percent':
                worksheet.write(row, 1, '', formats['percent'])
            elif fmt_type == 'number':
                worksheet.write(row, 1, '', formats['number'])
            worksheet.write(row, 2, note, formats['note'])
            row += 1
        
        row += 1
        
        # Section 2: Price Volatility Method
        worksheet.write(row, 0, 'Price Volatility Method', formats['section_header'])
        row += 1
        
        worksheet.write(row, 0, 'Use GBM (Geometric Brownian Motion)?', formats['input_label'])
        worksheet.write(row, 1, '', formats['input_cell'])  # Empty
        worksheet.write(row, 2, 'Yes = GBM method, No = Growth Rate method', formats['note'])
        row += 1
        
        # GBM Parameters
        worksheet.write(row, 0, 'GBM Parameters (if using GBM)', formats['input_label'])
        row += 1
        
        gbm_inputs = [
            ('GBM Drift (μ) - Expected Return', 'B14', 'percent', 'Expected annual price return (e.g., 3%)'),
            ('GBM Volatility (σ)', 'B15', 'percent', 'Annual price volatility (e.g., 15%)')
        ]
        
        for label, cell_ref, fmt_type, note in gbm_inputs:
            worksheet.write(row, 0, label, formats['input_label'])
            worksheet.write(row, 1, '', formats['percent'])  # Empty
            worksheet.write(row, 2, note, formats['note'])
            row += 1
        
        row += 1
        
        # Alternative: Growth Rate Parameters
        worksheet.write(row, 0, 'Growth Rate Parameters (if NOT using GBM)', formats['input_label'])
        row += 1
        
        growth_inputs = [
            ('Price Growth Base (Mean)', 'B17', 'percent', 'Mean annual price growth (e.g., 3%)'),
            ('Price Growth Std Dev', 'B18', 'percent', 'Std dev of price growth (e.g., 2%)'),
            ('Use Percentage Variation?', 'B19', 'text', 'Yes = % multipliers, No = growth rate deviations')
        ]
        
        for label, cell_ref, fmt_type, note in growth_inputs:
            worksheet.write(row, 0, label, formats['input_label'])
            if fmt_type == 'percent':
                worksheet.write(row, 1, '', formats['percent'])  # Empty
            else:
                worksheet.write(row, 1, '', formats['input_cell'])  # Empty
            worksheet.write(row, 2, note, formats['note'])
            row += 1
        
        row += 1
        
        # Section 3: Volume Parameters
        worksheet.write(row, 0, 'Volume Parameters', formats['section_header'])
        row += 1
        
        volume_inputs = [
            ('Volume Multiplier Base (Mean)', 'B21', 'number_2dec', 1.0, 'Mean volume multiplier (typically 1.0)'),
            ('Volume Std Dev', 'B22', 'percent', 0.15, 'Std dev of volume multiplier (e.g., 15%)')
        ]
        
        for label, cell_ref, fmt_type, default, note in volume_inputs:
            worksheet.write(row, 0, label, formats['input_label'])
            worksheet.write(row, 1, default, formats['input_cell'])
            if fmt_type == 'percent':
                worksheet.write(row, 1, default, formats['percent'])
            elif fmt_type == 'number_2dec':
                worksheet.write(row, 1, default, formats['number_2dec'])
            worksheet.write(row, 2, note, formats['note'])
            row += 1
        
        row += 1
        
        # Section 4: Run Analysis
        worksheet.write(row, 0, 'Run Analysis', formats['section_header'])
        row += 1
        worksheet.write(row, 0, 'Run Monte Carlo Simulation', formats['button'])
        worksheet.write(row, 2, 'Run: python3 scripts/run_monte_carlo_from_excel.py [this_file.xlsx]', formats['note'])
        worksheet.write(row, 3, 'Note: This may take several minutes for 5000+ simulations', formats['note'])
        row += 2
        
        # Section 5: Results
        worksheet.write(row, 0, 'Results', formats['section_header'])
        row += 1
        
        # IRR Results
        worksheet.write(row, 0, 'IRR Statistics', formats['result_label'])
        row += 1
        
        irr_results = [
            ('Mean IRR', 'B27'),
            ('P10 IRR (10th Percentile)', 'B28'),
            ('P50 IRR (Median)', 'B29'),
            ('P90 IRR (90th Percentile)', 'B30'),
            ('Std Dev IRR', 'B31'),
            ('Min IRR', 'B32'),
            ('Max IRR', 'B33')
        ]
        
        for label, cell_ref in irr_results:
            worksheet.write(row, 0, label, formats['result_label'])
            worksheet.write(row, 1, '', formats['result_cell'])
            worksheet.write(row, 1, '', formats['percent'])
            row += 1
        
        row += 1
        
        # NPV Results
        worksheet.write(row, 0, 'NPV Statistics', formats['result_label'])
        row += 1
        
        npv_results = [
            ('Mean NPV', 'B35'),
            ('P10 NPV (10th Percentile)', 'B36'),
            ('P50 NPV (Median)', 'B37'),
            ('P90 NPV (90th Percentile)', 'B38'),
            ('Std Dev NPV', 'B39'),
            ('Min NPV', 'B40'),
            ('Max NPV', 'B41')
        ]
        
        for label, cell_ref in npv_results:
            worksheet.write(row, 0, label, formats['result_label'])
            worksheet.write(row, 1, '', formats['result_cell'])
            worksheet.write(row, 1, '', formats['currency_2dec'])
            row += 1
        
        row += 1
        
        # Probabilities
        worksheet.write(row, 0, 'Probabilities', formats['result_label'])
        row += 1
        
        prob_results = [
            ('Prob(IRR > 20%)', 'B43'),
            ('Prob(IRR > 15%)', 'B44'),
            ('Prob(NPV > $0)', 'B45'),
            ('Prob(NPV > $10M)', 'B46')
        ]
        
        for label, cell_ref in prob_results:
            worksheet.write(row, 0, label, formats['result_label'])
            worksheet.write(row, 1, '', formats['result_cell'])
            worksheet.write(row, 1, '', formats['percent'])
            row += 1
        
        row += 1
        
        # Status
        worksheet.write(row, 0, 'Status', formats['input_label'])
        worksheet.write(row, 1, 'Ready', formats['result_cell'])
        worksheet.write(row, 2, 'Status of last calculation', formats['note'])
        
        # Set column widths
        # Reserve space for charts (columns E-H)
        row += 2
        worksheet.write(row, 4, 'Charts will appear here after running analysis', formats['note'])
        worksheet.merge_range(row, 4, row + 20, 7, '', formats['note'])
        
        worksheet.set_column(0, 0, 40)
        worksheet.set_column(1, 1, 20)
        worksheet.set_column(2, 3, 35)
        worksheet.set_column(4, 7, 20)  # Reserve space for charts
        
        return worksheet

