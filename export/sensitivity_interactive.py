"""
Interactive Sensitivity Analysis Sheet Creator

Creates an Excel sheet with input cells and results area for running
sensitivity analysis from within Excel.
"""

import xlsxwriter
from typing import Dict, Optional, List
import pandas as pd
import numpy as np


class InteractiveSensitivitySheet:
    """
    Creates an interactive Excel sheet for sensitivity analysis.
    
    Users can:
    - Set input ranges (credit volume, price multipliers)
    - Set streaming percentage
    - Run sensitivity analysis
    - See results in a 2D table
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
        sheet_name: str = "Sensitivity Analysis"
    ):
        """
        Create interactive sensitivity analysis sheet.
        
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
            'table_header': self.workbook.add_format({
                'bold': True, 'bg_color': '#4472C4', 'font_color': 'white',
                'border': 1, 'align': 'center', 'valign': 'vcenter'
            }),
            'table_cell': self.workbook.add_format({
                'border': 1, 'align': 'center', 'valign': 'vcenter'
            }),
            'percent': self.workbook.add_format({
                'num_format': '0.00%', 'border': 1, 'valign': 'vcenter'
            }),
            'number': self.workbook.add_format({
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
        worksheet.write(row, 0, 'Interactive Sensitivity Analysis', formats['title'])
        worksheet.merge_range(row, 0, row, 4, 
                            'Interactive Sensitivity Analysis', formats['title'])
        worksheet.set_row(row, 35)
        row += 2
        
        # Instructions
        instructions = [
            "INSTRUCTIONS:",
            "1. Set your input ranges in the 'Input Parameters' section below",
            "2. Set the streaming percentage to use",
            "3. Run: python3 scripts/run_sensitivity_from_excel.py [this_file.xlsx]",
            "4. Results will appear in the 'Sensitivity Table' section"
        ]
        
        for i, instruction in enumerate(instructions):
            if i == 0:
                worksheet.write(row, 0, instruction, formats['section_header'])
            else:
                worksheet.write(row, 0, instruction, formats['note'])
            row += 1
        
        row += 1
        
        # Section 1: Input Parameters
        worksheet.write(row, 0, 'Input Parameters', formats['section_header'])
        row += 1
        
        # Credit Volume Range (EMPTY - user fills)
        worksheet.write(row, 0, 'Credit Volume Range', formats['input_label'])
        worksheet.write(row, 1, 'Min Multiplier', formats['input_label'])
        worksheet.write(row, 2, '', formats['input_cell'])  # Empty
        worksheet.write(row, 2, '', formats['number'])
        worksheet.write(row, 3, 'Max Multiplier', formats['input_label'])
        worksheet.write(row, 4, '', formats['input_cell'])  # Empty
        worksheet.write(row, 4, '', formats['number'])
        worksheet.write(row, 5, 'Step', formats['input_label'])
        worksheet.write(row, 6, '', formats['input_cell'])  # Empty
        worksheet.write(row, 6, '', formats['number'])
        row += 1
        
        # Price Multiplier Range (EMPTY - user fills)
        worksheet.write(row, 0, 'Price Multiplier Range', formats['input_label'])
        worksheet.write(row, 1, 'Min Multiplier', formats['input_label'])
        worksheet.write(row, 2, '', formats['input_cell'])  # Empty
        worksheet.write(row, 2, '', formats['number'])
        worksheet.write(row, 3, 'Max Multiplier', formats['input_label'])
        worksheet.write(row, 4, '', formats['input_cell'])  # Empty
        worksheet.write(row, 4, '', formats['number'])
        worksheet.write(row, 5, 'Step', formats['input_label'])
        worksheet.write(row, 6, '', formats['input_cell'])  # Empty
        worksheet.write(row, 6, '', formats['number'])
        row += 1
        
        # Streaming Percentage (EMPTY - user fills)
        worksheet.write(row, 0, 'Streaming Percentage', formats['input_label'])
        worksheet.write(row, 1, '', formats['input_cell'])  # Empty
        worksheet.write(row, 1, '', formats['percent'])
        worksheet.write(row, 2, 'Percentage of credits streamed', formats['note'])
        row += 2
        
        # Section 2: Run Analysis
        worksheet.write(row, 0, 'Run Analysis', formats['section_header'])
        row += 1
        worksheet.write(row, 0, 'Run Sensitivity Analysis', formats['button'])
        worksheet.write(row, 2, 'Run: python3 scripts/run_sensitivity_from_excel.py [this_file.xlsx]', formats['note'])
        row += 2
        
        # Section 3: Sensitivity Table
        worksheet.write(row, 0, 'Sensitivity Table (IRR by Volume × Price)', formats['section_header'])
        row += 1
        
        # Table will be written here by the runner script
        # Placeholder: header row
        worksheet.write(row, 0, 'Credit Volume →', formats['table_header'])
        worksheet.write(row, 1, 'Price Multiplier ↓', formats['table_header'])
        worksheet.write(row, 2, '(Table will be populated here)', formats['note'])
        row += 2
        
        # Section 4: Summary Statistics
        worksheet.write(row, 0, 'Summary Statistics', formats['section_header'])
        row += 1
        
        summary_labels = [
            ('Minimum IRR', 'B30'),
            ('Maximum IRR', 'B31'),
            ('IRR Range', 'B32'),
            ('Base Case IRR (1.0x × 1.0x)', 'B33')
        ]
        
        for label, cell_ref in summary_labels:
            worksheet.write(row, 0, label, formats['result_label'])
            worksheet.write(row, 1, '', formats['result_cell'])
            worksheet.write(row, 1, '', formats['percent'])
            row += 1
        
        row += 1
        
        # Section 5: Status
        worksheet.write(row, 0, 'Status', formats['input_label'])
        worksheet.write(row, 1, 'Ready', formats['result_cell'])
        worksheet.write(row, 2, 'Status of last calculation', formats['note'])
        
        # Reserve space for charts (columns I-L)
        row += 2
        worksheet.write(row, 8, 'Charts will appear here after running analysis', formats['note'])
        worksheet.merge_range(row, 8, row + 15, 11, '', formats['note'])
        
        # Set column widths
        worksheet.set_column(0, 0, 30)
        worksheet.set_column(1, 1, 20)
        worksheet.set_column(2, 6, 15)
        worksheet.set_column(8, 11, 20)  # Reserve space for charts
        
        # Create named ranges for easy reference
        named_ranges = {
            'Input_CreditMin': 'C8',
            'Input_CreditMax': 'E8',
            'Input_CreditStep': 'G8',
            'Input_PriceMin': 'C9',
            'Input_PriceMax': 'E9',
            'Input_PriceStep': 'G9',
            'Input_StreamingPct': 'B10',
            'Status': 'B35'
        }
        
        for name, cell in named_ranges.items():
            self.workbook.define_name(name, f"'{sheet_name}'!{cell}")
        
        return worksheet

