"""
Interactive Breakeven Calculator Sheet Creator

Creates an Excel sheet with input cells and results area for running
breakeven analysis from within Excel.
"""

import xlsxwriter
from typing import Dict, Optional
import pandas as pd
import numpy as np


class InteractiveBreakevenSheet:
    """
    Creates an interactive Excel sheet for breakeven analysis.
    
    Users can:
    - Choose which breakeven to calculate (price, volume, streaming, or all)
    - Set target NPV
    - Set streaming percentage (for price/volume calculations)
    - Run breakeven analysis
    - See results
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
        sheet_name: str = "Breakeven Analysis"
    ):
        """
        Create interactive breakeven calculator sheet.
        
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
        worksheet.write(row, 0, 'Interactive Breakeven Calculator', formats['title'])
        worksheet.merge_range(row, 0, row, 4, 
                            'Interactive Breakeven Calculator', formats['title'])
        worksheet.set_row(row, 35)
        row += 2
        
        # Instructions
        instructions = [
            "INSTRUCTIONS:",
            "1. Choose which breakeven to calculate (price, volume, streaming, or all)",
            "2. Set target NPV (0 for true breakeven, or another value)",
            "3. Set streaming percentage (needed for price/volume calculations)",
            "4. Run: python3 scripts/run_breakeven_from_excel.py [this_file.xlsx]",
            "5. Results will appear in the 'Results' section"
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
        
        inputs = [
            ('Which Breakeven to Calculate', 'B8', 'text', 'Options: "all", "price", "volume", or "streaming"'),
            ('Target NPV', 'B9', 'currency_2dec', 'Target NPV (0 for true breakeven, or another value)'),
            ('Streaming Percentage', 'B10', 'percent', 'Needed for price/volume calculations')
        ]
        
        for label, cell_ref, fmt_type, note in inputs:
            worksheet.write(row, 0, label, formats['input_label'])
            # Empty cells - user fills
            if fmt_type == 'percent':
                worksheet.write(row, 1, '', formats['percent'])
            elif fmt_type == 'currency_2dec':
                worksheet.write(row, 1, '', formats['currency_2dec'])
            else:
                worksheet.write(row, 1, '', formats['input_cell'])
            worksheet.write(row, 2, note, formats['note'])
            row += 1
        
        row += 1
        
        # Section 2: Run Analysis
        worksheet.write(row, 0, 'Run Analysis', formats['section_header'])
        row += 1
        worksheet.write(row, 0, 'Run Breakeven Calculator', formats['button'])
        worksheet.write(row, 2, 'Run: python3 scripts/run_breakeven_from_excel.py [this_file.xlsx]', formats['note'])
        row += 2
        
        # Section 3: Results
        worksheet.write(row, 0, 'Results', formats['section_header'])
        row += 1
        
        # Breakeven Price Results
        worksheet.write(row, 0, 'Breakeven Price', formats['result_label'])
        row += 1
        
        price_results = [
            ('Breakeven Price per Ton', 'B15'),
            ('Base Price (Average)', 'B16'),
            ('Price Multiplier', 'B17'),
            ('Target NPV', 'B18')
        ]
        
        for label, cell_ref in price_results:
            worksheet.write(row, 0, label, formats['result_label'])
            worksheet.write(row, 1, '', formats['result_cell'])
            if 'Price' in label or 'NPV' in label:
                worksheet.write(row, 1, '', formats['currency_2dec'])
            elif 'Multiplier' in label:
                worksheet.write(row, 1, '', formats['number_2dec'])
            row += 1
        
        row += 1
        
        # Breakeven Volume Results
        worksheet.write(row, 0, 'Breakeven Volume', formats['result_label'])
        row += 1
        
        volume_results = [
            ('Breakeven Volume Multiplier', 'B20'),
            ('Base Volume (Average)', 'B21'),
            ('Breakeven Volume', 'B22'),
            ('Target NPV', 'B23')
        ]
        
        for label, cell_ref in volume_results:
            worksheet.write(row, 0, label, formats['result_label'])
            worksheet.write(row, 1, '', formats['result_cell'])
            if 'Volume' in label and 'Multiplier' not in label:
                worksheet.write(row, 1, '', formats['number'])
            elif 'Multiplier' in label:
                worksheet.write(row, 1, '', formats['number_2dec'])
            elif 'NPV' in label:
                worksheet.write(row, 1, '', formats['currency_2dec'])
            row += 1
        
        row += 1
        
        # Breakeven Streaming Results
        worksheet.write(row, 0, 'Breakeven Streaming Percentage', formats['result_label'])
        row += 1
        
        streaming_results = [
            ('Breakeven Streaming %', 'B25'),
            ('Current Streaming %', 'B26'),
            ('Target NPV', 'B27')
        ]
        
        for label, cell_ref in streaming_results:
            worksheet.write(row, 0, label, formats['result_label'])
            worksheet.write(row, 1, '', formats['result_cell'])
            if 'Streaming' in label or '%' in label:
                worksheet.write(row, 1, '', formats['percent'])
            elif 'NPV' in label:
                worksheet.write(row, 1, '', formats['currency_2dec'])
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
        worksheet.merge_range(row, 4, row + 15, 7, '', formats['note'])
        
        worksheet.set_column(0, 0, 40)
        worksheet.set_column(1, 1, 25)
        worksheet.set_column(2, 3, 40)
        worksheet.set_column(4, 7, 20)  # Reserve space for charts
        
        return worksheet

