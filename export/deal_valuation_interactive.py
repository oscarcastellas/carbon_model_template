"""
Interactive Deal Valuation Sheet Creator

Creates an Excel sheet with input cells and a button to run back-solver
from within Excel using xlwings.
"""

import xlsxwriter
from typing import Dict, Optional
import pandas as pd


class InteractiveDealValuationSheet:
    """
    Creates an interactive Excel sheet for deal valuation back-solver.
    
    Users can:
    - Set input variables (Target IRR, Streaming %, Purchase Price)
    - Click a button to run the back-solver
    - See results populated in Excel
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
        sheet_name: str = "Deal Valuation"
    ):
        """
        Create interactive deal valuation sheet.
        
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
        worksheet.write(row, 0, 'Interactive Deal Valuation Back-Solver', formats['title'])
        worksheet.merge_range(row, 0, row, 4, 
                            'Interactive Deal Valuation Back-Solver', formats['title'])
        worksheet.set_row(row, 35)
        row += 2
        
        # Instructions
        instructions = [
            "INSTRUCTIONS:",
            "1. Set your input variables in the 'Input Variables' section below",
            "2. Choose which calculation you want to run",
            "3. Click 'Run Back-Solver' button (or run Python script)",
            "4. Results will appear in the 'Results' section"
        ]
        
        for i, instruction in enumerate(instructions):
            if i == 0:
                worksheet.write(row, 0, instruction, formats['section_header'])
            else:
                worksheet.write(row, 0, instruction, formats['note'])
            row += 1
        
        row += 1
        
        # Section 1: Input Variables
        worksheet.write(row, 0, 'Input Variables', formats['section_header'])
        row += 1
        
        # Input cells (EMPTY - user fills manually)
        inputs = [
            ('Target IRR', 'B8', 'percent', 'Target IRR for purchase price calculation'),
            ('Streaming Percentage', 'B9', 'percent', 'Streaming percentage'),
            ('Purchase Price (for IRR calc)', 'B10', 'currency', 'Purchase price to calculate IRR'),
            ('Investment Tenor (Years)', 'B11', 'number', 'Investment deployment period')
        ]
        
        for label, cell_ref, fmt_type, note in inputs:
            worksheet.write(row, 0, label, formats['input_label'])
            # Write empty cell with proper format
            if fmt_type == 'percent':
                worksheet.write(row, 1, '', formats['percent'])
            elif fmt_type == 'currency':
                worksheet.write(row, 1, '', formats['currency_2dec'])
            else:
                worksheet.write(row, 1, '', formats['number'])
            worksheet.write(row, 2, note, formats['note'])
            row += 1
        
        row += 1
        
        # Section 2: Calculation Type
        worksheet.write(row, 0, 'Calculation Type', formats['section_header'])
        worksheet.merge_range(row, 0, row, 4, '', formats['section_header'])
        row += 1
        
        worksheet.write(row, 0, 'Select Calculation:', formats['input_label'])
        # Empty cell - user selects calculation type
        worksheet.write(row, 1, '', formats['input_cell'])
        worksheet.write(row, 2, 'Options: Solve for Purchase Price, Calculate IRR from Price, Solve for Streaming %', formats['note'])
        row += 1
        
        row += 1
        
        # Instructions for running script
        worksheet.write(row, 0, 'How to Run:', formats['section_header'])
        row += 1
        worksheet.write(row, 0, '1. Fill in input variables above', formats['note'])
        row += 1
        worksheet.write(row, 0, '2. Save this Excel file', formats['note'])
        row += 1
        worksheet.write(row, 0, '3. Run: python3 scripts/run_deal_valuation_from_excel.py "path/to/file.xlsx"', formats['note'])
        row += 1
        worksheet.write(row, 0, '4. Results and charts will appear below', formats['note'])
        row += 2
        
        # Section 4: Results
        worksheet.write(row, 0, 'Results', formats['section_header'])
        row += 1
        
        # Result cells (will be populated by Python script)
        results = [
            ('Maximum Purchase Price', 'B22', 'currency_2dec', 'Result from back-solver'),
            ('Actual IRR Achieved', 'B23', 'percent', 'Actual IRR achieved'),
            ('Target IRR', 'B24', 'percent', 'Target IRR input'),
            ('Difference', 'B25', 'percent', 'Difference between actual and target'),
            ('NPV at Calculated Price', 'B26', 'currency_2dec', 'NPV at calculated price'),
            ('Required Streaming %', 'B27', 'percent', 'Required streaming percentage'),
            ('Project IRR', 'B28', 'percent', 'Project IRR from purchase price')
        ]
        
        for label, cell_ref, fmt_type, note in results:
            worksheet.write(row, 0, label, formats['result_label'])
            worksheet.write(row, 1, '', formats['result_cell'])  # Empty, will be filled
            if fmt_type == 'percent':
                worksheet.write(row, 1, '', formats['percent'])
            elif fmt_type == 'currency_2dec':
                worksheet.write(row, 1, '', formats['currency_2dec'])
            worksheet.write(row, 2, note, formats['note'])
            row += 1
        
        row += 1
        
        # Section 5: Status/Error Messages
        worksheet.write(row, 0, 'Status', formats['input_label'])
        worksheet.write(row, 1, 'Ready', formats['result_cell'])
        worksheet.write(row, 2, 'Status of last calculation', formats['note'])
        row += 2
        
        # Reserve space for charts (columns E-H)
        worksheet.write(row, 4, 'Charts will appear here after running analysis', formats['note'])
        worksheet.merge_range(row, 4, row + 15, 7, '', formats['note'])
        
        # Set column widths
        worksheet.set_column(0, 0, 35)
        worksheet.set_column(1, 1, 25)
        worksheet.set_column(2, 3, 40)
        worksheet.set_column(4, 7, 20)  # Reserve space for charts
        
        # Create named ranges for easy reference
        named_ranges = {
            'Input_TargetIRR': 'B8',
            'Input_StreamingPct': 'B9',
            'Input_PurchasePrice': 'B10',
            'Input_InvestmentTenor': 'B11',
            'Input_CalcType': 'B13',
            'Result_PurchasePrice': 'B22',
            'Result_ActualIRR': 'B23',
            'Result_TargetIRR': 'B24',
            'Result_Difference': 'B25',
            'Result_NPV': 'B26',
            'Result_StreamingPct': 'B27',
            'Result_ProjectIRR': 'B28',
            'Status': 'B30'
        }
        
        for name, cell in named_ranges.items():
            self.workbook.define_name(name, f"'{sheet_name}'!{cell}")
        
        return worksheet

