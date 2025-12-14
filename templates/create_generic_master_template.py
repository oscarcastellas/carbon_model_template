#!/usr/bin/env python3
"""
Generic Master Template Creator

Creates a generic, reusable Excel template with:
- Generic "Investor" labels (supporting custom company names)
- Configurable year count (default: 20 years)
- All formulas pre-configured
- Professional formatting
- All sheets (Inputs, Valuation, Summary, Analysis modules)
- VBA support (.xlsm format)

Based on the old version's proven structure.
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

import xlsxwriter
import pandas as pd
import numpy as np
from typing import Optional


class GenericTemplateCreator:
    """Creates a generic master template for carbon credit investment analysis."""
    
    def __init__(self, company_name: str = "Investor", num_years: int = 20, start_year: int = 2025):
        """
        Initialize template creator.
        
        Parameters:
        -----------
        company_name : str
            Generic company name to use (default: "Investor")
        num_years : int
            Number of years in the model (default: 20)
        start_year : int
            Starting year for the model (default: 2025)
        """
        self.company_name = company_name
        self.num_years = num_years
        self.start_year = start_year
    
    def create_template(self, output_path: str) -> None:
        """
        Create the master template file.
        
        Parameters:
        -----------
        output_path : str
            Path where template will be saved (.xlsx file - xlsxwriter cannot create .xlsm)
        """
        print("=" * 70)
        print("CREATING GENERIC MASTER TEMPLATE")
        print("=" * 70)
        print(f"Company Name: {self.company_name}")
        print(f"Years: {self.num_years} ({self.start_year}-{self.start_year + self.num_years - 1})")
        print()
        
        # Note: xlsxwriter cannot create .xlsm files (only .xlsx)
        # If .xlsm is requested, we'll create .xlsx and rename it
        # The file can be converted to .xlsm later if VBA is needed
        actual_output = output_path
        if output_path.endswith('.xlsm'):
            # Create as .xlsx first (xlsxwriter limitation)
            actual_output = output_path.replace('.xlsm', '.xlsx')
            print(f"Note: Creating as .xlsx (xlsxwriter limitation). Can be converted to .xlsm later if needed.")
            print()
        
        # Create workbook (.xlsx format - xlsxwriter limitation)
        workbook = xlsxwriter.Workbook(actual_output, {'nan_inf_to_errors': True})
        
        # Define formats
        formats = self._create_formats(workbook)
        
        # Sheet 1: Inputs & Assumptions
        print("Creating: Inputs & Assumptions")
        inputs_sheet = workbook.add_worksheet('Inputs & Assumptions')
        self._write_inputs_sheet(workbook, inputs_sheet, formats)
        
        # Sheet 2: Valuation Schedule
        print("Creating: Valuation Schedule")
        valuation_sheet = workbook.add_worksheet('Valuation Schedule')
        self._write_valuation_schedule(workbook, valuation_sheet, formats, inputs_sheet)
        
        # Sheet 3: Summary & Results
        print("Creating: Summary & Results")
        summary_sheet = workbook.add_worksheet('Summary & Results')
        self._write_summary_sheet(workbook, summary_sheet, formats, inputs_sheet, valuation_sheet)
        
        # Sheet 4: Analysis (separator)
        print("Creating: Analysis (separator)")
        analysis_sheet = workbook.add_worksheet('Analysis')
        self._write_analysis_separator(workbook, analysis_sheet, formats)
        
        # Sheet 5-8: Interactive Analysis Sheets (placeholders)
        print("Creating: Deal Valuation (interactive)")
        deal_sheet = workbook.add_worksheet('Deal Valuation')
        self._write_deal_valuation_placeholder(workbook, deal_sheet, formats)
        
        print("Creating: Monte Carlo Results (interactive)")
        mc_sheet = workbook.add_worksheet('Monte Carlo Results')
        self._write_monte_carlo_placeholder(workbook, mc_sheet, formats)
        
        print("Creating: Sensitivity Analysis (interactive)")
        sens_sheet = workbook.add_worksheet('Sensitivity Analysis')
        self._write_sensitivity_placeholder(workbook, sens_sheet, formats)
        
        print("Creating: Breakeven Analysis (interactive)")
        be_sheet = workbook.add_worksheet('Breakeven Analysis')
        self._write_breakeven_placeholder(workbook, be_sheet, formats)
        
        workbook.close()
        
        # If .xlsm was requested but we created .xlsx, rename it
        if output_path.endswith('.xlsm') and actual_output != output_path:
            import shutil
            shutil.move(actual_output, output_path)
            print(f"Note: File renamed to {output_path} (but is actually .xlsx format)")
        
        print()
        print("=" * 70)
        print(f"✓ Template created successfully: {output_path}")
        print("=" * 70)
    
    def _create_formats(self, workbook: xlsxwriter.Workbook) -> dict:
        """Create formatting styles matching old version."""
        return {
            'header': workbook.add_format({
                'bold': True, 'bg_color': '#366092', 'font_color': 'white',
                'align': 'center', 'valign': 'vcenter', 'border': 1
            }),
            'title': workbook.add_format({
                'bold': True, 'font_size': 14, 'align': 'left'
            }),
            'subtitle': workbook.add_format({
                'bold': True, 'font_size': 12, 'align': 'left', 'bg_color': '#E7E6E6'
            }),
            'input_label': workbook.add_format({
                'bold': True, 'bg_color': '#D9E1F2', 'border': 1, 'align': 'right'
            }),
            'input_value': workbook.add_format({
                'bg_color': '#FFF2CC', 'border': 1, 'num_format': 'General'
            }),
            'formula_cell': workbook.add_format({
                'bg_color': '#E2EFDA', 'border': 1
            }),
            'currency': workbook.add_format({
                'num_format': '$#,##0', 'border': 1
            }),
            'currency_2dec': workbook.add_format({
                'num_format': '$#,##0.00', 'border': 1
            }),
            'currency_formula': workbook.add_format({
                'num_format': '$#,##0.00', 'border': 1, 'bg_color': '#E2EFDA'
            }),
            'percent': workbook.add_format({
                'num_format': '0.00%', 'border': 1
            }),
            'percent_formula': workbook.add_format({
                'num_format': '0.00%', 'border': 1, 'bg_color': '#E2EFDA'
            }),
            'number': workbook.add_format({
                'num_format': '#,##0', 'border': 1
            }),
            'number_2dec': workbook.add_format({
                'num_format': '#,##0.00', 'border': 1
            }),
            'number_formula': workbook.add_format({
                'num_format': '#,##0.00', 'border': 1, 'bg_color': '#E2EFDA'
            }),
            'bold': workbook.add_format({
                'bold': True, 'border': 1
            }),
            'bold_currency': workbook.add_format({
                'bold': True, 'num_format': '$#,##0.00', 'border': 1,
                'bg_color': '#D9E1F2'
            }),
            'bold_percent': workbook.add_format({
                'bold': True, 'num_format': '0.00%', 'border': 1,
                'bg_color': '#D9E1F2'
            }),
            'text': workbook.add_format({
                'border': 1
            })
        }
    
    def _write_inputs_sheet(self, workbook, worksheet, formats):
        """Write Inputs & Assumptions sheet with generic labels."""
        row = 0
        
        # Title
        title = f'Carbon Credit Investment Model - Inputs & Assumptions'
        worksheet.write(row, 0, title, formats['title'])
        row += 2
        
        # Base Financial Assumptions
        worksheet.write(row, 0, 'Base Financial Assumptions', formats['subtitle'])
        row += 1
        
        # Input labels and placeholder values
        input_labels = [
            'WACC',
            f'{self.company_name} Investment Total',
            'Investment Tenor (Years)',
            'Initial Streaming Percentage',
            'Target IRR',
            'Target Streaming Percentage'
        ]
        
        # Placeholder values (will be populated by GUI)
        input_values = [0.08, 0, 5, 0.48, 0.20, 0.48]
        input_formats = ['percent', 'currency', 'number', 'percent', 'percent', 'percent']
        
        for i, (label, value, fmt) in enumerate(zip(input_labels, input_values, input_formats)):
            worksheet.write(row, 0, label, formats['input_label'])
            if fmt == 'percent':
                worksheet.write(row, 1, value, formats['percent'] if i == 0 or i >= 3 else formats['input_value'])
            elif fmt == 'currency':
                worksheet.write(row, 1, value, formats['input_value'])
            else:
                worksheet.write(row, 1, value, formats['input_value'])
            row += 1
        
        # Monte Carlo Assumptions
        row += 1
        worksheet.write(row, 0, 'Monte Carlo Assumptions', formats['subtitle'])
        row += 1
        
        mc_labels = [
            'Price Growth Base (Mean)',
            'Price Growth Std Dev',
            'Volume Multiplier Base (Mean)',
            'Volume Std Dev',
            'Use GBM Method',
            'GBM Drift (μ)',
            'GBM Volatility (σ)',
            'Number of Simulations'
        ]
        
        mc_values = [0.03, 0.02, 1.0, 0.15, False, 0, 0, 5000]
        
        for i, (label, value) in enumerate(zip(mc_labels, mc_values)):
            worksheet.write(row, 0, label, formats['input_label'])
            if i == 4:  # Use GBM - boolean
                worksheet.write(row, 1, 'No', formats['input_value'])
            elif i == 7:  # Simulations - number
                worksheet.write(row, 1, int(value), formats['number'])
            elif i in [5, 6]:  # GBM parameters
                worksheet.write(row, 1, 'N/A', formats['text'])
            elif i in [0, 1]:  # Price growth - percent
                worksheet.write(row, 1, value, formats['percent'])
            else:  # Volume
                if i == 2:  # Multiplier
                    worksheet.write(row, 1, value, formats['number_2dec'])
                else:  # Std dev
                    worksheet.write(row, 1, value, formats['percent'])
            row += 1
        
        # Set column widths
        worksheet.set_column(0, 0, 35)
        worksheet.set_column(1, 1, 20)
    
    def _write_valuation_schedule(self, workbook, worksheet, formats, inputs_sheet):
        """Write Valuation Schedule with all formulas (matching old version structure)."""
        row = 0
        
        # Title
        worksheet.write(row, 0, f'Valuation Schedule - {self.num_years} Year Cash Flow', formats['title'])
        row += 2
        
        # Get input cell references (matching old version: B3-B8)
        wacc_cell = f"'{inputs_sheet.name}'!$B$3"
        investment_cell = f"'{inputs_sheet.name}'!$B$4"
        tenor_cell = f"'{inputs_sheet.name}'!$B$5"
        streaming_cell = f"'{inputs_sheet.name}'!$B$6"
        
        # Column A: Line item labels
        col_label = 0
        
        # Column B onwards: Years
        year_start_col = 1
        
        # Write year headers horizontally
        header_row = row
        years = [self.start_year + i for i in range(self.num_years)]
        for col_idx, year in enumerate(years):
            col = year_start_col + col_idx
            worksheet.write(header_row, col, year, formats['header'])
        
        # Add "Total" column at the end
        total_col = year_start_col + self.num_years
        worksheet.write(header_row, total_col, 'Total', formats['header'])
        
        row += 1
        
        # Define line items (generic names)
        line_items = [
            {
                'label': 'Carbon Credits Gross',
                'type': 'data',
                'data_col': 'carbon_credits_gross',
                'format': 'number',
                'include_total': True
            },
            {
                'label': f'{self.company_name} Share of Credits',
                'type': 'formula',
                'formula_base': 'credits_share',
                'format': 'number',
                'include_total': True
            },
            {
                'label': 'Base Carbon Price',
                'type': 'data',
                'data_col': 'base_carbon_price',
                'format': 'currency',
                'include_total': False
            },
            {
                'label': f'{self.company_name} Revenue',
                'type': 'formula',
                'formula_base': 'revenue',
                'format': 'currency',
                'include_total': True
            },
            {
                'label': 'Project Implementation Costs',
                'type': 'data',
                'data_col': 'project_implementation_costs',
                'format': 'currency',
                'include_total': True
            },
            {
                'label': f'{self.company_name} Investment Drawdown',
                'type': 'formula',
                'formula_base': 'investment',
                'format': 'currency',
                'include_total': True
            },
            {
                'label': f'{self.company_name} Net Cash Flow',
                'type': 'formula',
                'formula_base': 'net_cf',
                'format': 'currency',
                'include_total': True
            },
            {
                'label': 'Discount Factor',
                'type': 'formula',
                'formula_base': 'discount',
                'format': 'number',
                'include_total': False
            },
            {
                'label': 'Present Value',
                'type': 'formula',
                'formula_base': 'pv',
                'format': 'currency',
                'include_total': True
            },
            {
                'label': 'Cumulative Cash Flow',
                'type': 'formula',
                'formula_base': 'cum_cf',
                'format': 'currency',
                'include_total': False
            },
            {
                'label': 'Cumulative PV',
                'type': 'formula',
                'formula_base': 'cum_pv',
                'format': 'currency',
                'include_total': False
            }
        ]
        
        # Track row positions for formula references
        row_positions = {}
        
        # Write each line item
        for item_idx, item in enumerate(line_items):
            current_row = row + item_idx
            excel_row = current_row + 1  # Excel is 1-based
            
            # Write label in column A
            worksheet.write(current_row, col_label, item['label'], formats['input_label'])
            
            # Store row position for this line item
            row_positions[item['formula_base'] if item['type'] == 'formula' else item['data_col']] = excel_row
            
            # Write data/formulas for each year
            for year_idx in range(self.num_years):
                col = year_start_col + year_idx
                excel_col_letter = xlsxwriter.utility.xl_col_to_name(col)
                
                if item['type'] == 'data':
                    # Empty data cell (will be populated by GUI)
                    if item['format'] == 'currency':
                        worksheet.write(current_row, col, 0, formats['currency_2dec'])
                    else:
                        worksheet.write(current_row, col, 0, formats['number'])
                
                elif item['type'] == 'formula':
                    # Write formula based on type
                    if item['formula_base'] == 'credits_share':
                        # Share = Credits Gross * Streaming %
                        credits_row = row_positions['carbon_credits_gross']
                        formula = f"={excel_col_letter}{credits_row}*{streaming_cell}"
                        worksheet.write_formula(current_row, col, formula, formats['number_formula'])
                    
                    elif item['formula_base'] == 'revenue':
                        # Revenue = Share of Credits * Price
                        share_row = row_positions['credits_share']
                        price_row = row_positions['base_carbon_price']
                        formula = f"={excel_col_letter}{share_row}*{excel_col_letter}{price_row}"
                        worksheet.write_formula(current_row, col, formula, formats['currency_formula'])
                    
                    elif item['formula_base'] == 'investment':
                        # Investment = -Investment/Tenor if Year <= Tenor, else 0
                        year_num = year_idx + 1
                        formula = f"=IF({year_num}<={tenor_cell},-{investment_cell}/{tenor_cell},0)"
                        worksheet.write_formula(current_row, col, formula, formats['currency_formula'])
                    
                    elif item['formula_base'] == 'net_cf':
                        # Net CF = Revenue + Investment
                        revenue_row = row_positions['revenue']
                        investment_row = row_positions['investment']
                        formula = f"={excel_col_letter}{revenue_row}+{excel_col_letter}{investment_row}"
                        worksheet.write_formula(current_row, col, formula, formats['currency_formula'])
                    
                    elif item['formula_base'] == 'discount':
                        # Discount Factor = 1 / (1 + WACC)^(Year - 1)
                        year_num = year_idx + 1
                        formula = f"=1/((1+{wacc_cell})^({year_num}-1))"
                        worksheet.write_formula(current_row, col, formula, formats['number_formula'])
                    
                    elif item['formula_base'] == 'pv':
                        # PV = Net CF * Discount Factor
                        net_cf_row = row_positions['net_cf']
                        discount_row = row_positions['discount']
                        formula = f"={excel_col_letter}{net_cf_row}*{excel_col_letter}{discount_row}"
                        worksheet.write_formula(current_row, col, formula, formats['currency_formula'])
                    
                    elif item['formula_base'] == 'cum_cf':
                        # Cumulative CF = Previous + Current
                        net_cf_row = row_positions['net_cf']
                        if year_idx == 0:
                            formula = f"={excel_col_letter}{net_cf_row}"
                        else:
                            prev_col = xlsxwriter.utility.xl_col_to_name(col - 1)
                            formula = f"={prev_col}{excel_row}+{excel_col_letter}{net_cf_row}"
                        worksheet.write_formula(current_row, col, formula, formats['currency_formula'])
                    
                    elif item['formula_base'] == 'cum_pv':
                        # Cumulative PV = Previous + Current PV
                        pv_row = row_positions['pv']
                        if year_idx == 0:
                            formula = f"={excel_col_letter}{pv_row}"
                        else:
                            prev_col = xlsxwriter.utility.xl_col_to_name(col - 1)
                            formula = f"={prev_col}{excel_row}+{excel_col_letter}{pv_row}"
                        worksheet.write_formula(current_row, col, formula, formats['currency_formula'])
            
            # Write total formula if needed
            if item['include_total']:
                first_year_col = xlsxwriter.utility.xl_col_to_name(year_start_col)
                last_year_col = xlsxwriter.utility.xl_col_to_name(year_start_col + self.num_years - 1)
                sum_formula = f"=SUM({first_year_col}{excel_row}:{last_year_col}{excel_row})"
                
                if item['format'] == 'currency':
                    worksheet.write_formula(current_row, total_col, sum_formula, formats['bold_currency'])
                else:
                    worksheet.write_formula(current_row, total_col, sum_formula, formats['bold'])
        
        # Set column widths
        worksheet.set_column(col_label, col_label, 35)  # Label column
        for i in range(self.num_years):
            worksheet.set_column(year_start_col + i, year_start_col + i, 12)  # Year columns
        worksheet.set_column(total_col, total_col, 15)  # Total column
    
    def _write_summary_sheet(self, workbook, worksheet, formats, inputs_sheet, valuation_sheet):
        """Write Summary & Results sheet with formulas (matching old version)."""
        row = 0
        
        # Title
        worksheet.write(row, 0, 'Summary & Results', formats['title'])
        row += 2
        
        # Key Metrics Section
        worksheet.write(row, 0, 'Key Financial Metrics', formats['subtitle'])
        row += 1
        
        # NPV (sum of all PVs from Valuation Schedule)
        # PV is row 11 (0-indexed), Excel row 12, years are columns B-U (for 20 years) or B-{last_col}
        last_col_letter = xlsxwriter.utility.xl_col_to_name(1 + self.num_years - 1)  # Column B + num_years - 1
        worksheet.write(row, 0, 'Net Present Value (NPV)', formats['input_label'])
        npv_formula = f"=SUM('{valuation_sheet.name}'!B12:{last_col_letter}12)"
        worksheet.write_formula(row, 1, npv_formula, formats['bold_currency'])
        row += 1
        
        # IRR (calculated using Excel IRR function on cash flows)
        worksheet.write(row, 0, 'Internal Rate of Return (IRR)', formats['input_label'])
        # Use Excel's IRR function on the Net Cash Flow row (row 9, Excel row 10)
        irr_formula = f"=IRR('{valuation_sheet.name}'!B10:{last_col_letter}10)"
        worksheet.write_formula(row, 1, irr_formula, formats['bold_percent'])
        worksheet.write(row, 2, '(Calculated from cash flows)', formats['text'])
        row += 1
        
        # Payback Period
        worksheet.write(row, 0, 'Payback Period (Years)', formats['input_label'])
        # Payback is calculated by finding first positive cumulative CF
        # Formula: Find column where cumulative CF becomes positive (row 12, Excel row 13)
        payback_formula = f"=MATCH(0,'{valuation_sheet.name}'!B13:{last_col_letter}13,1)"
        worksheet.write_formula(row, 1, payback_formula, formats['bold'])
        worksheet.write(row, 2, '(Years to positive cumulative cash flow)', formats['text'])
        row += 1
        
        # Target Metrics
        row += 1
        worksheet.write(row, 0, 'Target Metrics', formats['subtitle'])
        row += 1
        
        target_irr_cell = f"'{inputs_sheet.name}'!$B$7"
        target_streaming_cell = f"'{inputs_sheet.name}'!$B$8"
        
        worksheet.write(row, 0, 'Target IRR', formats['input_label'])
        worksheet.write_formula(row, 1, f"={target_irr_cell}", formats['bold_percent'])
        row += 1
        
        worksheet.write(row, 0, 'Target Streaming Percentage', formats['input_label'])
        worksheet.write_formula(row, 1, f"={target_streaming_cell}", formats['bold_percent'])
        row += 1
        
        worksheet.write(row, 0, 'Actual IRR Achieved', formats['input_label'])
        worksheet.write(row, 1, 0, formats['bold_percent'])  # Will be populated
        row += 1
        
        # Set column widths
        worksheet.set_column(0, 0, 35)
        worksheet.set_column(1, 1, 20)
        worksheet.set_column(2, 2, 30)
    
    def _write_analysis_separator(self, workbook, worksheet, formats):
        """Write Analysis separator sheet."""
        worksheet.write(0, 0, 'Analysis Modules', formats['title'])
        note_format = workbook.add_format({'italic': True, 'font_color': '#666666', 'font_size': 9})
        worksheet.write(1, 0, 'Use the sheets below to run advanced analysis modules', note_format)
        worksheet.write(2, 0, 'Input variables in each sheet, then run Python scripts from Terminal', note_format)
    
    def _write_deal_valuation_placeholder(self, workbook, worksheet, formats):
        """Write Deal Valuation placeholder."""
        worksheet.write(0, 0, 'Deal Valuation', formats['title'])
        worksheet.write(2, 0, 'Input Variables:', formats['subtitle'])
        worksheet.write(3, 0, 'Target IRR', formats['input_label'])
        worksheet.write(3, 1, 0.20, formats['percent'])
        worksheet.write(4, 0, 'Calculation Type', formats['input_label'])
        worksheet.write(4, 1, 'Solve for Purchase Price', formats['text'])
        worksheet.write(6, 0, 'Results:', formats['subtitle'])
        worksheet.write(7, 0, 'Maximum Purchase Price', formats['input_label'])
        worksheet.write(8, 0, 'Actual IRR Achieved', formats['input_label'])
        worksheet.write(9, 0, 'Status', formats['input_label'])
        worksheet.write(9, 1, 'Ready - Run Python script', formats['text'])
    
    def _write_monte_carlo_placeholder(self, workbook, worksheet, formats):
        """Write Monte Carlo placeholder."""
        worksheet.write(0, 0, 'Monte Carlo Results', formats['title'])
        worksheet.write(2, 0, 'Input Variables:', formats['subtitle'])
        worksheet.write(3, 0, 'Number of Simulations', formats['input_label'])
        worksheet.write(3, 1, 5000, formats['number'])
        worksheet.write(4, 0, 'Use GBM Method', formats['input_label'])
        worksheet.write(4, 1, 'No', formats['text'])
        worksheet.write(5, 0, 'Status', formats['input_label'])
        worksheet.write(5, 1, 'Ready - Run Python script', formats['text'])
    
    def _write_sensitivity_placeholder(self, workbook, worksheet, formats):
        """Write Sensitivity Analysis placeholder."""
        worksheet.write(0, 0, 'Sensitivity Analysis', formats['title'])
        worksheet.write(2, 0, 'Input Variables:', formats['subtitle'])
        worksheet.write(3, 0, 'Credit Volume Range (Min)', formats['input_label'])
        worksheet.write(3, 1, 0.80, formats['number_2dec'])
        worksheet.write(4, 0, 'Credit Volume Range (Max)', formats['input_label'])
        worksheet.write(4, 1, 1.20, formats['number_2dec'])
        worksheet.write(5, 0, 'Price Multiplier Range (Min)', formats['input_label'])
        worksheet.write(5, 1, 0.90, formats['number_2dec'])
        worksheet.write(6, 0, 'Price Multiplier Range (Max)', formats['input_label'])
        worksheet.write(6, 1, 1.10, formats['number_2dec'])
        worksheet.write(7, 0, 'Status', formats['input_label'])
        worksheet.write(7, 1, 'Ready - Run Python script', formats['text'])
    
    def _write_breakeven_placeholder(self, workbook, worksheet, formats):
        """Write Breakeven Analysis placeholder."""
        worksheet.write(0, 0, 'Breakeven Analysis', formats['title'])
        worksheet.write(2, 0, 'Input Variables:', formats['subtitle'])
        worksheet.write(3, 0, 'Metric to Calculate', formats['input_label'])
        worksheet.write(3, 1, 'all', formats['text'])
        worksheet.write(4, 0, 'Target NPV', formats['input_label'])
        worksheet.write(4, 1, 0, formats['currency'])
        worksheet.write(5, 0, 'Status', formats['input_label'])
        worksheet.write(5, 1, 'Ready - Run Python script', formats['text'])


def main():
    """Main function to create the generic master template."""
    import argparse
    
    parser = argparse.ArgumentParser(description='Create generic master Excel template')
    parser.add_argument('--company-name', type=str, default='Investor',
                       help='Company name to use in template (default: Investor)')
    parser.add_argument('--years', type=int, default=20,
                       help='Number of years in model (default: 20)')
    parser.add_argument('--start-year', type=int, default=2025,
                       help='Starting year (default: 2025)')
    parser.add_argument('--output', type=str, default=None,
                       help='Output file path (default: templates/master_template.xlsm)')
    
    args = parser.parse_args()
    
    # Determine output path
    if args.output:
        output_path = args.output
    else:
        template_dir = Path(__file__).parent
        template_dir.mkdir(exist_ok=True)
        output_path = template_dir / 'master_template.xlsm'
    
    # Create template
    creator = GenericTemplateCreator(
        company_name=args.company_name,
        num_years=args.years,
        start_year=args.start_year
    )
    
    creator.create_template(str(output_path))


if __name__ == '__main__':
    main()

