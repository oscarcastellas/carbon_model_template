"""
Excel Exporter Module: Creates comprehensive, formula-based Excel reports.

This module exports model results with ALL calculations as Excel formulas,
making the model fully auditable and traceable for external review.
"""

import pandas as pd
import xlsxwriter
from typing import Dict, Optional
import numpy as np

# Handle imports for both package and direct usage
try:
    from ..core.payback import PaybackCalculator
except ImportError:
    try:
        from core.payback import PaybackCalculator
    except ImportError:
        from calculators.payback_calculator import PaybackCalculator


class ExcelExporter:
    """
    Exports carbon model results to formula-based Excel files.
    
    Creates comprehensive Excel files where ALL calculations are formulas,
    linked to input assumptions for full transparency and auditability.
    """
    
    def __init__(self):
        """Initialize the Excel Exporter."""
        self.payback_calculator = PaybackCalculator()
    
    def export_model_to_excel(
        self,
        filename: str,
        assumptions: Dict,
        target_streaming_percentage: float,
        target_irr: float,
        actual_irr: float,
        valuation_schedule: pd.DataFrame,
        sensitivity_table: Optional[pd.DataFrame] = None,
        payback_period: Optional[float] = None,
        monte_carlo_results: Optional[Dict] = None,
        risk_flags: Optional[Dict] = None,
        risk_score: Optional[Dict] = None,
        breakeven_results: Optional[Dict] = None,
        deal_valuation_results: Optional[Dict] = None,
        use_template: bool = True
    ) -> None:
        """
        Export complete model to formula-based Excel file.
        
        Creates comprehensive sheets with all calculations as formulas.
        If use_template=True, uses master template with interactive modules.
        
        Parameters:
        -----------
        filename : str
            Output Excel filename
        assumptions : Dict
            Dictionary of model assumptions
        target_streaming_percentage : float
            Target streaming percentage from goal-seeking
        target_irr : float
            Target IRR
        actual_irr : float
            Actual IRR achieved
        valuation_schedule : pd.DataFrame
            Detailed cash flow schedule
        sensitivity_table : pd.DataFrame, optional
            Sensitivity analysis table
        payback_period : float, optional
            Calculated payback period
        monte_carlo_results : Dict, optional
            Monte Carlo simulation results
        use_template : bool, optional
            If True, use master template with interactive modules (default: True)
        """
        # Try template-based export first (if enabled and template exists)
        if use_template:
            try:
                from .template_based_export import TemplateBasedExporter
                template_exporter = TemplateBasedExporter()
                if template_exporter.export_with_template(
                    filename=filename,
                    assumptions=assumptions,
                    target_streaming_percentage=target_streaming_percentage,
                    target_irr=target_irr,
                    actual_irr=actual_irr,
                    valuation_schedule=valuation_schedule,
                    sensitivity_table=sensitivity_table,
                    payback_period=payback_period,
                    monte_carlo_results=monte_carlo_results,
                    risk_flags=risk_flags,
                    risk_score=risk_score,
                    breakeven_results=breakeven_results,
                    deal_valuation_results=deal_valuation_results
                ):
                    # Template export successful!
                    return
            except Exception as e:
                print(f"Template export failed: {e}")
                print("Falling back to standard export...")
        
        # Standard export (fallback or if use_template=False)
        # Create workbook with NaN/INF handling
        workbook = xlsxwriter.Workbook(filename, {'nan_inf_to_errors': True})
        
        # Define formats
        formats = self._create_formats(workbook)
        
        # Sheet 1: Inputs & Assumptions (all inputs in one place)
        inputs_sheet = workbook.add_worksheet('Inputs & Assumptions')
        self._write_inputs_sheet(
            workbook, inputs_sheet, formats, assumptions,
            target_streaming_percentage, target_irr
        )
        
        # Sheet 2: Valuation Schedule (with formulas)
        valuation_sheet = workbook.add_worksheet('Valuation Schedule')
        self._write_valuation_schedule_with_formulas(
            workbook, valuation_sheet, formats, valuation_schedule, inputs_sheet
        )
        
        # Sheet 3: Summary & Results (key metrics with formulas)
        summary_sheet = workbook.add_worksheet('Summary & Results')
        self._write_summary_results_sheet(
            workbook, summary_sheet, formats, valuation_schedule,
            inputs_sheet, actual_irr, payback_period, monte_carlo_results,
            risk_flags, risk_score, breakeven_results
        )
        
        # Sheet 4: Sensitivity Analysis
        if sensitivity_table is not None:
            sens_sheet = workbook.add_worksheet('Sensitivity Analysis')
            self._write_sensitivity_sheet(
                workbook, sens_sheet, formats, sensitivity_table
            )
        
        # Sheet 5: Monte Carlo Results
        if monte_carlo_results is not None:
            mc_sheet = workbook.add_worksheet('Monte Carlo Results')
            self._write_monte_carlo_sheet(
                workbook, mc_sheet, formats, monte_carlo_results
            )
        
        # Sheet 6: Deal Valuation (Back-Solver Results)
        if deal_valuation_results is not None:
            deal_sheet = workbook.add_worksheet('Deal Valuation')
            self._write_deal_valuation_sheet(
                workbook, deal_sheet, formats, deal_valuation_results,
                inputs_sheet, summary_sheet
            )
        
        # Sheet 7: Deal Valuation Interactive (NEW - for running from Excel)
        try:
            from .deal_valuation_interactive import InteractiveDealValuationSheet
            interactive_creator = InteractiveDealValuationSheet(workbook)
            interactive_sheet = interactive_creator.create_interactive_sheet(
                base_assumptions=assumptions,
                sheet_name="Deal Valuation Interactive"
            )
        except ImportError:
            # If module not available, skip interactive sheet
            pass
        
        # Sheet 8: Sensitivity Analysis Interactive (NEW - for running from Excel)
        try:
            from .sensitivity_interactive import InteractiveSensitivitySheet
            sensitivity_creator = InteractiveSensitivitySheet(workbook)
            sensitivity_sheet = sensitivity_creator.create_interactive_sheet(
                base_assumptions=assumptions,
                sheet_name="Sensitivity Interactive"
            )
        except ImportError:
            # If module not available, skip interactive sheet
            pass
        
        # Sheet 9: Monte Carlo Interactive (NEW - for running from Excel)
        try:
            from .monte_carlo_interactive import InteractiveMonteCarloSheet
            mc_creator = InteractiveMonteCarloSheet(workbook)
            mc_sheet = mc_creator.create_interactive_sheet(
                base_assumptions=assumptions,
                sheet_name="Monte Carlo Interactive"
            )
        except ImportError:
            # If module not available, skip interactive sheet
            pass
        
        # Sheet 10: Breakeven Interactive (NEW - for running from Excel)
        try:
            from .breakeven_interactive import InteractiveBreakevenSheet
            breakeven_creator = InteractiveBreakevenSheet(workbook)
            breakeven_sheet = breakeven_creator.create_interactive_sheet(
                base_assumptions=assumptions,
                sheet_name="Breakeven Interactive"
            )
        except ImportError:
            # If module not available, skip interactive sheet
            pass
        
        # Close workbook
        workbook.close()
    
    def _create_formats(self, workbook: xlsxwriter.Workbook) -> Dict:
        """Create formatting styles for Excel output."""
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
    
    def _write_inputs_sheet(
        self,
        workbook: xlsxwriter.Workbook,
        worksheet: xlsxwriter.Workbook.worksheet_class,
        formats: Dict,
        assumptions: Dict,
        target_streaming_percentage: float,
        target_irr: float
    ) -> None:
        """Write Inputs & Assumptions sheet with all model inputs."""
        row = 0
        
        # Title
        worksheet.write(row, 0, 'Carbon Model - Inputs & Assumptions', formats['title'])
        row += 2
        
        # Base Assumptions
        worksheet.write(row, 0, 'Base Financial Assumptions', formats['subtitle'])
        row += 1
        
        # Create named ranges for key inputs
        input_labels = [
            'WACC',
            'Rubicon Investment Total',
            'Investment Tenor (Years)',
            'Initial Streaming Percentage',
            'Target IRR',
            'Target Streaming Percentage'
        ]
        
        input_values = [
            assumptions.get('wacc', 0),
            assumptions.get('rubicon_investment_total', 0),
            assumptions.get('investment_tenor', 0),
            assumptions.get('streaming_percentage_initial', 0),
            target_irr,
            target_streaming_percentage
        ]
        
        input_formats = ['percent', 'currency', 'number', 'percent', 'percent', 'percent']
        
        for i, (label, value, fmt) in enumerate(zip(input_labels, input_values, input_formats)):
            worksheet.write(row, 0, label, formats['input_label'])
            worksheet.write(row, 1, value, formats[input_formats[i] if fmt == 'percent' else 'input_value'])
            # Create named range for formula references
            cell_ref = xlsxwriter.utility.xl_rowcol_to_cell(row, 1)
            workbook.define_name(f'Input_{label.replace(" ", "_").replace("(", "").replace(")", "")}', f"'{worksheet.name}'!{cell_ref}")
            row += 1
        
        # Monte Carlo Assumptions
        if 'price_growth_base' in assumptions:
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
            
            mc_values = [
                assumptions.get('price_growth_base', 0),
                assumptions.get('price_growth_std_dev', 0),
                assumptions.get('volume_multiplier_base', 1.0),
                assumptions.get('volume_std_dev', 0),
                assumptions.get('use_gbm', False),
                assumptions.get('gbm_drift', 0),
                assumptions.get('gbm_volatility', 0),
                assumptions.get('simulations', 5000)
            ]
            
            for i, (label, value) in enumerate(zip(mc_labels, mc_values)):
                worksheet.write(row, 0, label, formats['input_label'])
                if i == 4:  # Use GBM - boolean
                    worksheet.write(row, 1, 'Yes' if value else 'No', formats['input_value'])
                elif i == 7:  # Simulations - number
                    worksheet.write(row, 1, int(value) if value else 5000, formats['number'])
                elif i in [5, 6]:  # GBM parameters - percent
                    if value:
                        worksheet.write(row, 1, value, formats['percent'])
                    else:
                        worksheet.write(row, 1, 'N/A', formats['text'])
                elif i in [0, 1]:  # Price growth - percent
                    worksheet.write(row, 1, value, formats['percent'])
                else:  # Volume - number or percent
                    if i == 2:  # Multiplier
                        worksheet.write(row, 1, value, formats['number_2dec'])
                    else:  # Std dev
                        worksheet.write(row, 1, value, formats['percent'])
                row += 1
        
        # Auto-adjust column widths
        worksheet.set_column(0, 0, 35)
        worksheet.set_column(1, 1, 20)
    
    def _write_valuation_schedule_with_formulas(
        self,
        workbook: xlsxwriter.Workbook,
        worksheet: xlsxwriter.Workbook.worksheet_class,
        formats: Dict,
        valuation_schedule: pd.DataFrame,
        inputs_sheet: xlsxwriter.Workbook.worksheet_class,
        start_year: int = 2025
    ) -> None:
        """
        Write Valuation Schedule in best-practice format: years horizontally, line items vertically.
        
        Parameters:
        -----------
        start_year : int
            Starting year for the model (default: 2025)
        """
        row = 0
        
        # Title
        worksheet.write(row, 0, 'Valuation Schedule - 20 Year Cash Flow', formats['title'])
        row += 2
        
        # Get input cell references
        wacc_cell = f"'{inputs_sheet.name}'!$B$3"
        investment_cell = f"'{inputs_sheet.name}'!$B$4"
        tenor_cell = f"'{inputs_sheet.name}'!$B$5"
        streaming_cell = f"'{inputs_sheet.name}'!$B$6"
        
        # Column A: Line item labels
        col_label = 0
        
        # Column B onwards: Years (2025, 2026, ..., 2044)
        year_start_col = 1
        
        # Write year headers horizontally
        header_row = row
        years = [start_year + i for i in range(20)]
        for col_idx, year in enumerate(years):
            col = year_start_col + col_idx
            worksheet.write(header_row, col, year, formats['header'])
        
        # Add "Total" column at the end
        total_col = year_start_col + 20
        worksheet.write(header_row, total_col, 'Total', formats['header'])
        
        row += 1
        
        # Define line items (in order they should appear)
        line_items = [
            {
                'label': 'Carbon Credits Gross',
                'type': 'data',
                'data_col': 'carbon_credits_gross',
                'format': 'number',
                'include_total': True
            },
            {
                'label': 'Rubicon Share of Credits',
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
                'label': 'Rubicon Revenue',
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
                'label': 'Rubicon Investment Drawdown',
                'type': 'formula',
                'formula_base': 'investment',
                'format': 'currency',
                'include_total': True
            },
            {
                'label': 'Rubicon Net Cash Flow',
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
            for year_idx, year in enumerate(valuation_schedule.index):
                col = year_start_col + year_idx
                excel_col_letter = xlsxwriter.utility.xl_col_to_name(col)
                
                if item['type'] == 'data':
                    # Write data value
                    value = valuation_schedule.loc[year, item['data_col']]
                    if item['format'] == 'currency':
                        worksheet.write(current_row, col, value, formats['currency_2dec'])
                    else:
                        worksheet.write(current_row, col, value, formats['number'])
                
                elif item['type'] == 'formula':
                    # Write formula based on type
                    if item['formula_base'] == 'credits_share':
                        # Rubicon Share = Credits Gross * Streaming %
                        credits_row = row_positions['carbon_credits_gross']
                        formula = f"={xlsxwriter.utility.xl_col_to_name(col)}{credits_row}*{streaming_cell}"
                        worksheet.write_formula(current_row, col, formula, formats['number_formula'])
                    
                    elif item['formula_base'] == 'revenue':
                        # Revenue = Share of Credits * Price
                        share_row = row_positions['credits_share']
                        price_row = row_positions['base_carbon_price']
                        formula = f"={xlsxwriter.utility.xl_col_to_name(col)}{share_row}*{xlsxwriter.utility.xl_col_to_name(col)}{price_row}"
                        worksheet.write_formula(current_row, col, formula, formats['currency_formula'])
                    
                    elif item['formula_base'] == 'investment':
                        # Investment = -Investment/Tenor if Year <= Tenor, else 0
                        # Year is in header, so we use year_idx + 1 (Year 1, 2, ...)
                        year_num = year_idx + 1
                        formula = f"=IF({year_num}<={tenor_cell},-{investment_cell}/{tenor_cell},0)"
                        worksheet.write_formula(current_row, col, formula, formats['currency_formula'])
                    
                    elif item['formula_base'] == 'net_cf':
                        # Net CF = Revenue + Investment
                        revenue_row = row_positions['revenue']
                        investment_row = row_positions['investment']
                        formula = f"={xlsxwriter.utility.xl_col_to_name(col)}{revenue_row}+{xlsxwriter.utility.xl_col_to_name(col)}{investment_row}"
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
                        formula = f"={xlsxwriter.utility.xl_col_to_name(col)}{net_cf_row}*{xlsxwriter.utility.xl_col_to_name(col)}{discount_row}"
                        worksheet.write_formula(current_row, col, formula, formats['currency_formula'])
                    
                    elif item['formula_base'] == 'cum_cf':
                        # Cumulative CF = Previous + Current
                        net_cf_row = row_positions['net_cf']
                        if year_idx == 0:
                            formula = f"={xlsxwriter.utility.xl_col_to_name(col)}{net_cf_row}"
                        else:
                            prev_col = xlsxwriter.utility.xl_col_to_name(col - 1)
                            formula = f"={prev_col}{excel_row}+{xlsxwriter.utility.xl_col_to_name(col)}{net_cf_row}"
                        worksheet.write_formula(current_row, col, formula, formats['currency_formula'])
                    
                    elif item['formula_base'] == 'cum_pv':
                        # Cumulative PV = Previous + Current PV
                        pv_row = row_positions['pv']
                        if year_idx == 0:
                            formula = f"={xlsxwriter.utility.xl_col_to_name(col)}{pv_row}"
                        else:
                            prev_col = xlsxwriter.utility.xl_col_to_name(col - 1)
                            formula = f"={prev_col}{excel_row}+{xlsxwriter.utility.xl_col_to_name(col)}{pv_row}"
                        worksheet.write_formula(current_row, col, formula, formats['currency_formula'])
            
            # Write total formula if needed
            if item['include_total']:
                first_year_col = xlsxwriter.utility.xl_col_to_name(year_start_col)
                last_year_col = xlsxwriter.utility.xl_col_to_name(year_start_col + 19)
                sum_formula = f"=SUM({first_year_col}{excel_row}:{last_year_col}{excel_row})"
                
                if item['format'] == 'currency':
                    worksheet.write_formula(current_row, total_col, sum_formula, formats['bold_currency'])
            else:
                    worksheet.write_formula(current_row, total_col, sum_formula, formats['bold'])
        
        # Add separator row before totals (optional - can add blank row)
        row += len(line_items)
        
        # Set column widths
        worksheet.set_column(col_label, col_label, 35)  # Label column
        for i in range(20):
            worksheet.set_column(year_start_col + i, year_start_col + i, 12)  # Year columns
        worksheet.set_column(total_col, total_col, 15)  # Total column
    
    def _write_summary_results_sheet(
        self,
        workbook: xlsxwriter.Workbook,
        worksheet: xlsxwriter.Workbook.worksheet_class,
        formats: Dict,
        valuation_schedule: pd.DataFrame,
        inputs_sheet: xlsxwriter.Workbook.worksheet_class,
        actual_irr: float,
        payback_period: Optional[float],
        monte_carlo_results: Optional[Dict],
        risk_flags: Optional[Dict] = None,
        risk_score: Optional[Dict] = None,
        breakeven_results: Optional[Dict] = None
    ) -> None:
        """Write Summary & Results sheet with key metrics as formulas."""
        row = 0
        
        # Title
        worksheet.write(row, 0, 'Summary & Results', formats['title'])
        row += 2
        
        # Key Metrics Section
        worksheet.write(row, 0, 'Key Financial Metrics', formats['subtitle'])
        row += 1
        
        # NPV (sum of all PVs from Valuation Schedule)
        # PV is row 11 (0-indexed), Excel row 12, years are columns B-U (2025-2044)
        worksheet.write(row, 0, 'Net Present Value (NPV)', formats['input_label'])
        npv_formula = "=SUM('Valuation Schedule'!B12:U12)"
        worksheet.write_formula(row, 1, npv_formula, formats['bold_currency'])
        row += 1
        
        # IRR (calculated using Excel IRR function on cash flows)
        worksheet.write(row, 0, 'Internal Rate of Return (IRR)', formats['input_label'])
        # Use Excel's IRR function on the Net Cash Flow row (row 9, Excel row 10, columns B-U)
        irr_formula = "=IRR('Valuation Schedule'!B10:U10)"
        worksheet.write_formula(row, 1, irr_formula, formats['bold_percent'])
        worksheet.write(row, 2, f'(Python calculated: {actual_irr:.2%})', formats['text'])
        row += 1
        
        # Payback Period
        if payback_period is not None:
            worksheet.write(row, 0, 'Payback Period (Years)', formats['input_label'])
            # Payback is calculated by finding first positive cumulative CF
            # Formula: Find column where cumulative CF becomes positive (row 12, Excel row 13, columns B-U)
            payback_formula = "=MATCH(0,'Valuation Schedule'!B13:U13,1)"
            worksheet.write_formula(row, 1, payback_formula, formats['bold'])
            worksheet.write(row, 2, f'(Actual: {payback_period:.2f} years)', formats['text'])
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
        worksheet.write(row, 1, actual_irr, formats['bold_percent'])
        row += 1
        
        # Monte Carlo Summary
        if monte_carlo_results is not None:
            row += 1
            worksheet.write(row, 0, 'Monte Carlo Simulation Summary', formats['subtitle'])
            row += 1
            
            mc_metrics = [
                ('MC Mean IRR', 'mc_mean_irr', 'percent'),
                ('MC P10 IRR (Downside)', 'mc_p10_irr', 'percent'),
                ('MC P90 IRR (Upside)', 'mc_p90_irr', 'percent'),
                ('MC Mean NPV', 'mc_mean_npv', 'currency'),
                ('MC P10 NPV', 'mc_p10_npv', 'currency'),
                ('MC P90 NPV', 'mc_p90_npv', 'currency'),
            ]
            
            for label, key, fmt_type in mc_metrics:
                worksheet.write(row, 0, label, formats['input_label'])
                value = monte_carlo_results.get(key, 0)
                # Handle NaN values
                if pd.isna(value) or not np.isfinite(value):
                    worksheet.write(row, 1, 'N/A', formats['text'])
                elif fmt_type == 'percent':
                    worksheet.write(row, 1, value, formats['bold_percent'])
                else:
                    worksheet.write(row, 1, value, formats['bold_currency'])
                row += 1
        
        # Risk Assessment Section
        if risk_flags is not None or risk_score is not None:
            row += 1
            worksheet.write(row, 0, 'Risk Assessment', formats['subtitle'])
            row += 1
            
            # Risk Score
            if risk_score is not None:
                worksheet.write(row, 0, 'Overall Risk Score', formats['input_label'])
                score = risk_score.get('overall_risk_score', 0)
                category = risk_score.get('risk_category', 'Unknown')
                
                # Color code based on risk level
                if category == 'Low':
                    score_format = workbook.add_format({'num_format': '0.0', 'bold': True, 'bg_color': '#C6EFCE'})
                elif category == 'Medium':
                    score_format = workbook.add_format({'num_format': '0.0', 'bold': True, 'bg_color': '#FFEB9C'})
                else:  # High
                    score_format = workbook.add_format({'num_format': '0.0', 'bold': True, 'bg_color': '#FFC7CE'})
                
                worksheet.write(row, 1, score, score_format)
                worksheet.write(row, 2, f'({category} Risk)', formats['text'])
                row += 1
                
                # Component risk scores
                worksheet.write(row, 0, '  Financial Risk', formats['text'])
                worksheet.write(row, 1, risk_score.get('financial_risk', 0), formats['number'])
                row += 1
                
                worksheet.write(row, 0, '  Volume Risk', formats['text'])
                worksheet.write(row, 1, risk_score.get('volume_risk', 0), formats['number'])
                row += 1
                
                worksheet.write(row, 0, '  Price Risk', formats['text'])
                worksheet.write(row, 1, risk_score.get('price_risk', 0), formats['number'])
                row += 1
                
                worksheet.write(row, 0, '  Operational Risk', formats['text'])
                worksheet.write(row, 1, risk_score.get('operational_risk', 0), formats['number'])
                row += 1
            
            # Risk Flags
            if risk_flags is not None:
                row += 1
                risk_level = risk_flags.get('risk_level', 'unknown')
                risk_level_format = workbook.add_format({'bold': True})
                
                if risk_level == 'red':
                    risk_level_format.set_bg_color('#FFC7CE')
                    risk_level_text = '🔴 HIGH RISK'
                elif risk_level == 'yellow':
                    risk_level_format.set_bg_color('#FFEB9C')
                    risk_level_text = '🟡 MEDIUM RISK'
                else:
                    risk_level_format.set_bg_color('#C6EFCE')
                    risk_level_text = '🟢 LOW RISK'
                
                worksheet.write(row, 0, 'Risk Level', formats['input_label'])
                worksheet.write(row, 1, risk_level_text, risk_level_format)
                row += 1
                
                # Flag counts
                flag_counts = risk_flags.get('flag_count', {})
                worksheet.write(row, 0, '  Red Flags', formats['text'])
                worksheet.write(row, 1, flag_counts.get('red', 0), formats['number'])
                row += 1
                
                worksheet.write(row, 0, '  Yellow Flags', formats['text'])
                worksheet.write(row, 1, flag_counts.get('yellow', 0), formats['number'])
                row += 1
                
                # List ALL flags with descriptions
                red_flags = risk_flags.get('red_flags', [])
                yellow_flags = risk_flags.get('yellow_flags', [])
                green_flags = risk_flags.get('green_flags', [])
                
                if red_flags:
                    row += 1
                    worksheet.write(row, 0, '🚨 Critical Risks (Red Flags):', formats['subtitle'])
                    row += 1
                    for i, flag in enumerate(red_flags, 1):
                        worksheet.write(row, 0, f'{i}. {flag}', formats['text'])
                        row += 1
                
                if yellow_flags:
                    row += 1
                    worksheet.write(row, 0, '⚠️  Warnings (Yellow Flags):', formats['subtitle'])
                    row += 1
                    for i, flag in enumerate(yellow_flags, 1):
                        worksheet.write(row, 0, f'{i}. {flag}', formats['text'])
                        row += 1
                
                if green_flags and (not red_flags and not yellow_flags):
                    row += 1
                    worksheet.write(row, 0, '✅ Positive Indicators:', formats['subtitle'])
                    row += 1
                    for i, flag in enumerate(green_flags, 1):
                        worksheet.write(row, 0, f'{i}. {flag}', formats['text'])
                        row += 1
        
        # Breakeven Analysis Section
        if breakeven_results is not None:
            row += 1
            worksheet.write(row, 0, 'Breakeven Analysis', formats['subtitle'])
            row += 1
            
            # Breakeven Price
            if 'breakeven_price' in breakeven_results:
                be_price = breakeven_results['breakeven_price']
                if be_price and 'breakeven_price' in be_price and not pd.isna(be_price.get('breakeven_price')):
                    worksheet.write(row, 0, 'Breakeven Carbon Price', formats['input_label'])
                    worksheet.write(row, 1, be_price['breakeven_price'], formats['bold_currency'])
                    if 'base_price' in be_price:
                        multiplier = be_price.get('price_multiplier', 1.0)
                        worksheet.write(row, 2, f'({multiplier:.2f}x base price)', formats['text'])
                    row += 1
            
            # Breakeven Volume
            if 'breakeven_volume' in breakeven_results:
                be_volume = breakeven_results['breakeven_volume']
                if be_volume and 'breakeven_volume_multiplier' in be_volume and not pd.isna(be_volume.get('breakeven_volume_multiplier')):
                    worksheet.write(row, 0, 'Breakeven Volume Multiplier', formats['input_label'])
                    worksheet.write(row, 1, be_volume['breakeven_volume_multiplier'], formats['bold_percent'])
                    row += 1
            
            # Breakeven Streaming
            if 'breakeven_streaming' in breakeven_results:
                be_streaming = breakeven_results['breakeven_streaming']
                if be_streaming and 'breakeven_streaming' in be_streaming and not pd.isna(be_streaming.get('breakeven_streaming')):
                    worksheet.write(row, 0, 'Breakeven Streaming %', formats['input_label'])
                    worksheet.write(row, 1, be_streaming['breakeven_streaming'], formats['bold_percent'])
                    row += 1
        
        # Auto-adjust column widths
        worksheet.set_column(0, 0, 35)
        worksheet.set_column(1, 1, 20)
        worksheet.set_column(2, 2, 30)
    
    def _write_sensitivity_sheet(
        self,
        workbook: xlsxwriter.Workbook,
        worksheet: xlsxwriter.Workbook.worksheet_class,
        formats: Dict,
        sensitivity_table: pd.DataFrame
    ) -> None:
        """Write Sensitivity Analysis sheet."""
        row = 0
        
        worksheet.write(row, 0, 'Sensitivity Analysis - IRR by Credit Volume and Price', formats['title'])
        row += 2
        
        # Write headers
        col = 1
        worksheet.write(row, 0, sensitivity_table.index.name or 'Credit Volume Multiplier', formats['header'])
        for price_mult in sensitivity_table.columns:
            worksheet.write(row, col, price_mult, formats['header'])
            col += 1
        row += 1
        
        # Write data (values, but could be formulas if we recalculate)
        for credit_mult in sensitivity_table.index:
            col = 0
            worksheet.write(row, col, credit_mult, formats['text'])
            col += 1
            for price_mult in sensitivity_table.columns:
                irr_value = sensitivity_table.loc[credit_mult, price_mult]
                if pd.notna(irr_value):
                    worksheet.write(row, col, irr_value, formats['percent'])
                else:
                    worksheet.write(row, col, 'N/A', formats['text'])
                col += 1
            row += 1
        
        worksheet.set_column(0, 0, 25)
        for i in range(1, len(sensitivity_table.columns) + 1):
            worksheet.set_column(i, i, 15)
    
    def _write_monte_carlo_sheet(
        self,
        workbook: xlsxwriter.Workbook,
        worksheet: xlsxwriter.Workbook.worksheet_class,
        formats: Dict,
        monte_carlo_results: Dict
    ) -> None:
        """Write Monte Carlo Results sheet with histogram."""
        row = 0
        
        worksheet.write(row, 0, 'Monte Carlo Simulation Results', formats['title'])
        row += 1
        
        # Show which method was used
        method_used = monte_carlo_results.get('method_used', 'Growth-Rate Based')
        method_format = workbook.add_format({
            'bold': True,
            'bg_color': '#D9E1F2',
            'border': 1
        })
        worksheet.write(row, 0, 'Simulation Method:', formats['input_label'])
        worksheet.write(row, 1, method_used, method_format)
        
        # Show GBM parameters if used
        if monte_carlo_results.get('use_gbm', False):
            row += 1
            worksheet.write(row, 0, 'GBM Drift (μ):', formats['text'])
            gbm_drift = monte_carlo_results.get('gbm_drift')
            if gbm_drift is not None:
                worksheet.write(row, 1, gbm_drift, formats['percent'])
            row += 1
            worksheet.write(row, 0, 'GBM Volatility (σ):', formats['text'])
            gbm_vol = monte_carlo_results.get('gbm_volatility')
            if gbm_vol is not None:
                worksheet.write(row, 1, gbm_vol, formats['percent'])
        
        row += 2
        
        # Summary Statistics
        worksheet.write(row, 0, 'Summary Statistics', formats['subtitle'])
        row += 1
        
        summary_data = [
            ('Mean IRR', 'mc_mean_irr', 'percent'),
            ('P10 IRR (10th Percentile)', 'mc_p10_irr', 'percent'),
            ('P90 IRR (90th Percentile)', 'mc_p90_irr', 'percent'),
            ('Std Dev IRR', 'mc_std_irr', 'percent'),
            ('Mean NPV', 'mc_mean_npv', 'currency'),
            ('P10 NPV', 'mc_p10_npv', 'currency'),
            ('P90 NPV', 'mc_p90_npv', 'currency'),
            ('Std Dev NPV', 'mc_std_npv', 'currency'),
            ('Total Simulations', 'simulations', 'number'),
            ('Valid Simulations', 'valid_simulations', 'number'),
        ]
        
        for label, key, fmt_type in summary_data:
            worksheet.write(row, 0, label, formats['text'])
            value = monte_carlo_results.get(key, 0)
            # Handle NaN values
            if pd.isna(value) or not np.isfinite(value):
                worksheet.write(row, 1, 'N/A', formats['text'])
            elif fmt_type == 'percent':
                worksheet.write(row, 1, value, formats['percent'])
            elif fmt_type == 'currency':
                worksheet.write(row, 1, value, formats['currency_2dec'])
            else:
                worksheet.write(row, 1, value, formats['number'])
            row += 1
        
        row += 2
        
        # Full Results Table
        worksheet.write(row, 0, 'Full Simulation Results', formats['subtitle'])
        row += 1
        
        worksheet.write(row, 0, 'Simulation', formats['header'])
        worksheet.write(row, 1, 'IRR', formats['header'])
        worksheet.write(row, 2, 'NPV', formats['header'])
        row += 1
        
        # Write simulation results
        irr_series = monte_carlo_results.get('irr_series', [])
        npv_series = monte_carlo_results.get('npv_series', [])
        
        for i in range(len(irr_series)):
            worksheet.write(row, 0, i + 1, formats['number'])
            irr_val = irr_series[i]
            npv_val = npv_series[i]
            
            if pd.notna(irr_val):
                worksheet.write(row, 1, irr_val, formats['percent'])
            else:
                worksheet.write(row, 1, 'N/A', formats['text'])
            
            if pd.notna(npv_val):
                worksheet.write(row, 2, npv_val, formats['currency_2dec'])
            else:
                worksheet.write(row, 2, 'N/A', formats['text'])
            
            row += 1
        
        # Create histogram charts
        row += 2
        worksheet.write(row, 0, 'Distribution Charts', formats['subtitle'])
        row += 2
        
        # IRR Histogram
        irr_valid = [x for x in irr_series if pd.notna(x)]
        if len(irr_valid) > 0:
            # Calculate bins for IRR histogram
            min_irr = min(irr_valid)
            max_irr = max(irr_valid)
            num_bins = 40  # More bins for better resolution
            bin_width = (max_irr - min_irr) / num_bins
            
            # Create histogram data
            hist_data_start_row = row
            worksheet.write(row, 0, 'IRR Histogram Data', formats['subtitle'])
            row += 1
            worksheet.write(row, 0, 'IRR Range', formats['header'])
            worksheet.write(row, 1, 'Frequency', formats['header'])
            row += 1
            
            bins = []
            frequencies = []
            hist_data_row_start = row
            
            for i in range(num_bins):
                bin_start = min_irr + i * bin_width
                bin_end = min_irr + (i + 1) * bin_width
                count = sum(1 for x in irr_valid if bin_start <= x < bin_end)
                if i == num_bins - 1:
                    count = sum(1 for x in irr_valid if bin_start <= x <= bin_end)
                bins.append(bin_start)
                frequencies.append(count)
                
                worksheet.write(row, 0, bin_start, formats['percent'])
                worksheet.write(row, 1, count, formats['number'])
                row += 1
            
            hist_data_row_end = row - 1
            
            # Create IRR histogram chart
            irr_chart = workbook.add_chart({'type': 'column'})
            irr_chart.add_series({
                'name': 'IRR Distribution',
                'categories': [f'Monte Carlo Results', hist_data_row_start, 0, hist_data_row_end, 0],
                'values': [f'Monte Carlo Results', hist_data_row_start, 1, hist_data_row_end, 1],
                'gap': 0,  # No gap between columns for histogram look
            })
            
            irr_chart.set_title({
                'name': 'Monte Carlo IRR Distribution',
                'name_font': {'size': 14, 'bold': True}
            })
            irr_chart.set_x_axis({
                'name': 'IRR',
                'num_format': '0.00%',
                'name_font': {'size': 11},
                'num_font': {'size': 9}
            })
            irr_chart.set_y_axis({
                'name': 'Frequency',
                'name_font': {'size': 11},
                'num_font': {'size': 9}
            })
            irr_chart.set_legend({'position': 'none'})  # Hide legend for histogram
            irr_chart.set_size({'width': 720, 'height': 400})
            irr_chart.set_style(10)
            
            # Insert chart below the data (more visible)
            chart_row = hist_data_row_end + 2
            worksheet.insert_chart(chart_row, 0, irr_chart)
            
            row = chart_row + 25  # Leave space for chart
        
        # NPV Histogram
        npv_valid = [x for x in npv_series if pd.notna(x)]
        if len(npv_valid) > 0:
            # Calculate bins for NPV histogram
            min_npv = min(npv_valid)
            max_npv = max(npv_valid)
            num_bins_npv = 40
            bin_width_npv = (max_npv - min_npv) / num_bins_npv
            
            # Create histogram data
            npv_hist_data_start_row = row
            worksheet.write(row, 0, 'NPV Histogram Data', formats['subtitle'])
            row += 1
            worksheet.write(row, 0, 'NPV Range', formats['header'])
            worksheet.write(row, 1, 'Frequency', formats['header'])
            row += 1
            
            npv_bins = []
            npv_frequencies = []
            npv_hist_data_row_start = row
            
            for i in range(num_bins_npv):
                bin_start = min_npv + i * bin_width_npv
                bin_end = min_npv + (i + 1) * bin_width_npv
                count = sum(1 for x in npv_valid if bin_start <= x < bin_end)
                if i == num_bins_npv - 1:
                    count = sum(1 for x in npv_valid if bin_start <= x <= bin_end)
                npv_bins.append(bin_start)
                npv_frequencies.append(count)
                
                worksheet.write(row, 0, bin_start, formats['currency_2dec'])
                worksheet.write(row, 1, count, formats['number'])
                row += 1
            
            npv_hist_data_row_end = row - 1
            
            # Create NPV histogram chart
            npv_chart = workbook.add_chart({'type': 'column'})
            npv_chart.add_series({
                'name': 'NPV Distribution',
                'categories': [f'Monte Carlo Results', npv_hist_data_row_start, 0, npv_hist_data_row_end, 0],
                'values': [f'Monte Carlo Results', npv_hist_data_row_start, 1, npv_hist_data_row_end, 1],
                'gap': 0,  # No gap between columns for histogram look
            })
            
            npv_chart.set_title({
                'name': 'Monte Carlo NPV Distribution',
                'name_font': {'size': 14, 'bold': True}
            })
            npv_chart.set_x_axis({
                'name': 'NPV ($)',
                'num_format': '$#,##0',
                'name_font': {'size': 11},
                'num_font': {'size': 9}
            })
            npv_chart.set_y_axis({
                'name': 'Frequency',
                'name_font': {'size': 11},
                'num_font': {'size': 9}
            })
            npv_chart.set_legend({'position': 'none'})  # Hide legend for histogram
            npv_chart.set_size({'width': 720, 'height': 400})
            npv_chart.set_style(10)
            
            # Insert chart below the data (more visible)
            npv_chart_row = npv_hist_data_row_end + 2
            worksheet.insert_chart(npv_chart_row, 0, npv_chart)
        
        worksheet.set_column(0, 0, 18)
        worksheet.set_column(1, 1, 15)
        worksheet.set_column(2, 2, 20)
    
    def _write_deal_valuation_sheet(
        self,
        workbook: xlsxwriter.Workbook,
        sheet,
        formats: Dict,
        deal_valuation_results: Dict,
        inputs_sheet,
        summary_sheet
    ) -> None:
        """
        Write Deal Valuation sheet with back-solver results.
        
        Creates interactive what-if analysis:
        - Shows solved purchase price given target IRR
        - Shows calculated IRR from purchase price
        - Shows required streaming % given price + IRR
        """
        row = 0
        
        # Header
        sheet.write(row, 0, 'Deal Valuation Analysis', formats['header'])
        sheet.merge_range(row, 0, row, 3, 'Deal Valuation Analysis', formats['header'])
        row += 2
        
        # Determine which solve method was used
        if 'purchase_price' in deal_valuation_results and 'target_irr' in deal_valuation_results:
            # Solve for Purchase Price
            section_header = formats.get('section_header', formats.get('subtitle', formats['header']))
            label_format = formats.get('label', formats.get('input_label', formats['text']))
            sheet.write(row, 0, 'Solve for Purchase Price', section_header)
            row += 1
            
            sheet.write(row, 0, 'Target IRR:', label_format)
            sheet.write(row, 1, deal_valuation_results.get('target_irr', 0), formats['percent'])
            row += 1
            
            sheet.write(row, 0, 'Streaming Percentage:', label_format)
            sheet.write(row, 1, deal_valuation_results.get('streaming_percentage', 0), formats['percent'])
            row += 1
            
            sheet.write(row, 0, 'Maximum Purchase Price:', label_format)
            sheet.write(row, 1, deal_valuation_results.get('purchase_price', 0), formats['currency_2dec'])
            row += 1
            
            sheet.write(row, 0, 'Actual IRR Achieved:', label_format)
            sheet.write(row, 1, deal_valuation_results.get('actual_irr', 0), formats['percent'])
            row += 1
            
            sheet.write(row, 0, 'Difference:', label_format)
            sheet.write(row, 1, deal_valuation_results.get('difference', 0), formats['percent'])
            row += 1
            
            sheet.write(row, 0, 'NPV at Calculated Price:', label_format)
            sheet.write(row, 1, deal_valuation_results.get('npv', 0), formats['currency_2dec'])
            row += 2
        
        elif 'purchase_price' in deal_valuation_results and 'irr' in deal_valuation_results:
            # Calculate IRR from Price
            section_header = formats.get('section_header', formats.get('subtitle', formats['header']))
            label_format = formats.get('label', formats.get('input_label', formats['text']))
            sheet.write(row, 0, 'Calculate IRR from Purchase Price', section_header)
            row += 1
            
            sheet.write(row, 0, 'Purchase Price:', label_format)
            sheet.write(row, 1, deal_valuation_results.get('purchase_price', 0), formats['currency_2dec'])
            row += 1
            
            sheet.write(row, 0, 'Streaming Percentage:', label_format)
            sheet.write(row, 1, deal_valuation_results.get('streaming_percentage', 0), formats['percent'])
            row += 1
            
            sheet.write(row, 0, 'Project IRR:', label_format)
            sheet.write(row, 1, deal_valuation_results.get('irr', 0), formats['percent'])
            row += 1
            
            sheet.write(row, 0, 'NPV:', label_format)
            sheet.write(row, 1, deal_valuation_results.get('npv', 0), formats['currency_2dec'])
            row += 2
        
        elif 'streaming_percentage' in deal_valuation_results and 'target_irr' in deal_valuation_results:
            # Solve for Streaming % given Price + IRR
            section_header = formats.get('section_header', formats.get('subtitle', formats['header']))
            label_format = formats.get('label', formats.get('input_label', formats['text']))
            sheet.write(row, 0, 'Solve for Streaming Percentage', section_header)
            row += 1
            
            sheet.write(row, 0, 'Purchase Price:', label_format)
            sheet.write(row, 1, deal_valuation_results.get('purchase_price', 0), formats['currency_2dec'])
            row += 1
            
            sheet.write(row, 0, 'Target IRR:', label_format)
            sheet.write(row, 1, deal_valuation_results.get('target_irr', 0), formats['percent'])
            row += 1
            
            sheet.write(row, 0, 'Required Streaming Percentage:', label_format)
            sheet.write(row, 1, deal_valuation_results.get('streaming_percentage', 0), formats['percent'])
            row += 1
            
            sheet.write(row, 0, 'Actual IRR Achieved:', label_format)
            sheet.write(row, 1, deal_valuation_results.get('actual_irr', 0), formats['percent'])
            row += 1
            
            sheet.write(row, 0, 'Difference:', label_format)
            sheet.write(row, 1, deal_valuation_results.get('difference', 0), formats['percent'])
            row += 1
            
            sheet.write(row, 0, 'NPV:', label_format)
            sheet.write(row, 1, deal_valuation_results.get('npv', 0), formats['currency_2dec'])
            row += 2
        
        # Investment Tenor
        if 'investment_tenor' in deal_valuation_results:
            label_format = formats.get('label', formats.get('input_label', formats['text']))
            sheet.write(row, 0, 'Investment Tenor:', label_format)
            sheet.write(row, 1, deal_valuation_results.get('investment_tenor', 0), formats['number'])
            row += 1
        
        # Note
        row += 1
        label_format = formats.get('label', formats.get('input_label', formats['text']))
        note_format = formats.get('note', formats.get('text', formats['number']))
        sheet.write(row, 0, 'Note:', label_format)
        sheet.write(row, 1, 'These results are calculated using the back-solver module.', note_format)
        sheet.write(row + 1, 1, 'To re-run analysis, use the Python API methods:', note_format)
        sheet.write(row + 2, 1, '- solve_for_purchase_price(target_irr, streaming_percentage)', note_format)
        sheet.write(row + 3, 1, '- solve_for_project_irr(purchase_price, streaming_percentage)', note_format)
        sheet.write(row + 4, 1, '- solve_for_streaming_given_price(purchase_price, target_irr)', note_format)
        
        # Set column widths
        sheet.set_column(0, 0, 30)
        sheet.set_column(1, 1, 20)
        sheet.set_column(2, 3, 15)
