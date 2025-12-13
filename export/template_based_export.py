"""
Enhanced Template-Based Export

This module provides comprehensive template-based export that populates
all standard sheets while preserving interactive sheets with VBA/buttons.
"""

import os
import shutil
from pathlib import Path
from typing import Dict, Optional
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows


class TemplateBasedExporter:
    """
    Exports Excel files using master template with all sheets and interactive modules.
    
    This ensures users get a fully-functional Excel file with:
    - All standard sheets (Inputs, Valuation, Summary, etc.) - populated with data
    - All interactive sheets (Deal Valuation, Sensitivity, Monte Carlo, Breakeven) - with VBA/buttons
    - Zero setup required!
    """
    
    def __init__(self):
        """Initialize template-based exporter."""
        self.template_path = Path(__file__).parent.parent / "templates" / "master_template_with_interactive_modules.xlsm"
        self.template_path_xlsx = Path(__file__).parent.parent / "templates" / "master_template_with_interactive_modules.xlsx"
    
    def export_with_template(
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
        deal_valuation_results: Optional[Dict] = None
    ) -> bool:
        """
        Export using master template.
        
        Returns True if template was used, False if template not found (fallback needed).
        """
        # Check if completed template (.xlsm) exists, otherwise try .xlsx
        template_file = self.template_path if self.template_path.exists() else self.template_path_xlsx
        
        if not template_file.exists():
            print(f"Template not found. Expected at: {self.template_path}")
            print("Falling back to standard export (without interactive modules)")
            return False
        
        try:
            # Step 1: Copy template to output file
            print(f"Using master template: {template_file.name}")
            shutil.copy(template_file, filename)
            
            # Step 2: Load with openpyxl (preserves VBA if .xlsm)
            keep_vba = template_file.suffix == '.xlsm'
            wb = load_workbook(filename, keep_vba=keep_vba)
            
            # Step 3: Populate all data sheets
            print("Populating data sheets...")
            
            # Populate Inputs & Assumptions
            if 'Inputs & Assumptions' in wb.sheetnames:
                self._populate_inputs_sheet_comprehensive(wb['Inputs & Assumptions'], assumptions, target_streaming_percentage, target_irr)
                # Generate and embed charts
                self._add_presentation_charts_to_inputs(wb['Inputs & Assumptions'], assumptions, target_streaming_percentage)
            
            # Populate Valuation Schedule
            if 'Valuation Schedule' in wb.sheetnames:
                self._populate_valuation_sheet_comprehensive(wb['Valuation Schedule'], valuation_schedule)
                # Generate and embed charts
                self._add_presentation_charts_to_valuation(wb['Valuation Schedule'], valuation_schedule)
            
            # Populate Summary & Results
            if 'Summary & Results' in wb.sheetnames:
                self._populate_summary_sheet_comprehensive(
                    wb['Summary & Results'],
                    valuation_schedule,
                    actual_irr,
                    target_irr,
                    payback_period,
                    monte_carlo_results,
                    risk_flags,
                    risk_score,
                    breakeven_results
                )
                # Generate and embed charts
                self._add_presentation_charts_to_summary(wb['Summary & Results'], actual_irr, target_irr, risk_score)
            
            # Populate Deal Valuation (if results available)
            if 'Deal Valuation' in wb.sheetnames and deal_valuation_results:
                self._populate_deal_valuation_sheet(wb['Deal Valuation'], deal_valuation_results)
            
            # Populate Monte Carlo Results (if available)
            if 'Monte Carlo Results' in wb.sheetnames and monte_carlo_results:
                self._populate_monte_carlo_sheet(wb['Monte Carlo Results'], monte_carlo_results)
            
            # Populate Sensitivity Analysis (if available)
            if 'Sensitivity Analysis' in wb.sheetnames and sensitivity_table is not None:
                self._populate_sensitivity_sheet(wb['Sensitivity Analysis'], sensitivity_table)
            
            # Interactive sheets remain untouched (VBA and buttons preserved)
            print("✓ Interactive sheets preserved (VBA and buttons intact)")
            
            # Step 4: Save (preserves VBA if .xlsm)
            wb.save(filename)
            wb.close()
            
            print(f"✓ Template-based export complete: {filename}")
            print("  All standard sheets populated with data")
            print("  All interactive sheets ready to use (buttons work!)")
            return True
            
        except Exception as e:
            print(f"Error using template: {e}")
            import traceback
            traceback.print_exc()
            print("Falling back to standard export")
            return False
    
    def _populate_inputs_sheet_comprehensive(self, ws, assumptions: Dict, streaming_pct: float, target_irr: float):
        """Comprehensively populate Inputs & Assumptions sheet."""
        # Find and populate all assumption cells
        # This searches for labels and populates corresponding values
        
        # Common assumption locations (will search if not found)
        assumption_mappings = {
            'WACC': assumptions.get('wacc', 0.08),
            'Weighted Average Cost of Capital': assumptions.get('wacc', 0.08),
            'Investment Total': assumptions.get('rubicon_investment_total', 20_000_000),
            'Rubicon Investment': assumptions.get('rubicon_investment_total', 20_000_000),
            'Investment Tenor': assumptions.get('investment_tenor', 5),
            'Tenor': assumptions.get('investment_tenor', 5),
            'Streaming Percentage': streaming_pct,
            'Target Streaming': streaming_pct,
            'Target IRR': target_irr,
            'Price Growth Base': assumptions.get('price_growth_base', 0.03),
            'Price Growth Std Dev': assumptions.get('price_growth_std_dev', 0.02),
            'Volume Multiplier Base': assumptions.get('volume_multiplier_base', 1.0),
            'Volume Std Dev': assumptions.get('volume_std_dev', 0.15),
        }
        
        # Search for labels and populate values
        for row in range(1, min(ws.max_row + 1, 100)):
            for col in range(1, min(ws.max_column + 1, 10)):
                cell = ws.cell(row=row, column=col)
                if cell.value:
                    cell_str = str(cell.value).strip()
                    for label, value in assumption_mappings.items():
                        if label.lower() in cell_str.lower():
                            # Value cell is typically next column
                            value_cell = ws.cell(row=row, column=col + 1)
                            value_cell.value = value
                            if 'percent' in label.lower() or 'irr' in label.lower() or 'growth' in label.lower():
                                value_cell.number_format = '0.00%'
                            elif 'investment' in label.lower() or 'total' in label.lower():
                                value_cell.number_format = '$#,##0'
                            break
    
    def _populate_valuation_sheet_comprehensive(self, ws, valuation_schedule: pd.DataFrame):
        """Comprehensively populate Valuation Schedule sheet."""
        # Clear existing data (keep headers in rows 1-2)
        for row in range(3, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=row, column=col)
                if cell.value and 'Year' not in str(cell.value):
                    cell.value = None
        
        # Write data starting from row 3
        start_row = 3
        for idx, (year, row_data) in enumerate(valuation_schedule.iterrows(), start=start_row):
            # Year
            ws.cell(row=idx, column=1).value = int(year) if pd.notna(year) else idx - start_row + 1
            
            # Cash Flow
            if 'cash_flow' in row_data:
                cf = row_data['cash_flow']
                if pd.notna(cf):
                    ws.cell(row=idx, column=2).value = float(cf)
                    ws.cell(row=idx, column=2).number_format = '$#,##0.00'
            
            # Present Value
            if 'present_value' in row_data:
                pv = row_data['present_value']
                if pd.notna(pv):
                    ws.cell(row=idx, column=3).value = float(pv)
                    ws.cell(row=idx, column=3).number_format = '$#,##0.00'
            
            # Add other columns if they exist in data
            col_idx = 4
            for col_name in ['npv', 'irr', 'cumulative_cf']:
                if col_name in row_data:
                    val = row_data[col_name]
                    if pd.notna(val):
                        ws.cell(row=idx, column=col_idx).value = float(val)
                        if 'irr' in col_name:
                            ws.cell(row=idx, column=col_idx).number_format = '0.00%'
                        else:
                            ws.cell(row=idx, column=col_idx).number_format = '$#,##0.00'
                        col_idx += 1
    
    def _populate_summary_sheet_comprehensive(self, ws, valuation_schedule, actual_irr, target_irr,
                                             payback_period, mc_results, risk_flags, risk_score, breakeven):
        """Comprehensively populate Summary & Results sheet."""
        # Find and populate key metrics
        metrics = {
            'IRR': actual_irr,
            'Actual IRR': actual_irr,
            'Target IRR': target_irr,
            'Payback Period': payback_period,
            'NPV': valuation_schedule['present_value'].sum() if 'present_value' in valuation_schedule.columns else None,
        }
        
        # Add Monte Carlo results if available
        if mc_results:
            metrics['MC Mean IRR'] = mc_results.get('mc_mean_irr')
            metrics['MC Mean NPV'] = mc_results.get('mc_mean_npv')
            metrics['MC P10 IRR'] = mc_results.get('mc_p10_irr')
            metrics['MC P90 IRR'] = mc_results.get('mc_p90_irr')
        
        # Add risk score if available
        if risk_score:
            metrics['Risk Score'] = risk_score.get('overall_score')
        
        # Search and populate
        for row in range(1, min(ws.max_row + 1, 100)):
            for col in range(1, min(ws.max_column + 1, 10)):
                cell = ws.cell(row=row, column=col)
                if cell.value:
                    cell_str = str(cell.value).strip()
                    for label, value in metrics.items():
                        if label.lower() in cell_str.lower() and value is not None:
                            value_cell = ws.cell(row=row, column=col + 1)
                            value_cell.value = float(value)
                            if 'irr' in label.lower() or 'percent' in label.lower():
                                value_cell.number_format = '0.00%'
                            elif 'npv' in label.lower() or 'investment' in label.lower():
                                value_cell.number_format = '$#,##0.00'
                            elif 'score' in label.lower():
                                value_cell.number_format = '#,##0'
                            break
    
    def _populate_deal_valuation_sheet(self, ws, deal_valuation_results: Dict):
        """Populate Deal Valuation sheet."""
        if not deal_valuation_results:
            return
        
        row = 2
        if 'maximum_purchase_price' in deal_valuation_results:
            ws.cell(row=row, column=1).value = 'Maximum Purchase Price'
            ws.cell(row=row, column=2).value = float(deal_valuation_results['maximum_purchase_price'])
            ws.cell(row=row, column=2).number_format = '$#,##0.00'
            row += 1
        
        if 'actual_irr' in deal_valuation_results:
            ws.cell(row=row, column=1).value = 'Actual IRR'
            ws.cell(row=row, column=2).value = float(deal_valuation_results['actual_irr'])
            ws.cell(row=row, column=2).number_format = '0.00%'
            row += 1
    
    def _populate_monte_carlo_sheet(self, ws, mc_results: Dict):
        """Populate Monte Carlo Results sheet."""
        if not mc_results:
            return
        
        row = 2
        metrics = [
            ('Mean IRR', 'mc_mean_irr', '0.00%'),
            ('P10 IRR', 'mc_p10_irr', '0.00%'),
            ('P50 IRR', 'mc_p50_irr', '0.00%'),
            ('P90 IRR', 'mc_p90_irr', '0.00%'),
            ('Mean NPV', 'mc_mean_npv', '$#,##0.00'),
            ('P10 NPV', 'mc_p10_npv', '$#,##0.00'),
            ('P50 NPV', 'mc_p50_npv', '$#,##0.00'),
            ('P90 NPV', 'mc_p90_npv', '$#,##0.00'),
        ]
        
        for label, key, fmt in metrics:
            if key in mc_results and mc_results[key] is not None:
                ws.cell(row=row, column=1).value = label
                ws.cell(row=row, column=2).value = float(mc_results[key])
                ws.cell(row=row, column=2).number_format = fmt
                row += 1
    
    def _populate_sensitivity_sheet(self, ws, sensitivity_table: pd.DataFrame):
        """Populate Sensitivity Analysis sheet."""
        if sensitivity_table is None or sensitivity_table.empty:
            return
        
        # Clear existing data (keep headers)
        for row in range(2, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=row, column=col)
                if cell.value:
                    cell.value = None
        
        # Write table starting from row 2
        # Column headers (price multipliers)
        col_idx = 2
        for price_mult in sensitivity_table.columns:
            ws.cell(row=2, column=col_idx).value = str(price_mult)
            ws.cell(row=2, column=col_idx).font = Font(bold=True)
            col_idx += 1
        
        # Row headers and data
        row_idx = 3
        for credit_mult in sensitivity_table.index:
            ws.cell(row=row_idx, column=1).value = str(credit_mult)
            ws.cell(row=row_idx, column=1).font = Font(bold=True)
            
            col_idx = 2
            for price_mult in sensitivity_table.columns:
                irr_value = sensitivity_table.loc[credit_mult, price_mult]
                if pd.notna(irr_value):
                    ws.cell(row=row_idx, column=col_idx).value = float(irr_value)
                    ws.cell(row=row_idx, column=col_idx).number_format = '0.00%'
                col_idx += 1
            row_idx += 1
    
    def _add_presentation_charts_to_inputs(self, ws, assumptions: Dict, streaming_pct: float):
        """Generate and embed charts in Inputs & Assumptions sheet."""
        try:
            from .presentation_charts import PresentationChartGenerator
            import tempfile
            
            chart_gen = PresentationChartGenerator()
            
            # Chart 1: Assumptions Summary (E2)
            chart_path = chart_gen.create_assumptions_summary_chart(assumptions, streaming_pct)
            chart_gen.embed_chart_in_excel(chart_path, ws, 'E2', width=400, height=300)
            
            # Chart 2: Price Projection (E17)
            chart_path = chart_gen.create_price_projection_chart(assumptions, years=20)
            chart_gen.embed_chart_in_excel(chart_path, ws, 'E17', width=400, height=300)
            
            # Chart 3: Volume Projection (E34)
            chart_path = chart_gen.create_volume_projection_chart(assumptions, years=20)
            chart_gen.embed_chart_in_excel(chart_path, ws, 'E34', width=400, height=300)
            
        except Exception as e:
            print(f"Warning: Could not add charts to Inputs & Assumptions: {e}")
    
    def _add_presentation_charts_to_valuation(self, ws, valuation_schedule: pd.DataFrame):
        """Generate and embed charts in Valuation Schedule sheet."""
        try:
            from .presentation_charts import PresentationChartGenerator
            
            chart_gen = PresentationChartGenerator()
            
            # Chart 1: Cash Flow Waterfall (below data, row 25)
            chart_path = chart_gen.create_cash_flow_waterfall(valuation_schedule, years=20)
            chart_gen.embed_chart_in_excel(chart_path, ws, 'A25', width=600, height=350)
            
            # Chart 2: Cumulative Cash Flow (I25)
            chart_path = chart_gen.create_cumulative_cash_flow(valuation_schedule, years=20)
            chart_gen.embed_chart_in_excel(chart_path, ws, 'I25', width=400, height=300)
            
            # Chart 3: NPV Trend (A45)
            chart_path = chart_gen.create_npv_trend(valuation_schedule, years=20)
            chart_gen.embed_chart_in_excel(chart_path, ws, 'A45', width=600, height=350)
            
        except Exception as e:
            print(f"Warning: Could not add charts to Valuation Schedule: {e}")
    
    def _add_presentation_charts_to_summary(self, ws, actual_irr: float, target_irr: float, risk_score: Dict):
        """Generate and embed charts in Summary & Results sheet."""
        try:
            from .presentation_charts import PresentationChartGenerator
            
            chart_gen = PresentationChartGenerator()
            
            # Chart 1: Financial Metrics Dashboard (E5) - placeholder for now
            # Could add sparklines or mini charts here
            
            # Chart 2: Risk Breakdown (E15)
            if risk_score:
                chart_path = chart_gen.create_risk_breakdown(risk_score)
                chart_gen.embed_chart_in_excel(chart_path, ws, 'E15', width=400, height=300)
            
            # Chart 3: Return Summary (E30)
            chart_path = chart_gen.create_return_summary(target_irr, actual_irr)
            chart_gen.embed_chart_in_excel(chart_path, ws, 'E30', width=400, height=300)
            
        except Exception as e:
            print(f"Warning: Could not add charts to Summary & Results: {e}")
