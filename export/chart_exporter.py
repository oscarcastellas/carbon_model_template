"""
Chart Exporter Module: Embeds volatility charts into Excel files.
"""

import os
from typing import Dict, Optional
import xlsxwriter


class ChartExporter:
    """
    Embeds volatility analysis charts into Excel files.
    """
    
    def __init__(self, workbook: xlsxwriter.Workbook):
        """
        Initialize chart exporter.
        
        Parameters:
        -----------
        workbook : xlsxwriter.Workbook
            Excel workbook to add charts to
        """
        self.workbook = workbook
    
    def add_chart_sheet(
        self,
        worksheet: xlsxwriter.Workbook.worksheet_class,
        chart_path: str,
        chart_name: str,
        row: int = 0,
        col: int = 0,
        width: int = 1200,
        height: int = 700
    ) -> int:
        """
        Add a chart image to Excel worksheet.
        
        Parameters:
        -----------
        worksheet : xlsxwriter worksheet
            Worksheet to add chart to
        chart_path : str
            Path to chart image file
        chart_name : str
            Name/description of chart
        row : int
            Starting row
        col : int
            Starting column
        width : int
            Image width in pixels
        height : int
            Image height in pixels
            
        Returns:
        --------
        int
            Next available row after chart
        """
        if not os.path.exists(chart_path):
            worksheet.write(row, col, f"Chart not found: {chart_name}", 
                          self.workbook.add_format({'bold': True, 'font_color': 'red'}))
            return row + 2
        
        # Write chart title
        title_format = self.workbook.add_format({
            'bold': True,
            'font_size': 14,
            'bg_color': '#366092',
            'font_color': 'white',
            'align': 'left',
            'valign': 'vcenter',
            'border': 1
        })
        worksheet.write(row, col, chart_name, title_format)
        worksheet.set_row(row, 30)
        row += 1
        
        # Insert image
        worksheet.insert_image(row, col, chart_path, {
            'x_scale': width / 1200,  # Scale to desired width
            'y_scale': height / 700,  # Scale to desired height
            'x_offset': 10,
            'y_offset': 10
        })
        
        # Calculate rows used (approximate: 1 row per 20 pixels)
        rows_used = int(height / 20) + 5
        row += rows_used
        
        return row
    
    def create_charts_sheet(
        self,
        charts: Dict[str, str],
        sheet_name: str = "Volatility Charts"
    ) -> xlsxwriter.Workbook.worksheet_class:
        """
        Create a new sheet with all volatility charts.
        
        Parameters:
        -----------
        charts : Dict[str, str]
            Dictionary mapping chart names to file paths
        sheet_name : str
            Name of the sheet
            
        Returns:
        --------
        xlsxwriter worksheet
            Created worksheet
        """
        worksheet = self.workbook.add_worksheet(sheet_name)
        
        # Set column width
        worksheet.set_column(0, 0, 100)
        
        row = 0
        
        # Title
        title_format = self.workbook.add_format({
            'bold': True,
            'font_size': 16,
            'bg_color': '#366092',
            'font_color': 'white',
            'align': 'center'
        })
        worksheet.merge_range(row, 0, row, 2, 
                            'Carbon Price Volatility Analysis - Charts', title_format)
        worksheet.set_row(row, 40)
        row += 2
        
        # Add each chart
        chart_order = [
            ('price_paths', '1. Price Volatility Paths'),
            ('price_distribution', '2. Price Distribution Over Time'),
            ('returns_distribution', '3. Returns Distribution (IRR & NPV)'),
            ('volatility_heatmap', '4. Volatility Heatmap'),
            ('correlation_analysis', '5. Correlation Analysis')
        ]
        
        for chart_key, chart_title in chart_order:
            if chart_key in charts:
                row = self.add_chart_sheet(
                    worksheet=worksheet,
                    chart_path=charts[chart_key],
                    chart_name=chart_title,
                    row=row,
                    col=0,
                    width=1400,
                    height=800
                )
                row += 3  # Space between charts
        
        return worksheet

