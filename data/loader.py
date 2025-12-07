"""
Data Loader Module: Handles robust data ingestion and cleaning.

This module provides functionality to load and clean unstructured, messy time-series
data from CSV and Excel files, with intelligent column detection and standardization.
Enhanced to handle messy data with assumptions and variables.
"""

import pandas as pd
import numpy as np
from typing import Dict, List, Optional, Union
import warnings


class DataLoader:
    """
    Handles loading and cleaning of project data from various file formats.
    
    Enhanced to robustly handle messy, unstructured data with:
    - Multiple sheets in Excel files
    - Various column naming conventions
    - Missing or extra columns
    - Assumptions and variables in different formats
    - Inconsistent data types
    - Headers in non-standard positions
    """
    
    # Expected column patterns for flexible matching (expanded)
    YEAR_PATTERNS = ['year', 'time', 'period', 't', 'yr', 'y']
    CREDITS_PATTERNS = [
        'carbon_credits', 'credits_issued', 'credits', 'gross', 
        'credit_volume', 'tonnage', 'tons', 'co2', 'carbon'
    ]
    COSTS_PATTERNS = [
        'project_implementation', 'capex', 'capital', 'costs', 
        'implementation', 'expenditure', 'expenses', 'spend',
        'investment', 'project_costs'
    ]
    PRICE_PATTERNS = [
        'base_carbon_price', 'carbon_price', 'price', 'price_per_ton',
        'credit_price', 'unit_price', 'price_per_credit', 'pricing'
    ]
    
    # Standardized column names
    REQUIRED_COLUMNS = {
        'carbon_credits_gross': 'Carbon Credits Issued (Gross)',
        'project_implementation_costs': 'Project Implementation Costs',
        'base_carbon_price': 'Base Carbon Price'
    }
    
    def __init__(self, num_years: int = 20):
        """
        Initialize the DataLoader.
        
        Parameters:
        -----------
        num_years : int
            Expected number of years in the time series (default: 20)
        """
        self.num_years = num_years
    
    def detect_transposed_format(self, df: pd.DataFrame) -> bool:
        """
        Detect if data is in transposed format (labels in first column, years as columns).
        
        Parameters:
        -----------
        df : pd.DataFrame
            DataFrame to check
            
        Returns:
        --------
        bool
            True if data appears to be transposed
        """
        # Check if first column contains text labels and other columns are numeric
        if len(df.columns) < 2:
            return False
        
        first_col = df.iloc[:, 0]
        other_cols = df.iloc[:, 1:]
        
        # Check if first column is mostly text
        text_ratio = sum(1 for val in first_col if isinstance(val, str) and pd.notna(val)) / len(first_col)
        
        # Check if other columns are mostly numeric
        numeric_ratio = 0
        if len(other_cols.columns) > 0:
            sample_col = other_cols.iloc[:, 0]
            numeric_ratio = sum(1 for val in sample_col if pd.api.types.is_number(val) and pd.notna(val)) / len(sample_col)
        
        # If first column is text and other columns are numeric, likely transposed
        return text_ratio > 0.5 and numeric_ratio > 0.3
    
    def extract_data_from_transposed_format(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Extract data from transposed format where labels are in rows and years are columns.
        
        Looks for rows containing:
        - "Carbon Credits Issued" or similar
        - "Project Implementation Costs" or similar  
        - "Carbon Price" or similar
        
        Parameters:
        -----------
        df : pd.DataFrame
            Transposed DataFrame
            
        Returns:
        --------
        pd.DataFrame
            Standard format DataFrame with years as rows
        """
        result_data = {}
        
        # Find the row indices for each data type
        credits_row = None
        costs_row = None
        price_row = None
        
        # Search through all rows and columns for labels
        for idx, row in df.iterrows():
            # Check all columns for labels (not just first column)
            row_text = ' '.join([str(val).lower() for val in row if pd.notna(val)])
            
            # Look for credits row
            if credits_row is None and any(term in row_text for term in ['carbon credit', 'credits issued', 'credit issued']):
                credits_row = idx
            
            # Look for costs row
            if costs_row is None and any(term in row_text for term in ['project implementation', 'implementation cost', 'project cost']):
                costs_row = idx
            
            # Look for price row
            if price_row is None and any(term in row_text for term in ['carbon price', 'price curve', 'carbon price curve']):
                price_row = idx
        
        # Extract data from identified rows
        # Find which columns contain year data
        # Look for row with year values (row 2 seems to have years like 2063, 2064)
        year_row_idx = None
        year_cols = []
        
        # First, try to find a row with year-like values
        for idx, row in df.iterrows():
            numeric_count = 0
            year_like_count = 0
            for val in row.iloc[1:]:  # Skip first column
                try:
                    num_val = pd.to_numeric(val, errors='coerce')
                    if pd.notna(num_val):
                        numeric_count += 1
                        # Check if it looks like a year (2000-2100) or sequential (1-20)
                        if (2000 <= num_val <= 2100) or (1 <= num_val <= 20):
                            year_like_count += 1
                except:
                    pass
            
            # If most values are year-like, this is probably the year row
            if numeric_count > 5 and year_like_count > numeric_count * 0.7:
                year_row_idx = idx
                # Extract year values from this row
                for col_idx, val in enumerate(row.iloc[1:], start=1):
                    try:
                        num_val = pd.to_numeric(val, errors='coerce')
                        if pd.notna(num_val) and ((2000 <= num_val <= 2100) or (1 <= num_val <= 20)):
                            year_cols.append(df.columns[col_idx])
                    except:
                        pass
                break
        
        # If no year columns found, use all columns except first
        if not year_cols:
            year_cols = [col for col in df.columns if col != df.columns[0]]
        
        # If we found the rows, extract data
        if credits_row is not None or costs_row is not None or price_row is not None:
            # Extract data from identified rows
            # For this specific format, data starts from column 3 (index 2)
            # Skip columns 0, 1, 2 (which may have labels, units, totals)
            # Extract 20 years of data
            
            if credits_row is not None:
                credits_data = []
                # Start from column index 3, extract 20 years
                for col_idx in range(3, min(len(df.columns), 23)):
                    val = df.iloc[credits_row, col_idx]
                    try:
                        credits_data.append(float(pd.to_numeric(val, errors='coerce') or 0))
                    except:
                        credits_data.append(0.0)
                result_data['carbon_credits_gross'] = credits_data[:20]
            
            if costs_row is not None:
                costs_data = []
                for col_idx in range(3, min(len(df.columns), 23)):
                    val = df.iloc[costs_row, col_idx]
                    try:
                        costs_data.append(float(pd.to_numeric(val, errors='coerce') or 0))
                    except:
                        costs_data.append(0.0)
                result_data['project_implementation_costs'] = costs_data[:20]
            
            if price_row is not None:
                price_data = []
                for col_idx in range(3, min(len(df.columns), 23)):
                    val = df.iloc[price_row, col_idx]
                    try:
                        price_data.append(float(pd.to_numeric(val, errors='coerce') or 0))
                    except:
                        price_data.append(0.0)
                result_data['base_carbon_price'] = price_data[:20]
            
            # Create DataFrame with years as index (1 to num_years)
            if result_data:
                # Use the length of the first data series to determine number of years
                first_key = list(result_data.keys())[0]
                num_data_points = len(result_data[first_key])
                
                # Limit to num_years
                if num_data_points > self.num_years:
                    for key in result_data:
                        result_data[key] = result_data[key][:self.num_years]
                    num_data_points = self.num_years
                elif num_data_points < self.num_years:
                    # Pad with zeros
                    for key in result_data:
                        result_data[key] = list(result_data[key]) + [0.0] * (self.num_years - num_data_points)
                
                # Create years index (1 to num_years)
                years = list(range(1, self.num_years + 1))
                
                # Create DataFrame
                result_df = pd.DataFrame(result_data, index=pd.Index(years, name='Year'))
                return result_df
        
        return df
    
    def transpose_data_if_needed(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Transpose data if it's in transposed format.
        
        Parameters:
        -----------
        df : pd.DataFrame
            DataFrame that may be transposed
            
        Returns:
        --------
        pd.DataFrame
            DataFrame in standard format (years as rows)
        """
        # Always try extraction first (more robust)
        extracted = self.extract_data_from_transposed_format(df)
        # Check if extraction worked by looking for required columns
        if 'carbon_credits_gross' in extracted.columns:
            return extracted
        
        # If extraction didn't work, try detection
        if self.detect_transposed_format(df):
            # Fallback: simple transpose
            # First column becomes index (row labels)
            df_transposed = df.set_index(df.columns[0])
            # Transpose so years become rows
            df_transposed = df_transposed.T
            # Reset index to make years a column
            df_transposed = df_transposed.reset_index()
            return df_transposed
        return df
    
    def load_file(self, file_path: str, sheet_name: Optional[Union[str, int]] = None) -> pd.DataFrame:
        """
        Read a file (CSV or Excel) into a pandas DataFrame.
        
        Enhanced to handle multiple sheets and messy data, including transposed formats.
        
        Parameters:
        -----------
        file_path : str
            Path to the input file
        sheet_name : str or int, optional
            Specific sheet to read. If None, tries common sheet names.
            
        Returns:
        --------
        pd.DataFrame
            Raw DataFrame from the file
        """
        if file_path.endswith('.csv'):
            # Try different encodings for CSV
            encodings = ['utf-8', 'latin-1', 'iso-8859-1', 'cp1252']
            for encoding in encodings:
                try:
                    df = pd.read_csv(file_path, encoding=encoding)
                    if len(df.columns) > 0:
                        # Check if transposed
                        df = self.transpose_data_if_needed(df)
                        return df
                except:
                    continue
            # If all encodings fail, try default
            df = pd.read_csv(file_path)
            df = self.transpose_data_if_needed(df)
            return df
            
        elif file_path.endswith(('.xlsx', '.xls')):
            # Get all sheet names
            try:
                xl_file = pd.ExcelFile(file_path)
                sheet_names = xl_file.sheet_names
            except:
                raise ValueError(f"Could not read Excel file: {file_path}")
            
            # If sheet_name specified, use it
            if sheet_name is not None:
                try:
                    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                    df = self.transpose_data_if_needed(df)
                    return df
                except:
                    warnings.warn(f"Could not read sheet '{sheet_name}', trying alternatives")
            
            # Try common sheet names
            preferred_sheets = ['Inputs', 'Input', 'Data', 'Main', 'Project Data']
            for preferred in preferred_sheets:
                if preferred in sheet_names:
                    try:
                        df = pd.read_excel(file_path, sheet_name=preferred, header=None)
                        df = self.transpose_data_if_needed(df)
                        return df
                    except:
                        continue
            
            # Try first sheet
            try:
                df = pd.read_excel(file_path, sheet_name=0, header=None)
                df = self.transpose_data_if_needed(df)
                return df
            except:
                # If that fails, try reading with header=None to handle messy headers
                df = pd.read_excel(file_path, sheet_name=0, header=None)
                df = self.transpose_data_if_needed(df)
                return df
        else:
            raise ValueError(
                f"Unsupported file format. Expected .csv, .xlsx, or .xls, "
                f"got: {file_path.split('.')[-1]}"
            )
    
    def find_header_row(self, df: pd.DataFrame) -> int:
        """
        Find the header row in a messy DataFrame.
        
        Looks for rows that contain column-like patterns.
        
        Parameters:
        -----------
        df : pd.DataFrame
            DataFrame that may have headers in a non-standard position
            
        Returns:
        --------
        int
            Row index where headers are located (0-based)
        """
        # If DataFrame already has named columns, return 0
        if not df.columns.tolist() == list(range(len(df.columns))):
            return 0
        
        # Look for row with text-like values that could be headers
        for idx in range(min(5, len(df))):
            row = df.iloc[idx]
            # Check if row contains mostly text/string values
            text_count = sum(1 for val in row if isinstance(val, str) and len(str(val).strip()) > 0)
            if text_count >= len(row) * 0.5:  # At least 50% text
                return idx
        
        return 0
    
    def clean_headers(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Clean and standardize column headers with enhanced robustness.
        
        Parameters:
        -----------
        df : pd.DataFrame
            DataFrame with raw column names
            
        Returns:
        --------
        pd.DataFrame
            DataFrame with cleaned column names
        """
        df = df.copy()
        
        # Handle numeric column names (from header=None)
        if df.columns.dtype in [np.int64, np.int32]:
            # Try to find header row
            header_row = self.find_header_row(df)
            if header_row > 0:
                # Use that row as headers
                new_headers = df.iloc[header_row].astype(str)
                df.columns = new_headers.values
                df = df.iloc[header_row + 1:].reset_index(drop=True)
        
        # Strip whitespace, lowercase, remove special chars, replace spaces with underscores
        df.columns = df.columns.astype(str)
        df.columns = df.columns.str.strip().str.lower()
        df.columns = df.columns.str.replace(r'[^\w\s]', '', regex=True)
        df.columns = df.columns.str.replace(r'\s+', '_', regex=True)
        
        # Remove empty column names
        df.columns = [col if col and col != 'nan' else f'unnamed_{i}' 
                     for i, col in enumerate(df.columns)]
        
        return df
    
    def find_year_column(self, df: pd.DataFrame) -> Optional[str]:
        """
        Find the Year/Time column in the DataFrame with enhanced matching.
        
        Parameters:
        -----------
        df : pd.DataFrame
            DataFrame with cleaned headers
            
        Returns:
        --------
        str or None
            Name of the year column, or None if not found
        """
        # Look for columns matching year patterns
        for col in df.columns:
            col_lower = str(col).lower()
            if any(term in col_lower for term in self.YEAR_PATTERNS):
                # Check if it's numeric or can be converted
                try:
                    sample = df[col].dropna().iloc[0] if len(df[col].dropna()) > 0 else None
                    if sample is not None:
                        float(sample)  # Test if numeric
                        return col
                except:
                    # Check if it's sequential integers
                    try:
                        numeric_col = pd.to_numeric(df[col], errors='coerce')
                        if numeric_col.notna().sum() >= len(df) * 0.8:  # 80% numeric
                            return col
                    except:
                        continue
        
        # If no year column found, check if index could be years
        if df.index.name and any(term in str(df.index.name).lower() for term in self.YEAR_PATTERNS):
            return None  # Use index
        
        # If no year column found, create one if we have the right number of rows
        if len(df) == self.num_years:
            return None  # Will create year column
        elif len(df) > 0:
            # Try to infer from row count
            return None  # Will attempt to create
        
        return None
    
    def standardize_index(self, df: pd.DataFrame, year_col: Optional[str]) -> pd.DataFrame:
        """
        Standardize the DataFrame index to Year (1 to num_years).
        
        Parameters:
        -----------
        df : pd.DataFrame
            DataFrame with year column (or None)
        year_col : str or None
            Name of the year column, or None to create one
            
        Returns:
        --------
        pd.DataFrame
            DataFrame indexed by Year (1 to num_years)
        """
        df = df.copy()
        
        if year_col is None:
            # Create year column from index or row number
            if df.index.name and any(term in str(df.index.name).lower() for term in self.YEAR_PATTERNS):
                # Index might be years
                try:
                    numeric_index = pd.to_numeric(df.index, errors='coerce')
                    if numeric_index.notna().sum() >= len(df) * 0.8:
                        df.index = numeric_index
                except:
                    pass
            
            # Ensure index is 1-num_years
            if df.index.min() == 0 or (df.index.min() > 1 and df.index.max() <= self.num_years):
                if df.index.min() == 0:
                    df.index = df.index + 1
                else:
                    df.index = range(1, len(df) + 1)
            df.index.name = 'Year'
        else:
            # Use the year column
            df = df.set_index(year_col)
            df.index.name = 'Year'
            
            # Ensure index starts at 1
            if df.index.min() == 0:
                df.index = df.index + 1
            elif df.index.min() != 1:
                # Reset to 1-num_years if needed
                df = df.reset_index(drop=True)
                df.index = range(1, len(df) + 1)
                df.index.name = 'Year'
        
        # Truncate or pad to exactly num_years
        if len(df) > self.num_years:
            df = df.iloc[:self.num_years]
        elif len(df) < self.num_years:
            # Pad with NaN if needed
            new_index = pd.Index(range(1, self.num_years + 1), name='Year')
            df = df.reindex(new_index)
        
        return df
    
    def map_columns(self, df: pd.DataFrame) -> Dict[str, str]:
        """
        Create a mapping from detected columns to standardized names.
        
        Enhanced with fuzzy matching and multiple attempts.
        
        Parameters:
        -----------
        df : pd.DataFrame
            DataFrame with cleaned headers
            
        Returns:
        --------
        Dict[str, str]
            Mapping from original column names to standardized names
        """
        column_mapping = {}
        used_cols = set()
        
        # Find Carbon Credits column (try multiple patterns)
        for col in df.columns:
            if col in used_cols:
                continue
            col_lower = str(col).lower()
            if any(term in col_lower for term in self.CREDITS_PATTERNS):
                # Prefer columns with 'gross' or 'total'
                if 'gross' in col_lower or 'total' in col_lower:
                    column_mapping[col] = 'carbon_credits_gross'
                    used_cols.add(col)
                    break
                elif 'carbon' in col_lower or 'credit' in col_lower:
                    # Use this as fallback if no gross/total found
                    if 'carbon_credits_gross' not in column_mapping.values():
                        column_mapping[col] = 'carbon_credits_gross'
                        used_cols.add(col)
        
        # Find Project Implementation Costs column
        for col in df.columns:
            if col in used_cols:
                continue
            col_lower = str(col).lower()
            if any(term in col_lower for term in self.COSTS_PATTERNS):
                column_mapping[col] = 'project_implementation_costs'
                used_cols.add(col)
                break
        
        # Find Base Carbon Price column
        for col in df.columns:
            if col in used_cols:
                continue
            col_lower = str(col).lower()
            if any(term in col_lower for term in self.PRICE_PATTERNS):
                if 'base' in col_lower or 'carbon' in col_lower or 'price' in col_lower:
                    column_mapping[col] = 'base_carbon_price'
                    used_cols.add(col)
                    break
        
        return column_mapping
    
    def validate_columns(self, df: pd.DataFrame, strict: bool = False) -> List[str]:
        """
        Validate that required columns exist, with optional strict mode.
        
        Parameters:
        -----------
        df : pd.DataFrame
            DataFrame to validate
        strict : bool
            If True, raises error on missing columns. If False, returns missing list.
            
        Returns:
        --------
        List[str]
            List of missing column names
            
        Raises:
        -------
        ValueError
            If strict=True and required columns are missing
        """
        required_cols = list(self.REQUIRED_COLUMNS.keys())
        missing_cols = [col for col in required_cols if col not in df.columns]
        
        if strict and missing_cols:
            raise ValueError(
                f"Could not identify required columns: {missing_cols}. "
                f"Found columns: {list(df.columns)}. "
                f"Please ensure your data contains columns for carbon credits, "
                f"project costs, and carbon price."
            )
        
        return missing_cols
    
    def clean_numeric_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Convert required columns to numeric and handle missing values robustly.
        
        Parameters:
        -----------
        df : pd.DataFrame
            DataFrame with standardized columns
            
        Returns:
        --------
        pd.DataFrame
            DataFrame with cleaned numeric data
        """
        df = df.copy()
        required_cols = list(self.REQUIRED_COLUMNS.keys())
        
        # Convert to numeric, handling any non-numeric values
        for col in required_cols:
            if col in df.columns:
                # Try multiple conversion strategies
                try:
                    # First, try direct conversion
                    df[col] = pd.to_numeric(df[col], errors='coerce')
                except:
                    # If that fails, try replacing common non-numeric characters
                    df[col] = df[col].astype(str).str.replace(r'[^\d\.\-]', '', regex=True)
                    df[col] = pd.to_numeric(df[col], errors='coerce')
        
        # Fill NaN values with 0 for financial calculations
        # But preserve the structure
        for col in required_cols:
            if col in df.columns:
                df[col] = df[col].fillna(0)
        
        return df
    
    def extract_assumptions(self, file_path: str) -> Dict[str, any]:
        """
        Extract assumptions and variables from the data file.
        
        Looks for common assumption patterns in Excel files (separate sheets,
        named ranges, or specific column patterns). Maps to standard assumption keys.
        
        Parameters:
        -----------
        file_path : str
            Path to the input file
            
        Returns:
        --------
        Dict[str, any]
            Dictionary of extracted assumptions with standardized keys:
            - 'wacc': Discount rate
            - 'rubicon_investment_total': Total investment amount
            - 'investment_tenor': Investment deployment period
            - 'streaming_percentage_initial': Initial streaming percentage
        """
        assumptions = {}
        raw_assumptions = {}
        
        if file_path.endswith(('.xlsx', '.xls')):
            try:
                xl_file = pd.ExcelFile(file_path)
                
                # Look for assumptions sheet
                assumption_sheets = ['Assumptions', 'Assumption', 'Inputs', 'Parameters', 'Settings', 'Model Inputs']
                for sheet_name in xl_file.sheet_names:
                    if any(term.lower() in sheet_name.lower() for term in assumption_sheets):
                        try:
                            assumptions_df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                            
                            # Try multiple formats:
                            # Format 1: Two columns (Name, Value)
                            if len(assumptions_df.columns) >= 2:
                                for idx, row in assumptions_df.iterrows():
                                    key = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ""
                                    value = row.iloc[1] if len(row) > 1 else None
                                    
                                    if key and pd.notna(value):
                                        # Clean the key
                                        key = key.replace(' ', '_').replace('-', '_')
                                        raw_assumptions[key] = value
                            
                            # Format 2: Headers in first row, values in second row
                            if len(assumptions_df) >= 2:
                                headers = assumptions_df.iloc[0].astype(str).str.strip().str.lower()
                                values = assumptions_df.iloc[1]
                                for header, val in zip(headers, values):
                                    if pd.notna(header) and pd.notna(val):
                                        key = header.replace(' ', '_').replace('-', '_')
                                        raw_assumptions[key] = val
                                        
                        except Exception as e:
                            warnings.warn(f"Could not parse assumptions from sheet {sheet_name}: {e}")
                            continue
            except Exception as e:
                warnings.warn(f"Could not read Excel file for assumptions: {e}")
        
        # Map raw assumptions to standardized keys
        # Flexible matching for common variations
        key_mappings = {
            'wacc': ['wacc', 'discount_rate', 'discountrate', 'rate', 'cost_of_capital', 'costofcapital'],
            'rubicon_investment_total': [
                'rubicon_investment_total', 'investment_total', 'investmenttotal',
                'total_investment', 'totalinvestment', 'investment', 'capital_investment',
                'rubicon_investment', 'initial_investment', 'initialinvestment'
            ],
            'investment_tenor': [
                'investment_tenor', 'investmenttenor', 'tenor', 'deployment_period',
                'deploymentperiod', 'investment_period', 'investmentperiod', 'years'
            ],
            'streaming_percentage_initial': [
                'streaming_percentage_initial', 'streaming_percentage', 'streamingpercentage',
                'streaming', 'initial_streaming', 'initialstreaming', 'streaming_pct',
                'streaming_pct_initial'
            ]
        }
        
        # Try to match raw assumptions to standard keys
        for standard_key, possible_keys in key_mappings.items():
            for raw_key, raw_value in raw_assumptions.items():
                if any(possible_key in raw_key for possible_key in possible_keys):
                    try:
                        # Convert to appropriate type
                        if standard_key == 'wacc' or standard_key == 'streaming_percentage_initial':
                            # Handle percentage formats (e.g., "8%" or "0.08" or "8")
                            value_str = str(raw_value).replace('%', '').strip()
                            value = float(value_str)
                            # If value > 1, assume it's a percentage (e.g., 8 means 8%)
                            if value > 1 and (standard_key == 'wacc' or standard_key == 'streaming_percentage_initial'):
                                value = value / 100.0
                            assumptions[standard_key] = value
                        elif standard_key == 'rubicon_investment_total':
                            # Handle currency formats
                            value_str = str(raw_value).replace('$', '').replace(',', '').strip()
                            assumptions[standard_key] = float(value_str)
                        elif standard_key == 'investment_tenor':
                            assumptions[standard_key] = int(float(raw_value))
                        break
                    except (ValueError, TypeError):
                        continue
        
        return assumptions
    
    def load_data(
        self, 
        file_path: str, 
        sheet_name: Optional[Union[str, int]] = None,
        strict: bool = False
    ) -> pd.DataFrame:
        """
        Complete data loading pipeline: load, clean, and standardize data.
        
        Enhanced to handle messy, unstructured data robustly.
        
        Parameters:
        -----------
        file_path : str
            Path to the input CSV or Excel file
        sheet_name : str or int, optional
            Specific sheet to read
        strict : bool
            If True, raises error on missing required columns. If False, attempts to continue.
            
        Returns:
        --------
        pd.DataFrame
            Clean DataFrame indexed by Year (1 to num_years) with standardized columns
        """
        # Load file
        df = self.load_file(file_path, sheet_name=sheet_name)
        
        # Check if we successfully extracted from transposed format
        # (extract_data_from_transposed_format returns a properly formatted DataFrame)
        if 'carbon_credits_gross' in df.columns or 'project_implementation_costs' in df.columns:
            # Already in correct format from extraction
            # Just need to clean and validate
            missing_cols = self.validate_columns(df, strict=strict)
            if missing_cols and strict:
                raise ValueError(f"Missing required columns: {missing_cols}")
            df = self.clean_numeric_data(df)
            return df
        
        # Otherwise, proceed with standard cleaning
        # Clean headers
        df = self.clean_headers(df)
        
        # Find and standardize year column
        year_col = self.find_year_column(df)
        df = self.standardize_index(df, year_col)
        
        # Map columns to standardized names
        column_mapping = self.map_columns(df)
        df = df.rename(columns=column_mapping)
        
        # Validate required columns exist
        missing_cols = self.validate_columns(df, strict=strict)
        if missing_cols and strict:
            raise ValueError(f"Missing required columns: {missing_cols}")
        
        # Clean numeric data
        df = self.clean_numeric_data(df)
        
        return df

