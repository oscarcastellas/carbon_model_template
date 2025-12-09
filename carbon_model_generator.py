"""
CarbonModelGenerator: Main class for DCF and Carbon Streaming analysis.

This class orchestrates data loading, DCF calculations, goal-seeking,
sensitivity analysis, and Excel reporting using modular components.
"""

import pandas as pd
import warnings
from typing import Dict, Optional, List, Tuple

from .data.loader import DataLoader
from .core.dcf import DCFCalculator
from .core.irr import IRRCalculator
from .core.payback import PaybackCalculator
from .analysis.goal_seeker import GoalSeeker
from .analysis.sensitivity import SensitivityAnalyzer
from .analysis.monte_carlo import MonteCarloSimulator
from .risk.flagger import RiskFlagger
from .risk.scorer import RiskScoreCalculator
from .valuation.breakeven import BreakevenCalculator
from .export.excel import ExcelExporter


class CarbonModelGenerator:
    """
    A template class for rapid DCF and Carbon Streaming analysis.
    
    This class processes unstructured time-series project data and performs
    automated financial calculations including NPV, IRR, goal-seeking,
    sensitivity analysis, and Excel reporting.
    
    The class uses modular components for:
    - Data loading and cleaning (DataLoader)
    - DCF calculations (DCFCalculator)
    - IRR calculations (IRRCalculator)
    - Goal-seeking optimization (GoalSeeker)
    - Sensitivity analysis (SensitivityAnalyzer)
    - Payback period calculation (PaybackCalculator)
    - Excel reporting (ExcelExporter)
    """
    
    def __init__(
        self,
        wacc: Optional[float] = None,
        rubicon_investment_total: Optional[float] = None,
        investment_tenor: Optional[int] = None,
        streaming_percentage_initial: Optional[float] = None,
        num_years: int = 20,
        assumptions: Optional[Dict[str, any]] = None,
        # Monte Carlo assumptions
        price_growth_base: Optional[float] = None,
        price_growth_std_dev: Optional[float] = None,
        volume_multiplier_base: Optional[float] = None,
        volume_std_dev: Optional[float] = None
    ):
        """
        Initialize the CarbonModelGenerator with financial assumptions.
        
        Assumptions can be provided in three ways:
        1. As individual parameters
        2. As a dictionary via 'assumptions' parameter
        3. Extracted from data file later using load_data_with_assumptions()
        
        Parameters:
        -----------
        wacc : float, optional
            Discount rate (e.g., 0.08 for 8.0%). If None, must be provided later.
        rubicon_investment_total : float, optional
            Initial investment amount in USD (e.g., 20,000,000). If None, must be provided later.
        investment_tenor : int, optional
            Period over which investment is deployed in years (e.g., 5). If None, must be provided later.
        streaming_percentage_initial : float, optional
            Initial percentage of credits Rubicon receives (e.g., 0.48 for 48.0%). If None, must be provided later.
        num_years : int
            Number of years in the time series (default: 20)
        assumptions : Dict[str, any], optional
            Dictionary of assumptions with keys:
            - 'wacc': Discount rate
            - 'rubicon_investment_total': Total investment
            - 'investment_tenor': Investment period
            - 'streaming_percentage_initial': Initial streaming percentage
            - 'price_growth_base': Mean annual price growth rate (e.g., 0.03 for 3%)
            - 'price_growth_std_dev': Std dev of price growth (e.g., 0.02 for 2%)
            - 'volume_multiplier_base': Mean volume multiplier (e.g., 1.0)
            - 'volume_std_dev': Std dev of volume multiplier (e.g., 0.15 for 15%)
        price_growth_base : float, optional
            Mean annual carbon price growth rate (e.g., 0.03 for 3%)
        price_growth_std_dev : float, optional
            Standard deviation of price growth rate (e.g., 0.02 for 2%)
        volume_multiplier_base : float, optional
            Mean volume multiplier (e.g., 1.0)
        volume_std_dev : float, optional
            Standard deviation of volume multiplier (e.g., 0.15 for 15%)
        """
        # Initialize modular components (will be updated when assumptions are set)
        self.data_loader = DataLoader(num_years=num_years)
        self.irr_calculator = IRRCalculator()
        self.num_years = num_years
        
        # Store assumptions (can be None initially)
        self._wacc = wacc
        self._rubicon_investment_total = rubicon_investment_total
        self._investment_tenor = investment_tenor
        self._streaming_percentage_initial = streaming_percentage_initial
        
        # Monte Carlo assumptions
        self._price_growth_base = price_growth_base
        self._price_growth_std_dev = price_growth_std_dev
        self._volume_multiplier_base = volume_multiplier_base
        self._volume_std_dev = volume_std_dev
        
        # If assumptions dict provided, use it (overrides individual params)
        if assumptions:
            self._wacc = assumptions.get('wacc', self._wacc)
            self._rubicon_investment_total = assumptions.get('rubicon_investment_total', self._rubicon_investment_total)
            self._investment_tenor = assumptions.get('investment_tenor', self._investment_tenor)
            self._streaming_percentage_initial = assumptions.get('streaming_percentage_initial', self._streaming_percentage_initial)
            self._price_growth_base = assumptions.get('price_growth_base', self._price_growth_base)
            self._price_growth_std_dev = assumptions.get('price_growth_std_dev', self._price_growth_std_dev)
            self._volume_multiplier_base = assumptions.get('volume_multiplier_base', self._volume_multiplier_base)
            self._volume_std_dev = assumptions.get('volume_std_dev', self._volume_std_dev)
        
        # Also check individual parameters for MC assumptions
        if price_growth_base is not None:
            self._price_growth_base = price_growth_base
        if price_growth_std_dev is not None:
            self._price_growth_std_dev = price_growth_std_dev
        if volume_multiplier_base is not None:
            self._volume_multiplier_base = volume_multiplier_base
        if volume_std_dev is not None:
            self._volume_std_dev = volume_std_dev
        
        # Initialize calculators if we have all assumptions
        self._initialize_calculators()
        
        # Data storage
        self.data: Optional[pd.DataFrame] = None
        
        # Results storage
        self.dcf_results: Optional[pd.DataFrame] = None
        self.npv: Optional[float] = None
        self.irr: Optional[float] = None
        self.target_streaming_percentage: Optional[float] = None
        self.target_irr: Optional[float] = None
        self.payback_period: Optional[float] = None
        self.goal_seeker: Optional[GoalSeeker] = None
        self.monte_carlo_results: Optional[Dict] = None
        self.risk_flags: Optional[Dict] = None
        self.risk_score: Optional[Dict] = None
        self.breakeven_results: Optional[Dict] = None
    
    def _initialize_calculators(self) -> None:
        """Initialize calculator components with current assumptions."""
        if self._has_all_assumptions():
            self.dcf_calculator = DCFCalculator(
                wacc=self._wacc,
                rubicon_investment_total=self._rubicon_investment_total,
                investment_tenor=self._investment_tenor,
                irr_calculator=self.irr_calculator
            )
            self.sensitivity_analyzer = SensitivityAnalyzer(self.dcf_calculator)
            self.monte_carlo_simulator = MonteCarloSimulator(
                dcf_calculator=self.dcf_calculator,
                irr_calculator=self.irr_calculator
            )
            # Initialize productivity tools (require DCF calculator)
            self.breakeven_calculator = BreakevenCalculator(
                dcf_calculator=self.dcf_calculator,
                irr_calculator=self.irr_calculator
            )
        else:
            # Create placeholder calculators (will be recreated when assumptions are set)
            self.dcf_calculator = None
            self.sensitivity_analyzer = None
            self.monte_carlo_simulator = None
            self.breakeven_calculator = None
        
        # Initialize simple tools (don't require DCF calculator)
        self.payback_calculator = PaybackCalculator()
        self.risk_flagger = RiskFlagger()
        self.risk_score_calculator = RiskScoreCalculator()
        self.excel_exporter = ExcelExporter()
    
    def _has_all_assumptions(self) -> bool:
        """Check if all required assumptions are set."""
        return all([
            self._wacc is not None,
            self._rubicon_investment_total is not None,
            self._investment_tenor is not None,
            self._streaming_percentage_initial is not None
        ])
    
    @property
    def wacc(self) -> Optional[float]:
        """Get WACC."""
        return self._wacc
    
    @property
    def rubicon_investment_total(self) -> Optional[float]:
        """Get Rubicon investment total."""
        return self._rubicon_investment_total
    
    @property
    def investment_tenor(self) -> Optional[int]:
        """Get investment tenor."""
        return self._investment_tenor
    
    @property
    def streaming_percentage_initial(self) -> Optional[float]:
        """Get initial streaming percentage."""
        return self._streaming_percentage_initial
    
    def set_assumptions(
        self,
        wacc: Optional[float] = None,
        rubicon_investment_total: Optional[float] = None,
        investment_tenor: Optional[int] = None,
        streaming_percentage_initial: Optional[float] = None,
        price_growth_base: Optional[float] = None,
        price_growth_std_dev: Optional[float] = None,
        volume_multiplier_base: Optional[float] = None,
        volume_std_dev: Optional[float] = None,
        assumptions: Optional[Dict[str, any]] = None
    ) -> None:
        """
        Set or update financial assumptions.
        
        Parameters:
        -----------
        wacc : float, optional
            Discount rate
        rubicon_investment_total : float, optional
            Total investment amount
        investment_tenor : int, optional
            Investment deployment period
        streaming_percentage_initial : float, optional
            Initial streaming percentage
        assumptions : Dict[str, any], optional
            Dictionary of assumptions (overrides individual parameters)
        """
        if assumptions:
            self._wacc = assumptions.get('wacc', self._wacc)
            self._rubicon_investment_total = assumptions.get('rubicon_investment_total', self._rubicon_investment_total)
            self._investment_tenor = assumptions.get('investment_tenor', self._investment_tenor)
            self._streaming_percentage_initial = assumptions.get('streaming_percentage_initial', self._streaming_percentage_initial)
            self._price_growth_base = assumptions.get('price_growth_base', self._price_growth_base)
            self._price_growth_std_dev = assumptions.get('price_growth_std_dev', self._price_growth_std_dev)
            self._volume_multiplier_base = assumptions.get('volume_multiplier_base', self._volume_multiplier_base)
            self._volume_std_dev = assumptions.get('volume_std_dev', self._volume_std_dev)
        else:
            if wacc is not None:
                self._wacc = wacc
            if rubicon_investment_total is not None:
                self._rubicon_investment_total = rubicon_investment_total
            if investment_tenor is not None:
                self._investment_tenor = investment_tenor
            if streaming_percentage_initial is not None:
                self._streaming_percentage_initial = streaming_percentage_initial
            if price_growth_base is not None:
                self._price_growth_base = price_growth_base
            if price_growth_std_dev is not None:
                self._price_growth_std_dev = price_growth_std_dev
            if volume_multiplier_base is not None:
                self._volume_multiplier_base = volume_multiplier_base
            if volume_std_dev is not None:
                self._volume_std_dev = volume_std_dev
        
        # Reinitialize calculators with new assumptions
        self._initialize_calculators()
        
        # Update goal seeker if data is loaded
        if self.data is not None and self._has_all_assumptions():
            self.goal_seeker = GoalSeeker(
                dcf_calculator=self.dcf_calculator,
                data=self.data
            )
    
    def get_assumptions(self) -> Dict[str, any]:
        """
        Get current assumptions as a dictionary.
        
        Returns:
        --------
        Dict[str, any]
            Dictionary of current assumptions including Monte Carlo parameters
        """
        return {
            'wacc': self._wacc,
            'rubicon_investment_total': self._rubicon_investment_total,
            'investment_tenor': self._investment_tenor,
            'streaming_percentage_initial': self._streaming_percentage_initial,
            'price_growth_base': self._price_growth_base,
            'price_growth_std_dev': self._price_growth_std_dev,
            'volume_multiplier_base': self._volume_multiplier_base,
            'volume_std_dev': self._volume_std_dev
        }
    
    def load_data_with_assumptions(
        self,
        file_path: str,
        sheet_name: Optional[str] = None,
        use_extracted_assumptions: bool = True,
        override_assumptions: Optional[Dict[str, any]] = None
    ) -> tuple[pd.DataFrame, Dict[str, any]]:
        """
        Load data and extract assumptions from the file.
        
        Parameters:
        -----------
        file_path : str
            Path to the input CSV or Excel file
        sheet_name : str or int, optional
            Specific sheet to read (for Excel files)
        use_extracted_assumptions : bool
            If True, extract and use assumptions from the file
        override_assumptions : Dict[str, any], optional
            Assumptions to override extracted values
            
        Returns:
        --------
        Tuple[pd.DataFrame, Dict[str, any]]
            (DataFrame, Dict) - Loaded data and extracted assumptions
        """
        # Load data
        data = self.data_loader.load_data(file_path, sheet_name=sheet_name)
        self.data = data
        
        # Extract assumptions from file
        extracted_assumptions = {}
        if use_extracted_assumptions:
            extracted_assumptions = self.data_loader.extract_assumptions(file_path)
        
        # Merge with overrides (overrides take precedence)
        if override_assumptions:
            extracted_assumptions.update(override_assumptions)
        
        # Set assumptions if any were found
        if extracted_assumptions:
            self.set_assumptions(assumptions=extracted_assumptions)
        
        # Update goal seeker
        if self.data is not None and self._has_all_assumptions():
            self.goal_seeker = GoalSeeker(
                dcf_calculator=self.dcf_calculator,
                data=self.data
            )
        
        return data, extracted_assumptions
    
    def load_data(self, file_path: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
        """
        Ingest raw CSV/Excel data, clean headers, and format into a clean DataFrame.
        
        This method delegates to the DataLoader module for all data processing.
        Enhanced to handle messy, unstructured data.
        
        Note: This method does NOT extract assumptions. Use load_data_with_assumptions()
        if you want to extract assumptions from the file.
        
        Parameters:
        -----------
        file_path : str
            Path to the input CSV or Excel file
        sheet_name : str or int, optional
            Specific sheet to read (for Excel files)
            
        Returns:
        --------
        pd.DataFrame
            Clean DataFrame indexed by Year (1 to num_years) with standardized columns
        """
        if not self._has_all_assumptions():
            raise ValueError(
                "Assumptions not set. Please either:\n"
                "1. Provide assumptions in __init__()\n"
                "2. Use set_assumptions() before loading data\n"
                "3. Use load_data_with_assumptions() to extract from file"
            )
        
        self.data = self.data_loader.load_data(file_path, sheet_name=sheet_name)
        
        # Update goal seeker with new data
        if self.data is not None and self._has_all_assumptions():
            self.goal_seeker = GoalSeeker(
                dcf_calculator=self.dcf_calculator,
                data=self.data
            )
        
        return self.data
    
    def run_dcf(self, streaming_percentage: Optional[float] = None) -> Dict:
        """
        Core calculation engine for DCF analysis.
        
        This method delegates to the DCFCalculator module for all calculations.
        
        Calculates:
        - Rubicon's Share of Credits
        - Rubicon's Revenue
        - Rubicon's Net Cash Flow (Revenue - Investment deployment)
        - NPV and IRR
        
        Parameters:
        -----------
        streaming_percentage : float, optional
            Percentage of credits Rubicon receives (0.0 to 1.0).
            If None, uses self.streaming_percentage_initial
            
        Returns:
        --------
        dict
            Dictionary containing:
            - 'results_df': DataFrame with all calculated metrics
            - 'npv': Net Present Value (sum of all discounted cash flows)
            - 'irr': Internal Rate of Return (discount rate where NPV = 0)
            - 'streaming_percentage': Streaming percentage used
            - 'payback_period': Payback period in years
        """
        if not self._has_all_assumptions():
            raise ValueError(
                "Assumptions not set. Please use set_assumptions() or "
                "load_data_with_assumptions() to set assumptions."
            )
        
        if self.data is None:
            raise ValueError("Data not loaded. Please call load_data() or load_data_with_assumptions() first.")
        
        if self.dcf_calculator is None:
            raise ValueError("DCF calculator not initialized. Please set assumptions first.")
        
        if streaming_percentage is None:
            streaming_percentage = self._streaming_percentage_initial
        
        # Run DCF using the calculator module
        results = self.dcf_calculator.run_dcf(
            data=self.data,
            streaming_percentage=streaming_percentage
        )
        
        # Store results
        self.dcf_results = results['results_df']
        self.npv = results['npv']
        self.irr = results['irr']
        
        # Validate calculations
        if pd.isna(self.npv):
            raise ValueError("NPV calculation resulted in NaN. Check input data.")
        if pd.isna(self.irr):
            warnings.warn("IRR calculation resulted in NaN. Cash flows may not have a valid IRR.")
        
        # Calculate payback period
        self.payback_period = self.payback_calculator.calculate_payback_period(
            results['cash_flows']
        )
        
        # Add streaming percentage to results
        results['streaming_percentage'] = streaming_percentage
        results['payback_period'] = self.payback_period
        
        # Auto-flag risks after DCF
        self._auto_flag_risks()
        
        return results
    
    def calculate_npv(self, cash_flows: Optional[pd.Series] = None) -> float:
        """
        Calculate Net Present Value explicitly.
        
        NPV is the sum of all discounted cash flows using the WACC.
        Formula: NPV = Σ(CF_t / (1 + WACC)^(t-1))
        where CF_t is cash flow in period t, and t starts at Year 1.
        
        Parameters:
        -----------
        cash_flows : pd.Series, optional
            Cash flow series indexed by Year. If None, uses current DCF results.
            
        Returns:
        --------
        float
            Net Present Value in USD
        """
        if cash_flows is None:
            if self.dcf_results is None:
                raise ValueError("No DCF results available. Run run_dcf() first.")
            cash_flows = self.dcf_results['rubicon_net_cash_flow']
        
        # Calculate discount factors: Year 1 is not discounted (Year - 1 = 0)
        discount_factors = 1 / ((1 + self.wacc) ** (cash_flows.index - 1))
        
        # Calculate present values
        present_values = cash_flows * discount_factors
        
        # Sum to get NPV
        npv = present_values.sum()
        
        if pd.isna(npv):
            raise ValueError("NPV calculation resulted in NaN. Check input data.")
        
        return float(npv)
    
    def calculate_irr(self, cash_flows: Optional[pd.Series] = None) -> float:
        """
        Calculate Internal Rate of Return explicitly.
        
        IRR is the discount rate that makes NPV = 0.
        Uses Brent's method with fallback to fsolve for robust calculation.
        
        Parameters:
        -----------
        cash_flows : pd.Series, optional
            Cash flow series indexed by Year. If None, uses current DCF results.
            
        Returns:
        --------
        float
            Internal Rate of Return as decimal (e.g., 0.20 for 20%)
            
        Raises:
        -------
        ValueError
            If IRR cannot be calculated (e.g., all positive or all negative cash flows)
        """
        if cash_flows is None:
            if self.dcf_results is None:
                raise ValueError("No DCF results available. Run run_dcf() first.")
            cash_flows = self.dcf_results['rubicon_net_cash_flow']
        
        # Convert to numpy array (preserves order: Year 1, Year 2, ..., Year 20)
        cash_flows_array = cash_flows.values
        
        # Validate cash flows have both positive and negative values
        if (cash_flows_array >= 0).all() or (cash_flows_array <= 0).all():
            raise ValueError(
                "IRR calculation requires both positive and negative cash flows. "
                "Current cash flows are all positive or all negative."
            )
        
        # Calculate IRR
        irr = self.irr_calculator.calculate_irr(cash_flows_array)
        
        if pd.isna(irr):
            raise ValueError(
                "IRR calculation failed. Cash flows may not have a valid IRR. "
                "Ensure there are both negative (investment) and positive (return) cash flows."
            )
        
        return float(irr)
    
    def calculate_npv(self, cash_flows: Optional[pd.Series] = None) -> float:
        """
        Calculate Net Present Value explicitly.
        
        NPV is the sum of all discounted cash flows using the WACC.
        
        Parameters:
        -----------
        cash_flows : pd.Series, optional
            Cash flow series. If None, uses current DCF results.
            
        Returns:
        --------
        float
            Net Present Value in USD
        """
        if cash_flows is None:
            if self.dcf_results is None:
                raise ValueError("No DCF results available. Run run_dcf() first.")
            cash_flows = self.dcf_results['rubicon_net_cash_flow']
        
        # Calculate discount factors
        discount_factors = 1 / ((1 + self.wacc) ** (cash_flows.index - 1))
        
        # Calculate present values
        present_values = cash_flows * discount_factors
        
        # Sum to get NPV
        npv = present_values.sum()
        
        return float(npv)
    
    def calculate_irr(self, cash_flows: Optional[pd.Series] = None) -> float:
        """
        Calculate Internal Rate of Return explicitly.
        
        IRR is the discount rate that makes NPV = 0.
        
        Parameters:
        -----------
        cash_flows : pd.Series, optional
            Cash flow series. If None, uses current DCF results.
            
        Returns:
        --------
        float
            Internal Rate of Return as decimal (e.g., 0.20 for 20%)
        """
        if cash_flows is None:
            if self.dcf_results is None:
                raise ValueError("No DCF results available. Run run_dcf() first.")
            cash_flows = self.dcf_results['rubicon_net_cash_flow']
        
        # Convert to numpy array
        cash_flows_array = cash_flows.values
        
        # Calculate IRR
        irr = self.irr_calculator.calculate_irr(cash_flows_array)
        
        if pd.isna(irr):
            raise ValueError(
                "IRR calculation failed. Cash flows may not have a valid IRR. "
                "Ensure there are both negative (investment) and positive (return) cash flows."
            )
        
        return float(irr)
    
    def find_target_irr_stream(self, target_irr: float, tolerance: float = 1e-4) -> Dict:
        """
        Goal-seeking function to find streaming_percentage that achieves target IRR.
        
        This method delegates to the GoalSeeker module for optimization.
        
        Parameters:
        -----------
        target_irr : float
            Target IRR as decimal (e.g., 0.20 for 20%)
        tolerance : float
            Tolerance for convergence (default: 1e-4)
            
        Returns:
        --------
        dict
            Dictionary containing:
            - 'streaming_percentage': The calculated streaming percentage
            - 'actual_irr': The actual IRR achieved
            - 'target_irr': The target IRR
            - 'difference': The difference between actual and target
            - 'results_df': Full DCF results at the calculated streaming percentage
            - 'npv': NPV at the calculated streaming percentage
        """
        if self.data is None:
            raise ValueError("Data not loaded. Please call load_data() first.")
        
        # Update goal seeker tolerance if needed
        if self.goal_seeker is None:
            self.goal_seeker = GoalSeeker(
                dcf_calculator=self.dcf_calculator,
                data=self.data,
                tolerance=tolerance
            )
        else:
            self.goal_seeker.tolerance = tolerance
        
        # Use goal seeker to find optimal streaming percentage
        results = self.goal_seeker.find_target_irr_stream(target_irr)
        
        # Update stored results
        self.dcf_results = results['results_df']
        self.npv = results['npv']
        self.irr = results['actual_irr']
        self.target_streaming_percentage = results['streaming_percentage']
        self.target_irr = target_irr
        
        # Calculate payback period for target scenario
        self.payback_period = self.payback_calculator.calculate_payback_period(
            results['results_df']['rubicon_net_cash_flow']
        )
        results['payback_period'] = self.payback_period
        
        # Auto-flag risks after goal-seeking
        self._auto_flag_risks()
        
        return results
    
    def run_sensitivity_table(
        self,
        credit_range: List[float],
        price_range: List[float],
        streaming_percentage: Optional[float] = None
    ) -> pd.DataFrame:
        """
        Run sensitivity analysis by varying credit volumes and prices.
        
        Creates a 2D table showing IRR for each combination of:
        - Credit Volume Multipliers (rows)
        - Carbon Price Multipliers (columns)
        
        Uses the target streaming percentage from goal-seeking if available,
        otherwise uses the provided or initial streaming percentage.
        
        Parameters:
        -----------
        credit_range : List[float]
            Range of credit volume multipliers (e.g., [0.9, 1.0, 1.1])
        price_range : List[float]
            Range of carbon price multipliers (e.g., [0.8, 1.0, 1.2])
        streaming_percentage : float, optional
            Streaming percentage to use. If None, uses target_streaming_percentage
            from goal-seeking, or falls back to initial streaming percentage.
            
        Returns:
        --------
        pd.DataFrame
            2D DataFrame with credit multipliers as index, price multipliers as columns,
            and IRR values as cells
        """
        if self.data is None:
            raise ValueError("Data not loaded. Please call load_data() first.")
        
        # Determine streaming percentage to use
        if streaming_percentage is None:
            if self.target_streaming_percentage is not None:
                streaming_percentage = self.target_streaming_percentage
            else:
                streaming_percentage = self._streaming_percentage_initial
        
        # Run sensitivity analysis
        sensitivity_table = self.sensitivity_analyzer.run_sensitivity_table(
            data=self.data,
            streaming_percentage=streaming_percentage,
            credit_range=credit_range,
            price_range=price_range
        )
        
        return sensitivity_table
    
    def run_monte_carlo(
        self,
        simulations: int = 5000,
        streaming_percentage: Optional[float] = None,
        price_growth_base: Optional[float] = None,
        price_growth_std_dev: Optional[float] = None,
        volume_multiplier_base: Optional[float] = None,
        volume_std_dev: Optional[float] = None,
        random_seed: Optional[int] = None,
        use_percentage_variation: bool = False,
        use_gbm: bool = False,
        gbm_drift: Optional[float] = None,
        gbm_volatility: Optional[float] = None
    ) -> Dict:
        """
        Run Monte Carlo simulation with dual-variable stochastic modeling.
        
        This simulation uses YOUR original price forecasts and credit volumes from
        the loaded data as the center of the stochastic distribution. It then applies
        variability around those forecasts to assess probabilistic risk.
        
        The simulation builds sensitivities from YOUR original carbon price forecasts
        (e.g., from Analyst_Model_Test_OCC.xlsx), ensuring your price assumptions are
        respected while modeling uncertainty.
        
        Parameters:
        -----------
        simulations : int
            Number of simulations to run (default: 5000)
        streaming_percentage : float, optional
            Streaming percentage to use. If None, uses target_streaming_percentage
            from goal-seeking, or falls back to initial streaming percentage.
        price_growth_base : float, optional
            Mean annual price growth rate. If None, uses stored assumption.
            Used when use_percentage_variation=False.
        price_growth_std_dev : float, optional
            Standard deviation controlling price volatility around YOUR forecasts.
            If None, uses stored assumption.
        volume_multiplier_base : float, optional
            Mean volume multiplier (typically 1.0). If None, uses stored assumption.
        volume_std_dev : float, optional
            Standard deviation of volume multiplier. If None, uses stored assumption.
        random_seed : int, optional
            Random seed for reproducibility
        use_percentage_variation : bool
            If True, applies percentage multipliers directly to prices.
            If False (default), applies stochastic deviations to growth rates
            implied by your original price curve.
            
        Returns:
        --------
        Dict
            Dictionary containing:
            - 'irr_series': Array of simulated IRRs
            - 'npv_series': Array of simulated NPVs
            - 'mc_mean_irr': Mean IRR across simulations
            - 'mc_mean_npv': Mean NPV across simulations
            - 'mc_p10_irr': 10th percentile IRR (downside metric)
            - 'mc_p90_irr': 90th percentile IRR
            - 'mc_p10_npv': 10th percentile NPV
            - 'mc_p90_npv': 90th percentile NPV
            - 'mc_std_irr': Standard deviation of IRR
            - 'mc_std_npv': Standard deviation of NPV
        """
        if self.data is None:
            raise ValueError("Data not loaded. Please call load_data() or load_data_with_assumptions() first.")
        
        if self.monte_carlo_simulator is None:
            raise ValueError(
                "Monte Carlo simulator not initialized. Please ensure all "
                "base assumptions are set."
            )
        
        # Use provided parameters or fall back to stored assumptions
        if streaming_percentage is None:
            if self.target_streaming_percentage is not None:
                streaming_percentage = self.target_streaming_percentage
            else:
                streaming_percentage = self._streaming_percentage_initial
        
        price_growth_base = price_growth_base if price_growth_base is not None else self._price_growth_base
        price_growth_std_dev = price_growth_std_dev if price_growth_std_dev is not None else self._price_growth_std_dev
        volume_multiplier_base = volume_multiplier_base if volume_multiplier_base is not None else self._volume_multiplier_base
        volume_std_dev = volume_std_dev if volume_std_dev is not None else self._volume_std_dev
        
        # Validate Monte Carlo assumptions
        if price_growth_base is None or price_growth_std_dev is None:
            raise ValueError(
                "Monte Carlo price assumptions not set. Please provide "
                "price_growth_base and price_growth_std_dev."
            )
        if volume_multiplier_base is None or volume_std_dev is None:
            raise ValueError(
                "Monte Carlo volume assumptions not set. Please provide "
                "volume_multiplier_base and volume_std_dev."
            )
        
        # Run Monte Carlo simulation
        print(f"Running {simulations} Monte Carlo simulations...")
        results = self.monte_carlo_simulator.run_monte_carlo(
            base_data=self.data,
            streaming_percentage=streaming_percentage,
            price_growth_base=price_growth_base,
            price_growth_std_dev=price_growth_std_dev,
            volume_multiplier_base=volume_multiplier_base,
            volume_std_dev=volume_std_dev,
            simulations=simulations,
            random_seed=random_seed,
            use_percentage_variation=use_percentage_variation,
            use_gbm=use_gbm,
            gbm_drift=gbm_drift,
            gbm_volatility=gbm_volatility
        )
        
        # Store results
        self.monte_carlo_results = results
        
        # Re-flag risks with Monte Carlo volatility data
        self._auto_flag_risks()
        
        print(f"Monte Carlo simulation complete!")
        print(f"  Mean IRR: {results['mc_mean_irr']:.2%}")
        print(f"  P10 IRR: {results['mc_p10_irr']:.2%}")
        print(f"  P90 IRR: {results['mc_p90_irr']:.2%}")
        print(f"  Mean NPV: ${results['mc_mean_npv']:,.2f}")
        
        return results
    
    def _auto_flag_risks(self) -> None:
        """
        Automatically flag risks after DCF or goal-seeking.
        
        This is called automatically after run_dcf() and find_target_irr_stream().
        """
        if self.irr is None or self.npv is None:
            return
        
        # Get volatility from Monte Carlo if available
        irr_volatility = None
        volume_volatility = None
        price_volatility = None
        
        if self.monte_carlo_results:
            irr_series = self.monte_carlo_results.get('irr_series', [])
            if len(irr_series) > 0:
                import numpy as np
                irr_valid = [x for x in irr_series if pd.notna(x)]
                if len(irr_valid) > 0:
                    irr_volatility = np.std(irr_valid)
            
            # Extract volume and price volatility if available
            volume_volatility = self.monte_carlo_results.get('volume_volatility')
            price_volatility = self.monte_carlo_results.get('price_volatility')
        
        # Flag risks
        self.risk_flags = self.risk_flagger.flag_risks(
            irr=self.irr,
            npv=self.npv,
            payback_period=self.payback_period,
            irr_volatility=irr_volatility,
            credit_volumes=self.data['carbon_credits_gross'] if self.data is not None else None,
            project_costs=self.data['project_implementation_costs'] if self.data is not None else None
        )
        
        # Calculate risk score
        self.risk_score = self.risk_score_calculator.calculate_overall_risk_score(
            irr=self.irr,
            npv=self.npv,
            payback_period=self.payback_period,
            credit_volumes=self.data['carbon_credits_gross'] if self.data is not None else None,
            base_prices=self.data['base_carbon_price'] if self.data is not None else None,
            project_costs=self.data['project_implementation_costs'] if self.data is not None else None,
            volume_volatility=volume_volatility,
            price_volatility=price_volatility,
            total_investment=self._rubicon_investment_total
        )
    
    def flag_risks(self) -> Dict:
        """
        Manually flag risks for the current project.
        
        Returns:
        --------
        Dict
            Risk flags dictionary with risk_level, flags, red_flags, yellow_flags, green_flags
        """
        if self.irr is None or self.npv is None:
            raise ValueError("Please run DCF analysis first using run_dcf()")
        
        self._auto_flag_risks()
        return self.risk_flags
    
    def get_risk_summary(self) -> str:
        """
        Get a human-readable risk summary.
        
        Returns:
        --------
        str
            Formatted risk summary
        """
        if self.risk_flags is None:
            self.flag_risks()
        
        return self.risk_flagger.get_risk_summary(self.risk_flags)
    
    def calculate_risk_score(self) -> Dict:
        """
        Calculate overall risk score for the current project.
        
        Returns:
        --------
        Dict
            Dictionary containing:
            - 'overall_risk_score': Overall score (0-100)
            - 'financial_risk': Financial risk score
            - 'volume_risk': Volume risk score
            - 'price_risk': Price risk score
            - 'operational_risk': Operational risk score
            - 'risk_category': 'Low', 'Medium', or 'High'
        """
        if self.irr is None or self.npv is None:
            raise ValueError("Please run DCF analysis first using run_dcf()")
        
        self._auto_flag_risks()
        return self.risk_score
    
    def calculate_breakeven(
        self,
        metric: str = 'all',
        target_npv: float = 0.0
    ) -> Dict:
        """
        Calculate breakeven points for key variables.
        
        Parameters:
        -----------
        metric : str
            Which breakeven to calculate: 'price', 'volume', 'streaming', or 'all'
        target_npv : float
            Target NPV (default: 0.0 for true breakeven)
            
        Returns:
        --------
        Dict
            Dictionary with breakeven calculations
        """
        if self.breakeven_calculator is None:
            raise ValueError("Breakeven calculator not initialized. Please set assumptions first.")
        
        if self.data is None:
            raise ValueError("Please load data first using load_data()")
        
        streaming_percentage = (
            self.target_streaming_percentage 
            if self.target_streaming_percentage is not None 
            else self._streaming_percentage_initial
        )
        
        if metric == 'price':
            self.breakeven_results = {
                'breakeven_price': self.breakeven_calculator.calculate_breakeven_price(
                    self.data, streaming_percentage, target_npv
                )
            }
        elif metric == 'volume':
            self.breakeven_results = {
                'breakeven_volume': self.breakeven_calculator.calculate_breakeven_volume(
                    self.data, streaming_percentage, target_npv
                )
            }
        elif metric == 'streaming':
            self.breakeven_results = {
                'breakeven_streaming': self.breakeven_calculator.calculate_breakeven_streaming(
                    self.data, target_npv
                )
            }
        else:  # 'all'
            self.breakeven_results = self.breakeven_calculator.calculate_all_breakevens(
                self.data, streaming_percentage, target_npv
            )
        
        return self.breakeven_results
    
    def export_model_to_excel(self, filename: str) -> None:
        """
        Export complete model to formatted Excel file with multiple sheets.
        
        Creates four sheets:
        1. Inputs & Summary: Key assumptions, target streaming percentage, IRR, payback period, MC results
        2. Valuation Schedule: Detailed 20-year cash flow table
        3. Sensitivity Analysis: 2D IRR sensitivity table (if available)
        4. Monte Carlo Results: Full simulation results with histogram chart
        
        Parameters:
        -----------
        filename : str
            Output Excel filename (should end with .xlsx)
        """
        if self.dcf_results is None:
            raise ValueError(
                "No DCF results available. Please run run_dcf() or "
                "find_target_irr_stream() first."
            )
        
        # Prepare assumptions dictionary
        assumptions = {
            'wacc': self._wacc,
            'rubicon_investment_total': self._rubicon_investment_total,
            'investment_tenor': self._investment_tenor,
            'streaming_percentage_initial': self._streaming_percentage_initial,
            'price_growth_base': self._price_growth_base,
            'price_growth_std_dev': self._price_growth_std_dev,
            'volume_multiplier_base': self._volume_multiplier_base,
            'volume_std_dev': self._volume_std_dev
        }
        
        # Add GBM parameters if Monte Carlo was run
        if self.monte_carlo_results:
            assumptions['use_gbm'] = self.monte_carlo_results.get('use_gbm', False)
            assumptions['gbm_drift'] = self.monte_carlo_results.get('gbm_drift')
            assumptions['gbm_volatility'] = self.monte_carlo_results.get('gbm_volatility')
            assumptions['simulations'] = len(self.monte_carlo_results.get('irr_series', []))
        
        # Determine target streaming and IRR
        target_streaming = (
            self.target_streaming_percentage 
            if self.target_streaming_percentage is not None 
            else self._streaming_percentage_initial
        )
        target_irr = self.target_irr if self.target_irr is not None else None
        actual_irr = self.irr if self.irr is not None else 0.0
        
        # Run sensitivity analysis if we have target streaming
        sensitivity_table = None
        if self.target_streaming_percentage is not None:
            # Default sensitivity ranges
            credit_range = [0.8, 0.9, 1.0, 1.1, 1.2]
            price_range = [0.7, 0.85, 1.0, 1.15, 1.3]
            try:
                sensitivity_table = self.run_sensitivity_table(
                    credit_range=credit_range,
                    price_range=price_range,
                    streaming_percentage=target_streaming
                )
            except Exception as e:
                # If sensitivity fails, continue without it
                print(f"Warning: Could not generate sensitivity table: {e}")
        
        # Export to Excel
        self.excel_exporter.export_model_to_excel(
            filename=filename,
            assumptions=assumptions,
            target_streaming_percentage=target_streaming,
            target_irr=target_irr if target_irr is not None else actual_irr,
            actual_irr=actual_irr,
            valuation_schedule=self.dcf_results,
            sensitivity_table=sensitivity_table,
            payback_period=self.payback_period,
            monte_carlo_results=self.monte_carlo_results,
            risk_flags=self.risk_flags,
            risk_score=self.risk_score,
            breakeven_results=self.breakeven_results
        )
