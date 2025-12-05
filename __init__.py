"""
Carbon Model Template Package

A modular Python package for DCF and Carbon Streaming analysis.

Package Structure:
------------------
- data/: Data loading and cleaning modules
- calculators/: Financial calculation modules
- reporting/: Excel export and reporting modules
- carbon_model_generator: Main class that orchestrates all modules
"""

from .carbon_model_generator import CarbonModelGenerator

# Import submodules for direct access if needed
from .data.data_loader import DataLoader
from .calculators.dcf_calculator import DCFCalculator
from .calculators.irr_calculator import IRRCalculator
from .calculators.goal_seeker import GoalSeeker
from .calculators.sensitivity_analyzer import SensitivityAnalyzer
from .calculators.payback_calculator import PaybackCalculator
from .reporting.excel_exporter import ExcelExporter

__all__ = [
    'CarbonModelGenerator',
    'DataLoader',
    'DCFCalculator',
    'IRRCalculator',
    'GoalSeeker',
    'SensitivityAnalyzer',
    'PaybackCalculator',
    'ExcelExporter'
]

__version__ = '1.1.0'

