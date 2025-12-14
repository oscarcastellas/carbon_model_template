"""
Simple GUI Application for Carbon Model Analysis Tool

A user-friendly interface for running carbon model analysis without
requiring any Python knowledge.
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import sys
import platform
from pathlib import Path

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from analysis_config import AnalysisConfig
from analysis.volatility_visualizer import VolatilityVisualizer
from data.loader import DataLoader
from data.multi_file_loader import MultiFileLoader
from core.dcf import DCFCalculator
from core.irr import IRRCalculator
from analysis.monte_carlo import MonteCarloSimulator
from analysis.gbm_simulator import GBMPriceSimulator
from risk.flagger import RiskFlagger
from risk.scorer import RiskScoreCalculator
from valuation.breakeven import BreakevenCalculator
from core.payback import PaybackCalculator
from export.excel import ExcelExporter
import pandas as pd
import numpy as np


class CarbonModelGUI:
    """Main GUI application for Carbon Model Analysis."""
    
    def __init__(self, root):
        """Initialize the GUI application."""
        self.root = root
        self.setup_window()
        self.create_widgets()
        
        # State variables
        self.input_files = []
        self.output_file = None
        self.is_running = False
        self.analysis_thread = None
        
    def setup_window(self):
        """Configure the main window."""
        self.root.title("Carbon Model Analysis Tool")
        self.root.geometry("700x700")
        self.root.resizable(True, True)  # Allow resizing
        self.root.minsize(650, 600)  # Minimum size
        
        # Set window icon (if available)
        try:
            # Try to set icon - will fail gracefully if icon doesn't exist
            pass
        except:
            pass
        
        # Center window on screen
        self.center_window()
        
        # Set background color
        self.root.configure(bg='#F5F5F5')
        
        # Create scrollable frame
        self.create_scrollable_frame()
        
    def create_scrollable_frame(self):
        """Create a scrollable frame for the content."""
        # Create main canvas with scrollbar
        canvas = tk.Canvas(self.root, bg='#F5F5F5', highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        self.scrollable_frame = tk.Frame(canvas, bg='#F5F5F5')
        
        def update_scrollregion(event=None):
            canvas.update_idletasks()
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        self.scrollable_frame.bind("<Configure>", update_scrollregion)
        
        canvas_window = canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Make canvas window expand with canvas
        def configure_canvas_width(event):
            canvas_width = event.width
            canvas.itemconfig(canvas_window, width=canvas_width)
        canvas.bind('<Configure>', configure_canvas_width)
        
        # Pack canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Universal scroll handler - simplified for better Mac compatibility
        def _on_scroll(event):
            """Handle scrolling for mousewheel and trackpad on all platforms."""
            try:
                # Windows mousewheel
                if hasattr(event, 'delta'):
                    delta = event.delta
                    if platform.system() == "Darwin":  # macOS
                        # Mac trackpad: delta can be positive or negative
                        # Scale appropriately for smooth scrolling
                        scroll_amount = int(-1 * delta / 3)  # Divide by 3 for smoother Mac scrolling
                        if scroll_amount == 0:
                            scroll_amount = -1 if delta < 0 else 1
                        canvas.yview_scroll(scroll_amount, "units")
                    else:  # Windows
                        canvas.yview_scroll(int(-1 * (delta / 120)), "units")
                # Linux/Mac trackpad button events
                elif hasattr(event, 'num'):
                    if event.num == 4:
                        canvas.yview_scroll(-1, "units")
                    elif event.num == 5:
                        canvas.yview_scroll(1, "units")
            except Exception:
                pass  # Silently handle any scroll errors
        
        # Bind scroll events with multiple approaches for maximum compatibility
        # Method 1: bind_all (works for most cases)
        canvas.bind_all("<MouseWheel>", _on_scroll)  # Windows & Mac
        canvas.bind_all("<Button-4>", _on_scroll)      # Linux/Mac trackpad
        canvas.bind_all("<Button-5>", _on_scroll)      # Linux/Mac trackpad
        
        # Method 2: bind directly to canvas
        canvas.bind("<MouseWheel>", _on_scroll)
        canvas.bind("<Button-4>", _on_scroll)
        canvas.bind("<Button-5>", _on_scroll)
        
        # Method 3: bind to root window (sometimes needed on Mac)
        self.root.bind("<MouseWheel>", _on_scroll)
        self.root.bind("<Button-4>", _on_scroll)
        self.root.bind("<Button-5>", _on_scroll)
        
        # Method 4: bind to scrollable frame
        self.scrollable_frame.bind("<MouseWheel>", _on_scroll)
        self.scrollable_frame.bind("<Button-4>", _on_scroll)
        self.scrollable_frame.bind("<Button-5>", _on_scroll)
        
        # Mac-specific: Ensure canvas gets focus for trackpad scrolling
        def _on_canvas_enter(event):
            """When mouse enters canvas, give it focus for better scrolling."""
            canvas.focus_set()
            # Also try to make canvas focusable
            canvas.configure(takefocus=True)
        
        canvas.bind("<Enter>", _on_canvas_enter)
        canvas.configure(takefocus=True)
        
        # Also bind to scrollable frame to ensure focus
        def _on_frame_enter(event):
            canvas.focus_set()
            canvas.configure(takefocus=True)
        
        self.scrollable_frame.bind("<Enter>", _on_frame_enter)
        
        # Store references
        self.main_canvas = canvas
        self.canvas_window = canvas_window
        
    def center_window(self):
        """Center the window on the screen."""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
    def create_widgets(self):
        """Create all GUI widgets."""
        # Use scrollable frame instead of root
        parent = self.scrollable_frame
        
        # Header
        self.create_header(parent)
        
        # Input file section
        self.create_input_section(parent)
        
        # Output file section
        self.create_output_section(parent)
        
        # Options section
        self.create_options_section(parent)
        
        # Status section
        self.create_status_section(parent)
        
        # Action buttons
        self.create_action_buttons(parent)
        
    def create_header(self, parent=None):
        """Create header section."""
        if parent is None:
            parent = self.root
        header_frame = tk.Frame(parent, bg='#366092', height=80)
        header_frame.pack(fill=tk.X, padx=0, pady=0)
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(
            header_frame,
            text="Carbon Model Analysis Tool",
            font=('Arial', 18, 'bold'),
            fg='white',
            bg='#366092'
        )
        title_label.pack(pady=10)
        
        subtitle_label = tk.Label(
            header_frame,
            text="Professional Financial Modeling for Carbon Credits",
            font=('Arial', 10),
            fg='#E0E0E0',
            bg='#366092'
        )
        subtitle_label.pack(pady=(0, 10))
        
    def create_input_section(self, parent=None):
        """Create input file selection section (supports multiple files)."""
        if parent is None:
            parent = self.root
        input_frame = tk.LabelFrame(
            parent,
            text=" 📊 Input Data Files (Excel, Word, PDF)",
            font=('Arial', 11, 'bold'),
            bg='#F5F5F5',
            fg='#333333',
            padx=15,
            pady=10
        )
        input_frame.pack(fill=tk.X, padx=20, pady=15)
        
        # File list display with scrollbar
        list_frame = tk.Frame(input_frame, bg='#F5F5F5')
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 10))
        
        # Scrollbar
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Listbox for files
        self.file_listbox = tk.Listbox(
            list_frame,
            font=('Arial', 9),
            bg='white',
            fg='#333333',
            selectmode=tk.EXTENDED,
            yscrollcommand=scrollbar.set,
            height=4
        )
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.file_listbox.yview)
        
        # Store selected files
        self.input_files = []
        
        # Button frame
        button_frame = tk.Frame(input_frame, bg='#F5F5F5')
        button_frame.pack(fill=tk.X)
        
        # Add button
        add_btn = tk.Button(
            button_frame,
            text="➕ Add Files...",
            command=self.browse_input_files,
            font=('Arial', 10, 'bold'),
            bg='#4CAF50',
            fg='black',
            activebackground='#45a049',
            activeforeground='black',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        )
        add_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # Remove button
        remove_btn = tk.Button(
            button_frame,
            text="➖ Remove Selected",
            command=self.remove_selected_files,
            font=('Arial', 10),
            bg='#f44336',
            fg='black',
            activebackground='#da190b',
            activeforeground='black',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        )
        remove_btn.pack(side=tk.LEFT)
        
        # Info label
        info_label = tk.Label(
            input_frame,
            text="Supports: Excel (.xlsx, .xls), Word (.docx, .doc), PDF (.pdf)",
            font=('Arial', 8, 'italic'),
            bg='#F5F5F5',
            fg='#666666'
        )
        info_label.pack(pady=(5, 0))
        
    def create_output_section(self, parent=None):
        """Create output file selection section."""
        if parent is None:
            parent = self.root
        output_frame = tk.LabelFrame(
            parent,
            text=" 💾 Output File (Optional)",
            font=('Arial', 11, 'bold'),
            bg='#F5F5F5',
            fg='#333333',
            padx=15,
            pady=10
        )
        output_frame.pack(fill=tk.X, padx=20, pady=10)
        
        # Default output file
        default_output = os.path.join(os.getcwd(), "results.xlsx")
        self.output_path_var = tk.StringVar(value=default_output)
        
        output_entry = tk.Entry(
            output_frame,
            textvariable=self.output_path_var,
            font=('Arial', 10),
            bg='white',
            fg='#333333'
        )
        output_entry.pack(fill=tk.X, pady=(5, 10))
        
        # Browse button
        output_browse_btn = tk.Button(
            output_frame,
            text="Browse...",
            command=self.browse_output_file,
            font=('Arial', 10),
            bg='#E0E0E0',
            fg='#333333',
            activebackground='#D0D0D0',
            relief=tk.FLAT,
            padx=20,
            pady=5,
            cursor='hand2'
        )
        output_browse_btn.pack()
        
    def create_options_section(self, parent=None):
        """Create analysis options section."""
        if parent is None:
            parent = self.root
        options_frame = tk.LabelFrame(
            parent,
            text=" ⚙️  Analysis Options",
            font=('Arial', 11, 'bold'),
            bg='#F5F5F5',
            fg='#333333',
            padx=15,
            pady=10
        )
        options_frame.pack(fill=tk.X, padx=20, pady=10)
        
        # Monte Carlo checkbox
        self.run_mc_var = tk.BooleanVar(value=True)
        mc_check = tk.Checkbutton(
            options_frame,
            text="Run Monte Carlo Simulation",
            variable=self.run_mc_var,
            font=('Arial', 10),
            bg='#F5F5F5',
            fg='#333333',
            activebackground='#F5F5F5',
            activeforeground='#333333',
            selectcolor='white'
        )
        mc_check.pack(anchor=tk.W, pady=5)
        
        # GBM checkbox
        self.use_gbm_var = tk.BooleanVar(value=True)
        gbm_check = tk.Checkbutton(
            options_frame,
            text="Use GBM (Geometric Brownian Motion)",
            variable=self.use_gbm_var,
            font=('Arial', 10),
            bg='#F5F5F5',
            fg='#333333',
            activebackground='#F5F5F5',
            activeforeground='#333333',
            selectcolor='white'
        )
        gbm_check.pack(anchor=tk.W, pady=5)
        
        # Generate charts checkbox
        self.generate_charts_var = tk.BooleanVar(value=True)
        charts_check = tk.Checkbutton(
            options_frame,
            text="Generate Charts",
            variable=self.generate_charts_var,
            font=('Arial', 10),
            bg='#F5F5F5',
            fg='#333333',
            activebackground='#F5F5F5',
            activeforeground='#333333',
            selectcolor='white'
        )
        charts_check.pack(anchor=tk.W, pady=5)
        
        # Simulations input
        sim_frame = tk.Frame(options_frame, bg='#F5F5F5')
        sim_frame.pack(fill=tk.X, pady=10)
        
        sim_label = tk.Label(
            sim_frame,
            text="Simulations:",
            font=('Arial', 10),
            bg='#F5F5F5',
            fg='#333333'
        )
        sim_label.pack(side=tk.LEFT, padx=(0, 10))
        
        self.simulations_var = tk.StringVar(value="5000")
        sim_entry = tk.Entry(
            sim_frame,
            textvariable=self.simulations_var,
            font=('Arial', 10),
            width=10,
            bg='white',
            fg='#333333'
        )
        sim_entry.pack(side=tk.LEFT)
        
    def create_status_section(self, parent=None):
        """Create status and progress section."""
        if parent is None:
            parent = self.root
        status_frame = tk.LabelFrame(
            parent,
            text=" Status",
            font=('Arial', 11, 'bold'),
            bg='#F5F5F5',
            fg='#333333',
            padx=15,
            pady=10
        )
        status_frame.pack(fill=tk.X, padx=20, pady=10)
        
        # Status label
        self.status_var = tk.StringVar(value="Ready")
        status_label = tk.Label(
            status_frame,
            textvariable=self.status_var,
            font=('Arial', 10, 'bold'),
            bg='#F5F5F5',
            fg='#4CAF50'
        )
        status_label.pack(anchor=tk.W, pady=(0, 10))
        
        # Progress bar
        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress_bar = ttk.Progressbar(
            status_frame,
            variable=self.progress_var,
            maximum=100,
            length=400,
            mode='determinate'
        )
        self.progress_bar.pack(fill=tk.X, pady=(0, 10))
        
        # Current step label
        self.current_step_var = tk.StringVar(value="Waiting to start...")
        step_label = tk.Label(
            status_frame,
            textvariable=self.current_step_var,
            font=('Arial', 9, 'italic'),
            bg='#F5F5F5',
            fg='#666666'
        )
        step_label.pack(anchor=tk.W)
        
    def create_action_buttons(self, parent=None):
        """Create action buttons."""
        if parent is None:
            parent = self.root
        button_frame = tk.Frame(parent, bg='#F5F5F5')
        button_frame.pack(fill=tk.X, padx=20, pady=20)
        
        # Run Analysis button
        self.run_btn = tk.Button(
            button_frame,
            text="▶ Run Analysis",
            command=self.run_analysis,
            font=('Arial', 12, 'bold'),
            bg='#4CAF50',
            fg='black',
            activebackground='#45a049',
            activeforeground='black',
            relief=tk.FLAT,
            padx=30,
            pady=12,
            cursor='hand2'
        )
        self.run_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Help button
        help_btn = tk.Button(
            button_frame,
            text="ℹ Help",
            command=self.show_help,
            font=('Arial', 10),
            bg='#2196F3',
            fg='black',
            activebackground='#0b7dda',
            activeforeground='black',
            relief=tk.FLAT,
            padx=20,
            pady=10,
            cursor='hand2'
        )
        help_btn.pack(side=tk.LEFT)
        
    def browse_input_files(self):
        """Open file dialog for multiple input files."""
        filenames = filedialog.askopenfilenames(
            title="Select Input Files (Excel, Word, PDF)",
            filetypes=[
                ("All supported", "*.xlsx *.xls *.docx *.doc *.pdf"),
                ("Excel files", "*.xlsx *.xls"),
                ("Word files", "*.docx *.doc"),
                ("PDF files", "*.pdf"),
                ("All files", "*.*")
            ]
        )
        
        if filenames:
            for filename in filenames:
                if filename not in self.input_files:
                    self.input_files.append(filename)
                    # Add to listbox (show just filename)
                    from pathlib import Path
                    file_name = Path(filename).name
                    self.file_listbox.insert(tk.END, file_name)
                    
    def remove_selected_files(self):
        """Remove selected files from list."""
        selected_indices = self.file_listbox.curselection()
        
        # Remove in reverse order to maintain indices
        for index in reversed(selected_indices):
            self.file_listbox.delete(index)
            del self.input_files[index]
            
    def browse_output_file(self):
        """Open file dialog for output Excel file."""
        filename = filedialog.asksaveasfilename(
            title="Save Results As",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("All files", "*.*")
            ]
        )
        
        if filename:
            self.output_file = filename
            self.output_path_var.set(filename)
            
    def validate_inputs(self):
        """Validate user inputs before running analysis."""
        if not self.input_files:
            messagebox.showerror(
                "Error",
                "Please select at least one input file (Excel, Word, or PDF)."
            )
            return False
        
        # Validate all files exist
        for file_path in self.input_files:
            if not os.path.exists(file_path):
                messagebox.showerror(
                    "Error",
                    f"File not found: {file_path}"
                )
                return False
            
        # Validate simulations
        try:
            sims = int(self.simulations_var.get())
            if sims < 100 or sims > 100000:
                messagebox.showerror(
                    "Error",
                    "Number of simulations must be between 100 and 100,000."
                )
                return False
        except ValueError:
            messagebox.showerror(
                "Error",
                "Please enter a valid number for simulations."
            )
            return False
            
        return True
        
    def run_analysis(self):
        """Run the analysis in a background thread."""
        if self.is_running:
            return
            
        if not self.validate_inputs():
            return
            
        # Disable run button
        self.run_btn.config(state=tk.DISABLED, text="Running...")
        self.is_running = True
        
        # Reset progress
        self.progress_var.set(0.0)
        self.status_var.set("Running")
        self.current_step_var.set("Initializing...")
        
        # Get output file
        output_file = self.output_path_var.get() or "results.xlsx"
        
        # Start analysis in background thread
        self.analysis_thread = threading.Thread(
            target=self._run_analysis_thread,
            args=(self.input_files, output_file),
            daemon=True
        )
        self.analysis_thread.start()
        
    def _run_analysis_thread(self, input_files, output_file):
        """Run analysis in background thread."""
        try:
            # Step 1: Load data from multiple files (10%)
            self.update_progress(10, "Loading data from multiple files...")
            
            multi_loader = MultiFileLoader()
            config = AnalysisConfig()  # Initialize config early
            
            # Load from multiple files
            extracted_assumptions = {}
            
            if len(input_files) == 1:
                # Single file - use existing loader for Excel, or multi-loader for others
                file_type = multi_loader.detect_file_type(input_files[0])
                if file_type == 'excel':
                    loader = DataLoader()
                    data = loader.load_data(input_files[0])
                    base_prices = data['base_carbon_price']
                    # Try to extract assumptions from Excel
                    try:
                        extracted_assumptions = loader.extract_assumptions(input_files[0])
                    except:
                        pass
                else:
                    # Use multi-file loader for Word/PDF
                    results = multi_loader.load_multiple_files(input_files)
                    if results['combined_data'] is not None:
                        data = results['combined_data']
                        extracted_assumptions = results.get('assumptions', {})
                        # Ensure required columns exist
                        if 'base_carbon_price' in data.columns:
                            base_prices = data['base_carbon_price']
                        else:
                            raise ValueError("Could not find carbon price data. Please ensure your file contains price information.")
                    else:
                        raise ValueError("Could not extract data table from file. Please ensure file contains a data table with Year, Credits, Price, and Costs columns.")
            else:
                # Multiple files - use multi-file loader
                results = multi_loader.load_multiple_files(input_files)
                if results['combined_data'] is not None:
                    data = results['combined_data']
                    extracted_assumptions = results.get('assumptions', {})
                    # Ensure required columns exist
                    if 'base_carbon_price' in data.columns:
                        base_prices = data['base_carbon_price']
                    else:
                        raise ValueError("Could not find carbon price data. Please ensure at least one file contains price information.")
                else:
                    raise ValueError("Could not extract data table from files. Please ensure at least one file contains a data table with Year, Credits, Price, and Costs columns.")
            
            # Apply extracted assumptions to config
            if extracted_assumptions:
                for key, value in extracted_assumptions.items():
                    if hasattr(config, key):
                        try:
                            setattr(config, key, value)
                        except:
                            pass
            
            # Validate data has required columns
            required_columns = ['carbon_credits_gross', 'base_carbon_price', 'project_implementation_costs']
            missing = [col for col in required_columns if col not in data.columns]
            if missing:
                # Try to map common column names
                column_mapping = {
                    'carbon_credits_gross': ['credits', 'carbon credits', 'gross credits', 'tonnes', 'tons'],
                    'base_carbon_price': ['price', 'carbon price', 'price per ton', 'price/ton'],
                    'project_implementation_costs': ['costs', 'project costs', 'capex', 'capital', 'implementation costs']
                }
                
                # Attempt automatic mapping
                for missing_col in missing:
                    for alt_name in column_mapping.get(missing_col, []):
                        matching_cols = [col for col in data.columns if alt_name.lower() in str(col).lower()]
                        if matching_cols:
                            data = data.rename(columns={matching_cols[0]: missing_col})
                            break
                
                # Check again
                missing = [col for col in required_columns if col not in data.columns]
                if missing:
                    raise ValueError(f"Missing required columns: {missing}. Please ensure your data has these columns or similar names.")
            
            # Ensure we have a Year column/index
            if 'Year' not in data.columns and 'year' not in [str(c).lower() for c in data.columns]:
                if data.index.name and 'year' in str(data.index.name).lower():
                    data = data.reset_index()
                elif isinstance(data.index, pd.RangeIndex):
                    # Create Year column from index
                    data['Year'] = data.index + 1
                    data = data.set_index('Year')
            
            # Get base prices
            if 'base_carbon_price' not in data.columns:
                raise ValueError("Could not find carbon price data in files.")
            base_prices = data['base_carbon_price']
            
            # Step 2: Initialize components (15%)
            self.update_progress(15, "Initializing analysis components...")
            # Update config with GUI options (assumptions may have been set from file extraction)
            config.use_gbm = self.use_gbm_var.get()
            config.gbm_drift = 0.03
            config.gbm_volatility = 0.15
            config.simulations = int(self.simulations_var.get())
            config.random_seed = 42
            
            irr_calc = IRRCalculator()
            dcf_calc = DCFCalculator(
                wacc=config.wacc,
                rubicon_investment_total=config.rubicon_investment_total,
                investment_tenor=config.investment_tenor,
                irr_calculator=irr_calc
            )
            
            # Step 3: Run DCF (25%)
            self.update_progress(25, "Running DCF analysis...")
            dcf_results = dcf_calc.run_dcf(data, config.streaming_percentage_initial)
            
            # Step 4: Calculate payback (30%)
            self.update_progress(30, "Calculating payback period...")
            payback_calc = PaybackCalculator()
            payback = payback_calc.calculate_payback_period(dcf_results['cash_flows'])
            
            # Step 5: Risk analysis (35%)
            self.update_progress(35, "Analyzing risks...")
            risk_flagger = RiskFlagger()
            risk_flags = risk_flagger.flag_risks(
                dcf_results['irr'],
                dcf_results['npv'],
                payback,
                credit_volumes=data['carbon_credits_gross'],
                project_costs=data['project_implementation_costs']
            )
            
            risk_scorer = RiskScoreCalculator()
            risk_score = risk_scorer.calculate_overall_risk_score(
                dcf_results['irr'],
                dcf_results['npv'],
                payback,
                credit_volumes=data['carbon_credits_gross'],
                base_prices=data['base_carbon_price'],
                project_costs=data['project_implementation_costs'],
                total_investment=config.rubicon_investment_total
            )
            
            # Step 6: Breakeven (40%)
            self.update_progress(40, "Calculating breakeven...")
            breakeven_calc = BreakevenCalculator(dcf_calc, irr_calc)
            breakeven = breakeven_calc.calculate_all_breakevens(
                data, config.streaming_percentage_initial, 0.0
            )
            
            # Step 6.5: Deal Valuation Back-Solver (42%)
            self.update_progress(42, "Running deal valuation back-solver...")
            deal_valuation_results = None
            try:
                from valuation.deal_valuation import DealValuationSolver
                deal_solver = DealValuationSolver(
                    dcf_calculator=dcf_calc,
                    data=data,
                    tolerance=1e-4
                )
                # Solve for purchase price given target IRR of 20%
                deal_valuation_results = deal_solver.solve_for_purchase_price(
                    target_irr=0.20,
                    streaming_percentage=config.streaming_percentage_initial,
                    investment_tenor=config.investment_tenor
                )
            except Exception as e:
                # If back-solver fails, continue without it
                print(f"Warning: Deal valuation back-solver failed: {e}")
                deal_valuation_results = None
            
            # Step 7: Monte Carlo (if enabled) (40-85%)
            mc_results = None
            if self.run_mc_var.get():
                self.update_progress(45, "Running Monte Carlo simulation...")
                mc_sim = MonteCarloSimulator(dcf_calc, irr_calc)
                
                # Update progress during MC
                def mc_progress_callback(current, total):
                    progress = 45 + int((current / total) * 40)
                    self.update_progress(
                        progress,
                        f"Monte Carlo: {current:,} of {total:,} simulations..."
                    )
                
                mc_results = mc_sim.run_monte_carlo(
                    base_data=data,
                    streaming_percentage=config.streaming_percentage_initial,
                    price_growth_base=config.price_growth_base,
                    price_growth_std_dev=config.price_growth_std_dev,
                    volume_multiplier_base=config.volume_multiplier_base,
                    volume_std_dev=config.volume_std_dev,
                    simulations=config.simulations,
                    random_seed=config.random_seed,
                    use_percentage_variation=False,
                    use_gbm=config.use_gbm,
                    gbm_drift=config.gbm_drift,
                    gbm_volatility=config.gbm_volatility
                )
                self.update_progress(85, "Monte Carlo complete!")
            else:
                self.update_progress(85, "Skipping Monte Carlo...")
            
            # Step 8: Generate charts (if enabled) (85-95%)
            saved_charts = {}
            if self.generate_charts_var.get() and config.use_gbm and mc_results:
                self.update_progress(90, "Generating charts...")
                visualizer = VolatilityVisualizer(output_dir="volatility_charts")
                
                gbm_sim = GBMPriceSimulator()
                gbm_paths = []
                for i in range(1000):
                    path = gbm_sim.generate_gbm_path_from_base(
                        base_prices=base_prices,
                        drift=config.gbm_drift,
                        volatility=config.gbm_volatility,
                        random_seed=None
                    )
                    gbm_paths.append(path)
                
                saved_charts = visualizer.generate_full_report(
                    base_prices=base_prices,
                    gbm_paths=gbm_paths,
                    monte_carlo_results=mc_results,
                    output_prefix="carbon_price_volatility"
                )
            
            # Step 9: Export to Excel (95-100%)
            self.update_progress(95, "Exporting to Excel...")
            
            # Calculate number of years from data
            num_years = len(data) if data is not None and len(data) > 0 else 20
            
            assumptions = {
                'wacc': config.wacc,
                'rubicon_investment_total': config.rubicon_investment_total,
                'investment_tenor': config.investment_tenor,
                'streaming_percentage_initial': config.streaming_percentage_initial,
                'price_growth_base': config.price_growth_base,
                'price_growth_std_dev': config.price_growth_std_dev,
                'volume_multiplier_base': config.volume_multiplier_base,
                'volume_std_dev': config.volume_std_dev,
                'use_gbm': config.use_gbm,
                'gbm_drift': config.gbm_drift,
                'gbm_volatility': config.gbm_volatility,
                'simulations': config.simulations,
                # Template customization
                'company_name': 'Investor',  # Generic default, can be customized
                'num_years': num_years  # Use actual data length
            }
            
            # Use template-based export (automatically includes interactive modules with VBA/buttons)
            excel_exporter = ExcelExporter()
            
            # Try template-based export first (includes all interactive modules)
            # This will automatically use the master template if available
            excel_exporter.export_model_to_excel(
                filename=output_file,
                assumptions=assumptions,
                target_streaming_percentage=config.streaming_percentage_initial,
                target_irr=0.20,
                actual_irr=dcf_results['irr'],
                valuation_schedule=dcf_results['results_df'],
                sensitivity_table=None,  # Can add if needed
                payback_period=payback,
                monte_carlo_results=mc_results,
                risk_flags=risk_flags,
                risk_score=risk_score,
                breakeven_results=breakeven,
                deal_valuation_results=deal_valuation_results,
                use_template=True  # Use template with interactive modules
            )
            
            # Complete!
            self.update_progress(100, "Analysis complete!")
            self.analysis_complete(True, f"Analysis complete! Results saved to:\n{output_file}")
            
        except Exception as e:
            error_msg = f"An error occurred:\n{str(e)}"
            self.analysis_complete(False, error_msg)
            
    def update_progress(self, value, text):
        """Update progress bar and status text (thread-safe)."""
        self.root.after(0, self._update_progress_ui, value, text)
        
    def _update_progress_ui(self, value, text):
        """Update UI elements (called from main thread)."""
        self.progress_var.set(value)
        self.current_step_var.set(text)
        
    def analysis_complete(self, success, message):
        """Handle analysis completion."""
        self.is_running = False
        self.run_btn.config(state=tk.NORMAL, text="▶ Run Analysis")
        
        if success:
            self.status_var.set("Complete")
            self.progress_var.set(100.0)
            
            # Ask if user wants to open results
            result = messagebox.askyesno(
                "Analysis Complete!",
                f"{message}\n\nWould you like to open the results file?",
                icon='question'
            )
            
            if result:
                # Open Excel file
                import subprocess
                import platform
                
                if platform.system() == 'Darwin':  # macOS
                    subprocess.call(['open', self.output_path_var.get() or "results.xlsx"])
                elif platform.system() == 'Windows':  # Windows
                    os.startfile(self.output_path_var.get() or "results.xlsx")
                else:  # Linux
                    subprocess.call(['xdg-open', self.output_path_var.get() or "results.xlsx"])
        else:
            self.status_var.set("Error")
            messagebox.showerror("Error", message)
            self.progress_var.set(0.0)
            self.current_step_var.set("Ready to try again...")
            
    def show_help(self):
        """Show help window."""
        help_window = tk.Toplevel(self.root)
        help_window.title("Help - Carbon Model Analysis Tool")
        help_window.geometry("500x400")
        help_window.configure(bg='white')
        
        # Center help window
        help_window.update_idletasks()
        x = (help_window.winfo_screenwidth() // 2) - (500 // 2)
        y = (help_window.winfo_screenheight() // 2) - (400 // 2)
        help_window.geometry(f'500x400+{x}+{y}')
        
        # Help text
        help_text = """
Carbon Model Analysis Tool - Help

HOW TO USE:

1. Select Input File
   - Click "Browse" button
   - Choose your Excel data file
   - File path will appear

2. (Optional) Set Output File
   - Default: results.xlsx in current folder
   - Click "Browse" to change location

3. Configure Options
   - Check/uncheck analysis options
   - Adjust number of simulations (100-100,000)

4. Run Analysis
   - Click "Run Analysis" button
   - Watch progress bar
   - Wait for completion message

5. View Results
   - Click "Yes" when asked to open results
   - Or manually open the Excel file

TIPS:

- Analysis typically takes 1-3 minutes
- More simulations = more accurate but slower
- Charts are only generated if GBM is enabled
- All results are saved in Excel format

TROUBLESHOOTING:

Problem: "No file selected"
Solution: Make sure you've selected an input Excel file

Problem: Analysis takes too long
Solution: Reduce number of simulations

Problem: Error message appears
Solution: Check that your Excel file has the required columns
        """
        
        help_label = tk.Label(
            help_window,
            text=help_text,
            font=('Arial', 10),
            bg='white',
            fg='#333333',
            justify=tk.LEFT,
            padx=20,
            pady=20
        )
        help_label.pack(fill=tk.BOTH, expand=True)
        
        # Close button
        close_btn = tk.Button(
            help_window,
            text="Close",
            command=help_window.destroy,
            font=('Arial', 10, 'bold'),
            bg='#4CAF50',
            fg='white',
            activebackground='#45a049',
            relief=tk.FLAT,
            padx=20,
            pady=5
        )
        close_btn.pack(pady=10)


def main():
    """Main entry point for GUI application."""
    root = tk.Tk()
    app = CarbonModelGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

