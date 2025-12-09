"""
Carbon Price Volatility Visualization Module

Generates comprehensive charts and graphs for GBM price volatility analysis.
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from typing import Dict, Optional, List, Tuple
import os

# Try to import seaborn (optional)
try:
    import seaborn as sns
    sns.set_style("whitegrid")
    HAS_SEABORN = True
except ImportError:
    HAS_SEABORN = False

# Set style
plt.rcParams['figure.figsize'] = (12, 8)
plt.rcParams['font.size'] = 10
if not HAS_SEABORN:
    plt.style.use('default')


class VolatilityVisualizer:
    """
    Creates visualizations for carbon price volatility analysis.
    """
    
    def __init__(self, output_dir: str = "volatility_charts"):
        """
        Initialize the visualizer.
        
        Parameters:
        -----------
        output_dir : str
            Directory to save charts
        """
        self.output_dir = output_dir
        os.makedirs(output_dir, exist_ok=True)
    
    def plot_price_paths(
        self,
        base_prices: pd.Series,
        gbm_paths: List[pd.Series],
        title: str = "Carbon Price Volatility Paths",
        save_path: Optional[str] = None
    ) -> plt.Figure:
        """
        Plot multiple GBM price paths showing volatility.
        
        Parameters:
        -----------
        base_prices : pd.Series
            Base price forecast
        gbm_paths : List[pd.Series]
            List of GBM-generated price paths
        title : str
            Chart title
        save_path : str, optional
            Path to save figure
            
        Returns:
        --------
        plt.Figure
            Matplotlib figure
        """
        fig, ax = plt.subplots(figsize=(14, 8))
        
        # Plot base forecast
        ax.plot(base_prices.index, base_prices.values, 
                'k-', linewidth=3, label='Base Forecast', alpha=0.8)
        
        # Plot sample GBM paths (show first 50 for clarity)
        sample_paths = gbm_paths[:50]
        for i, path in enumerate(sample_paths):
            ax.plot(path.index, path.values, 
                   alpha=0.3, linewidth=0.8, color='steelblue')
        
        # Calculate and plot percentiles
        if len(gbm_paths) > 0:
            all_paths_df = pd.DataFrame(gbm_paths)
            p10 = all_paths_df.quantile(0.10, axis=0)
            p50 = all_paths_df.quantile(0.50, axis=0)
            p90 = all_paths_df.quantile(0.90, axis=0)
            
            ax.fill_between(p10.index, p10.values, p90.values,
                           alpha=0.2, color='steelblue', label='P10-P90 Range')
            ax.plot(p50.index, p50.values, '--', 
                   linewidth=2, color='darkblue', label='Median (P50)')
        
        ax.set_xlabel('Year', fontsize=12, fontweight='bold')
        ax.set_ylabel('Carbon Price ($/ton)', fontsize=12, fontweight='bold')
        ax.set_title(title, fontsize=14, fontweight='bold', pad=20)
        ax.legend(loc='best', fontsize=10)
        ax.grid(True, alpha=0.3)
        
        plt.tight_layout()
        
        if save_path:
            plt.savefig(save_path, dpi=300, bbox_inches='tight')
        
        return fig
    
    def plot_price_distribution(
        self,
        gbm_paths: List[pd.Series],
        years: List[int] = [5, 10, 15, 20],
        title: str = "Price Distribution Over Time",
        save_path: Optional[str] = None
    ) -> plt.Figure:
        """
        Plot price distribution at different time horizons.
        
        Parameters:
        -----------
        gbm_paths : List[pd.Series]
            List of GBM-generated price paths
        years : List[int]
            Years to show distributions for
        title : str
            Chart title
        save_path : str, optional
            Path to save figure
        """
        fig, axes = plt.subplots(2, 2, figsize=(16, 12))
        axes = axes.flatten()
        
        all_paths_df = pd.DataFrame(gbm_paths)
        
        for idx, year in enumerate(years):
            if year in all_paths_df.columns:
                prices_at_year = all_paths_df[year].dropna()
                
                axes[idx].hist(prices_at_year, bins=50, alpha=0.7, 
                              color='steelblue', edgecolor='black')
                axes[idx].axvline(prices_at_year.mean(), color='red', 
                                linestyle='--', linewidth=2, label=f'Mean: ${prices_at_year.mean():.2f}')
                axes[idx].axvline(prices_at_year.median(), color='green', 
                                 linestyle='--', linewidth=2, label=f'Median: ${prices_at_year.median():.2f}')
                
                axes[idx].set_xlabel('Price ($/ton)', fontsize=11, fontweight='bold')
                axes[idx].set_ylabel('Frequency', fontsize=11, fontweight='bold')
                axes[idx].set_title(f'Year {year} Price Distribution\n'
                                   f'Mean: ${prices_at_year.mean():.2f}, '
                                   f'Std: ${prices_at_year.std():.2f}',
                                   fontsize=11, fontweight='bold')
                axes[idx].legend(fontsize=9)
                axes[idx].grid(True, alpha=0.3)
        
        plt.suptitle(title, fontsize=14, fontweight='bold', y=0.995)
        plt.tight_layout()
        
        if save_path:
            plt.savefig(save_path, dpi=300, bbox_inches='tight')
        
        return fig
    
    def plot_irr_npv_distribution(
        self,
        irr_series: List[float],
        npv_series: List[float],
        title: str = "IRR and NPV Distribution from Price Volatility",
        save_path: Optional[str] = None
    ) -> plt.Figure:
        """
        Plot IRR and NPV distributions from Monte Carlo.
        
        Parameters:
        -----------
        irr_series : List[float]
            List of simulated IRRs
        npv_series : List[float]
            List of simulated NPVs
        title : str
            Chart title
        save_path : str, optional
            Path to save figure
        """
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 6))
        
        # Filter out NaN values
        irr_valid = [x for x in irr_series if not (np.isnan(x) or not np.isfinite(x))]
        npv_valid = [x for x in npv_series if not (np.isnan(x) or not np.isfinite(x))]
        
        # IRR Distribution
        ax1.hist(irr_valid, bins=50, alpha=0.7, color='steelblue', edgecolor='black')
        mean_irr = np.mean(irr_valid)
        p10_irr = np.percentile(irr_valid, 10)
        p90_irr = np.percentile(irr_valid, 90)
        
        ax1.axvline(mean_irr, color='red', linestyle='--', linewidth=2, 
                   label=f'Mean: {mean_irr:.2%}')
        ax1.axvline(p10_irr, color='orange', linestyle='--', linewidth=2, 
                   label=f'P10: {p10_irr:.2%}')
        ax1.axvline(p90_irr, color='green', linestyle='--', linewidth=2, 
                   label=f'P90: {p90_irr:.2%}')
        
        ax1.set_xlabel('IRR (%)', fontsize=12, fontweight='bold')
        ax1.set_ylabel('Frequency', fontsize=12, fontweight='bold')
        ax1.set_title('IRR Distribution\n'
                     f'Mean: {mean_irr:.2%}, Std: {np.std(irr_valid):.2%}',
                     fontsize=12, fontweight='bold')
        ax1.legend(fontsize=10)
        ax1.grid(True, alpha=0.3)
        
        # NPV Distribution
        ax2.hist(npv_valid, bins=50, alpha=0.7, color='darkgreen', edgecolor='black')
        mean_npv = np.mean(npv_valid)
        p10_npv = np.percentile(npv_valid, 10)
        p90_npv = np.percentile(npv_valid, 90)
        
        ax2.axvline(mean_npv, color='red', linestyle='--', linewidth=2, 
                   label=f'Mean: ${mean_npv:,.0f}')
        ax2.axvline(p10_npv, color='orange', linestyle='--', linewidth=2, 
                   label=f'P10: ${p10_npv:,.0f}')
        ax2.axvline(p90_npv, color='green', linestyle='--', linewidth=2, 
                   label=f'P90: ${p90_npv:,.0f}')
        
        ax2.set_xlabel('NPV ($)', fontsize=12, fontweight='bold')
        ax2.set_ylabel('Frequency', fontsize=12, fontweight='bold')
        ax2.set_title('NPV Distribution\n'
                     f'Mean: ${mean_npv:,.0f}, Std: ${np.std(npv_valid):,.0f}',
                     fontsize=12, fontweight='bold')
        ax2.legend(fontsize=10)
        ax2.grid(True, alpha=0.3)
        
        plt.suptitle(title, fontsize=14, fontweight='bold', y=1.02)
        plt.tight_layout()
        
        if save_path:
            plt.savefig(save_path, dpi=300, bbox_inches='tight')
        
        return fig
    
    def plot_volatility_impact(
        self,
        volatility_levels: List[float],
        results_by_volatility: Dict[float, Dict],
        title: str = "Impact of Price Volatility on Returns",
        save_path: Optional[str] = None
    ) -> plt.Figure:
        """
        Plot how different volatility levels affect IRR and NPV.
        
        Parameters:
        -----------
        volatility_levels : List[float]
            List of volatility levels tested
        results_by_volatility : Dict[float, Dict]
            Results for each volatility level
        title : str
            Chart title
        save_path : str, optional
            Path to save figure
        """
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 6))
        
        mean_irrs = []
        std_irrs = []
        mean_npvs = []
        std_npvs = []
        
        for vol in volatility_levels:
            if vol in results_by_volatility:
                results = results_by_volatility[vol]
                mean_irrs.append(results.get('mean_irr', 0))
                std_irrs.append(results.get('std_irr', 0))
                mean_npvs.append(results.get('mean_npv', 0))
                std_npvs.append(results.get('std_npv', 0))
        
        # IRR vs Volatility
        ax1.plot(volatility_levels, mean_irrs, 'o-', linewidth=2, 
                markersize=8, color='steelblue', label='Mean IRR')
        ax1.fill_between(volatility_levels,
                        [m - s for m, s in zip(mean_irrs, std_irrs)],
                        [m + s for m, s in zip(mean_irrs, std_irrs)],
                        alpha=0.2, color='steelblue', label='±1 Std Dev')
        
        ax1.set_xlabel('Price Volatility (σ)', fontsize=12, fontweight='bold')
        ax1.set_ylabel('IRR (%)', fontsize=12, fontweight='bold')
        ax1.set_title('IRR vs. Price Volatility', fontsize=12, fontweight='bold')
        ax1.legend(fontsize=10)
        ax1.grid(True, alpha=0.3)
        
        # NPV vs Volatility
        ax2.plot(volatility_levels, mean_npvs, 'o-', linewidth=2, 
                markersize=8, color='darkgreen', label='Mean NPV')
        ax2.fill_between(volatility_levels,
                         [m - s for m, s in zip(mean_npvs, std_npvs)],
                         [m + s for m, s in zip(mean_npvs, std_npvs)],
                         alpha=0.2, color='darkgreen', label='±1 Std Dev')
        
        ax2.set_xlabel('Price Volatility (σ)', fontsize=12, fontweight='bold')
        ax2.set_ylabel('NPV ($)', fontsize=12, fontweight='bold')
        ax2.set_title('NPV vs. Price Volatility', fontsize=12, fontweight='bold')
        ax2.legend(fontsize=10)
        ax2.grid(True, alpha=0.3)
        
        plt.suptitle(title, fontsize=14, fontweight='bold', y=1.02)
        plt.tight_layout()
        
        if save_path:
            plt.savefig(save_path, dpi=300, bbox_inches='tight')
        
        return fig
    
    def plot_price_volatility_heatmap(
        self,
        base_prices: pd.Series,
        gbm_paths: List[pd.Series],
        title: str = "Price Volatility Heatmap Over Time",
        save_path: Optional[str] = None
    ) -> plt.Figure:
        """
        Create heatmap showing price volatility distribution over time.
        
        Parameters:
        -----------
        base_prices : pd.Series
            Base price forecast
        gbm_paths : List[pd.Series]
            List of GBM-generated price paths
        title : str
            Chart title
        save_path : str, optional
            Path to save figure
        """
        fig, ax = plt.subplots(figsize=(14, 8))
        
        all_paths_df = pd.DataFrame(gbm_paths)
        
        # Calculate percentiles for each year
        percentiles = [10, 25, 50, 75, 90]
        heatmap_data = []
        
        for year in all_paths_df.columns:
            prices = all_paths_df[year].dropna()
            row = [np.percentile(prices, p) for p in percentiles]
            heatmap_data.append(row)
        
        heatmap_df = pd.DataFrame(heatmap_data, 
                                 index=all_paths_df.columns,
                                 columns=[f'P{p}' for p in percentiles])
        
        if HAS_SEABORN:
            sns.heatmap(heatmap_df.T, annot=True, fmt='.0f', cmap='RdYlGn',
                       cbar_kws={'label': 'Price ($/ton)'}, ax=ax)
        else:
            # Fallback to matplotlib imshow
            im = ax.imshow(heatmap_df.T.values, aspect='auto', cmap='RdYlGn', interpolation='nearest')
            ax.set_xticks(range(len(heatmap_df.index)))
            ax.set_xticklabels(heatmap_df.index)
            ax.set_yticks(range(len(heatmap_df.columns)))
            ax.set_yticklabels(heatmap_df.columns)
            plt.colorbar(im, ax=ax, label='Price ($/ton)')
            # Add text annotations
            for i in range(len(heatmap_df.columns)):
                for j in range(len(heatmap_df.index)):
                    text = ax.text(j, i, f'{heatmap_df.T.values[i, j]:.0f}',
                                 ha="center", va="center", color="black", fontsize=9)
        
        ax.set_xlabel('Year', fontsize=12, fontweight='bold')
        ax.set_ylabel('Percentile', fontsize=12, fontweight='bold')
        ax.set_title(title, fontsize=14, fontweight='bold', pad=20)
        
        plt.tight_layout()
        
        if save_path:
            plt.savefig(save_path, dpi=300, bbox_inches='tight')
        
        return fig
    
    def plot_correlation_analysis(
        self,
        price_paths: List[pd.Series],
        irr_series: List[float],
        npv_series: List[float],
        title: str = "Price Volatility vs. Returns Correlation",
        save_path: Optional[str] = None
    ) -> plt.Figure:
        """
        Analyze correlation between price volatility and returns.
        
        Parameters:
        -----------
        price_paths : List[pd.Series]
            List of price paths (may not match length of irr_series)
        irr_series : List[float]
            List of IRRs from Monte Carlo
        npv_series : List[float]
            List of NPVs from Monte Carlo
        title : str
            Chart title
        save_path : str, optional
            Path to save figure
        """
        fig, axes = plt.subplots(2, 2, figsize=(16, 12))
        
        # Calculate price volatility for each path (only for available paths)
        price_volatilities = []
        final_prices = []
        
        for path in price_paths:
            returns = path.pct_change().dropna()
            if len(returns) > 0:
                vol = returns.std() * np.sqrt(len(returns))  # Annualized
                price_volatilities.append(vol)
                final_prices.append(path.iloc[-1])
            else:
                price_volatilities.append(0.0)
                final_prices.append(path.iloc[0] if len(path) > 0 else 0.0)
        
        # Filter valid data - align with irr_series length
        # If price_paths is shorter, use what we have
        min_length = min(len(irr_series), len(price_volatilities))
        
        valid_data = []
        for i in range(min_length):
            if not (np.isnan(irr_series[i]) or not np.isfinite(irr_series[i])):
                if i < len(price_volatilities) and i < len(final_prices):
                    valid_data.append({
                        'vol': price_volatilities[i],
                        'final_price': final_prices[i],
                        'irr': irr_series[i],
                        'npv': npv_series[i] if i < len(npv_series) else np.nan
                    })
        
        if len(valid_data) == 0:
            # Fallback: use just IRR/NPV data without price correlation
            valid_indices = [i for i in range(len(irr_series)) 
                            if not (np.isnan(irr_series[i]) or not np.isfinite(irr_series[i]))]
            irr_valid = [irr_series[i] for i in valid_indices]
            npv_valid = [npv_series[i] for i in valid_indices if i < len(npv_series)]
            
            # Simple scatter of IRR vs NPV
            axes[0, 0].scatter(irr_valid, npv_valid, alpha=0.5, s=20)
            axes[0, 0].set_xlabel('IRR (%)', fontsize=11, fontweight='bold')
            axes[0, 0].set_ylabel('NPV ($)', fontsize=11, fontweight='bold')
            axes[0, 0].set_title('IRR vs. NPV', fontsize=11, fontweight='bold')
            axes[0, 0].grid(True, alpha=0.3)
            
            # Hide other subplots
            for ax in [axes[0, 1], axes[1, 0], axes[1, 1]]:
                ax.axis('off')
            
            plt.suptitle(title, fontsize=14, fontweight='bold', y=0.995)
            plt.tight_layout()
            if save_path:
                plt.savefig(save_path, dpi=300, bbox_inches='tight')
            return fig
        
        price_vols_valid = [d['vol'] for d in valid_data]
        final_prices_valid = [d['final_price'] for d in valid_data]
        irr_valid = [d['irr'] for d in valid_data]
        npv_valid = [d['npv'] for d in valid_data if not np.isnan(d['npv'])]
        
        # Price Volatility vs IRR
        axes[0, 0].scatter(price_vols_valid, irr_valid, alpha=0.5, s=20)
        axes[0, 0].set_xlabel('Price Volatility', fontsize=11, fontweight='bold')
        axes[0, 0].set_ylabel('IRR (%)', fontsize=11, fontweight='bold')
        axes[0, 0].set_title(f'Price Volatility vs. IRR\n'
                             f'Correlation: {np.corrcoef(price_vols_valid, irr_valid)[0,1]:.3f}',
                             fontsize=11, fontweight='bold')
        axes[0, 0].grid(True, alpha=0.3)
        
        # Price Volatility vs NPV
        axes[0, 1].scatter(price_vols_valid, npv_valid, alpha=0.5, s=20, color='green')
        axes[0, 1].set_xlabel('Price Volatility', fontsize=11, fontweight='bold')
        axes[0, 1].set_ylabel('NPV ($)', fontsize=11, fontweight='bold')
        axes[0, 1].set_title(f'Price Volatility vs. NPV\n'
                            f'Correlation: {np.corrcoef(price_vols_valid, npv_valid)[0,1]:.3f}',
                            fontsize=11, fontweight='bold')
        axes[0, 1].grid(True, alpha=0.3)
        
        # Final Price vs IRR
        axes[1, 0].scatter(final_prices_valid, irr_valid, alpha=0.5, s=20, color='orange')
        axes[1, 0].set_xlabel('Final Price ($/ton)', fontsize=11, fontweight='bold')
        axes[1, 0].set_ylabel('IRR (%)', fontsize=11, fontweight='bold')
        axes[1, 0].set_title(f'Final Price vs. IRR\n'
                             f'Correlation: {np.corrcoef(final_prices_valid, irr_valid)[0,1]:.3f}',
                             fontsize=11, fontweight='bold')
        axes[1, 0].grid(True, alpha=0.3)
        
        # Final Price vs NPV
        axes[1, 1].scatter(final_prices_valid, npv_valid, alpha=0.5, s=20, color='purple')
        axes[1, 1].set_xlabel('Final Price ($/ton)', fontsize=11, fontweight='bold')
        axes[1, 1].set_ylabel('NPV ($)', fontsize=11, fontweight='bold')
        axes[1, 1].set_title(f'Final Price vs. NPV\n'
                            f'Correlation: {np.corrcoef(final_prices_valid, npv_valid)[0,1]:.3f}',
                            fontsize=11, fontweight='bold')
        axes[1, 1].grid(True, alpha=0.3)
        
        plt.suptitle(title, fontsize=14, fontweight='bold', y=0.995)
        plt.tight_layout()
        
        if save_path:
            plt.savefig(save_path, dpi=300, bbox_inches='tight')
        
        return fig
    
    def generate_full_report(
        self,
        base_prices: pd.Series,
        gbm_paths: List[pd.Series],
        monte_carlo_results: Dict,
        output_prefix: str = "volatility_analysis"
    ) -> Dict[str, str]:
        """
        Generate complete volatility analysis report with all charts.
        
        Parameters:
        -----------
        base_prices : pd.Series
            Base price forecast
        gbm_paths : List[pd.Series]
            List of GBM-generated price paths
        monte_carlo_results : Dict
            Monte Carlo simulation results
        output_prefix : str
            Prefix for output files
            
        Returns:
        --------
        Dict[str, str]
            Dictionary mapping chart names to file paths
        """
        saved_files = {}
        
        # 1. Price Paths
        fig1 = self.plot_price_paths(
            base_prices, gbm_paths,
            title="Carbon Price Volatility: GBM Simulation Paths"
        )
        path1 = os.path.join(self.output_dir, f"{output_prefix}_price_paths.png")
        fig1.savefig(path1, dpi=300, bbox_inches='tight')
        saved_files['price_paths'] = path1
        plt.close(fig1)
        
        # 2. Price Distribution Over Time
        fig2 = self.plot_price_distribution(
            gbm_paths,
            title="Price Distribution at Key Time Horizons"
        )
        path2 = os.path.join(self.output_dir, f"{output_prefix}_price_distribution.png")
        fig2.savefig(path2, dpi=300, bbox_inches='tight')
        saved_files['price_distribution'] = path2
        plt.close(fig2)
        
        # 3. IRR and NPV Distribution
        fig3 = self.plot_irr_npv_distribution(
            monte_carlo_results.get('irr_series', []),
            monte_carlo_results.get('npv_series', []),
            title="Investment Returns Distribution from Price Volatility"
        )
        path3 = os.path.join(self.output_dir, f"{output_prefix}_returns_distribution.png")
        fig3.savefig(path3, dpi=300, bbox_inches='tight')
        saved_files['returns_distribution'] = path3
        plt.close(fig3)
        
        # 4. Volatility Heatmap
        fig4 = self.plot_price_volatility_heatmap(
            base_prices, gbm_paths,
            title="Price Volatility Distribution: Percentile Heatmap"
        )
        path4 = os.path.join(self.output_dir, f"{output_prefix}_volatility_heatmap.png")
        fig4.savefig(path4, dpi=300, bbox_inches='tight')
        saved_files['volatility_heatmap'] = path4
        plt.close(fig4)
        
        # 5. Correlation Analysis
        fig5 = self.plot_correlation_analysis(
            gbm_paths,
            monte_carlo_results.get('irr_series', []),
            monte_carlo_results.get('npv_series', []),
            title="Price Volatility Impact on Investment Returns"
        )
        path5 = os.path.join(self.output_dir, f"{output_prefix}_correlation_analysis.png")
        fig5.savefig(path5, dpi=300, bbox_inches='tight')
        saved_files['correlation_analysis'] = path5
        plt.close(fig5)
        
        print(f"\n✓ Generated {len(saved_files)} charts in '{self.output_dir}/'")
        print("\nGenerated Charts:")
        for name, path in saved_files.items():
            print(f"  • {name}: {path}")
        
        return saved_files

