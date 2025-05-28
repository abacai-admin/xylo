"""
Financial analysis utility functions for calculating financial ratios and metrics.
"""
import pandas as pd
import numpy as np
from typing import Dict, List, Optional, Union, Tuple


def calculate_financial_ratios(data: pd.DataFrame) -> pd.DataFrame:
    """
    Calculate common financial ratios from financial data.
    
    Args:
        data: DataFrame containing financial metrics
        
    Returns:
        DataFrame with original data plus calculated ratios
    """
    # Make a copy to avoid modifying the original
    df = data.copy()
    
    # List of functions to apply
    ratio_functions = [
        calculate_profitability_ratios,
        calculate_liquidity_ratios,
        calculate_efficiency_ratios,
        calculate_leverage_ratios,
        calculate_valuation_ratios
    ]
    
    # Apply each set of ratio calculations
    for func in ratio_functions:
        df = func(df)
    
    return df


def calculate_profitability_ratios(df: pd.DataFrame) -> pd.DataFrame:
    """Calculate profitability ratios"""
    # Find relevant columns - handle both direct and company-specific column names
    rev_cols = [col for col in df.columns if 'TOTAL_REV' in col]
    ni_cols = [col for col in df.columns if 'NI' in col and col != 'NI_MARGIN']
    ebitda_cols = [col for col in df.columns if 'EBITDA' in col and col != 'EBITDA_MARGIN']
    ebit_cols = [col for col in df.columns if 'EBIT' in col and col != 'EBIT_MARGIN']
    asset_cols = [col for col in df.columns if 'TOTAL_ASSETS' in col]
    
    # Calculate EBITDA Margin (EBITDA / Revenue)
    for ebitda_col in ebitda_cols:
        # Find matching revenue column (for either direct or company-specific column)
        for rev_col in rev_cols:
            # Check if they match (either both are direct or both have same company suffix)
            if (('_' not in ebitda_col and '_' not in rev_col) or 
                ('_' in ebitda_col and '_' in rev_col and ebitda_col.split('_')[-1] == rev_col.split('_')[-1])):
                
                # Create appropriate suffix for the new column
                suffix = ""
                if '_' in ebitda_col:
                    suffix = f"_{ebitda_col.split('_')[-1]}"
                
                margin_col = f"EBITDA_MARGIN{suffix}"
                df[margin_col] = df[ebitda_col] / df[rev_col] * 100
    
    # Calculate Net Profit Margin (Net Income / Revenue)
    for ni_col in ni_cols:
        for rev_col in rev_cols:
            if (('_' not in ni_col and '_' not in rev_col) or 
                ('_' in ni_col and '_' in rev_col and ni_col.split('_')[-1] == rev_col.split('_')[-1])):
                
                suffix = ""
                if '_' in ni_col:
                    suffix = f"_{ni_col.split('_')[-1]}"
                
                margin_col = f"NET_PROFIT_MARGIN{suffix}"
                df[margin_col] = df[ni_col] / df[rev_col] * 100
    
    # Calculate Return on Assets (Net Income / Total Assets)
    for ni_col in ni_cols:
        for asset_col in asset_cols:
            if (('_' not in ni_col and '_' not in asset_col) or 
                ('_' in ni_col and '_' in asset_col and ni_col.split('_')[-1] == asset_col.split('_')[-1])):
                
                suffix = ""
                if '_' in ni_col:
                    suffix = f"_{ni_col.split('_')[-1]}"
                
                roa_col = f"ROA{suffix}"
                df[roa_col] = df[ni_col] / df[asset_col] * 100
    
    # Calculate EBIT Margin (EBIT / Revenue)
    for ebit_col in ebit_cols:
        for rev_col in rev_cols:
            if (('_' not in ebit_col and '_' not in rev_col) or 
                ('_' in ebit_col and '_' in rev_col and ebit_col.split('_')[-1] == rev_col.split('_')[-1])):
                
                suffix = ""
                if '_' in ebit_col:
                    suffix = f"_{ebit_col.split('_')[-1]}"
                
                margin_col = f"EBIT_MARGIN{suffix}"
                df[margin_col] = df[ebit_col] / df[rev_col] * 100
    
    return df


def calculate_liquidity_ratios(df: pd.DataFrame) -> pd.DataFrame:
    """Calculate liquidity ratios"""
    # Find relevant columns
    cash_cols = [col for col in df.columns if 'CASH_EQUIV' in col]
    asset_cols = [col for col in df.columns if 'TOTAL_ASSETS' in col]
    liab_cols = [col for col in df.columns if 'TOTAL_LIAB' in col]
    
    # Calculate Cash Ratio (Cash and Equivalents / Total Liabilities)
    for cash_col in cash_cols:
        for liab_col in liab_cols:
            if (('_' not in cash_col and '_' not in liab_col) or 
                ('_' in cash_col and '_' in liab_col and cash_col.split('_')[-1] == liab_col.split('_')[-1])):
                
                suffix = ""
                if '_' in cash_col:
                    suffix = f"_{cash_col.split('_')[-1]}"
                
                cash_ratio_col = f"CASH_RATIO{suffix}"
                df[cash_ratio_col] = df[cash_col] / df[liab_col]
    
    # Calculate Debt to Asset Ratio (Total Liabilities / Total Assets)
    for liab_col in liab_cols:
        for asset_col in asset_cols:
            if (('_' not in liab_col and '_' not in asset_col) or 
                ('_' in liab_col and '_' in asset_col and liab_col.split('_')[-1] == asset_col.split('_')[-1])):
                
                suffix = ""
                if '_' in liab_col:
                    suffix = f"_{liab_col.split('_')[-1]}"
                
                debt_asset_ratio_col = f"DEBT_TO_ASSET_RATIO{suffix}"
                df[debt_asset_ratio_col] = df[liab_col] / df[asset_col]
    
    return df


def calculate_efficiency_ratios(df: pd.DataFrame) -> pd.DataFrame:
    """Calculate efficiency ratios"""
    # We'd need more data points for most efficiency ratios, like inventory, receivables, etc.
    # This is a placeholder for future implementation
    return df


def calculate_leverage_ratios(df: pd.DataFrame) -> pd.DataFrame:
    """Calculate leverage ratios"""
    # Find relevant columns
    asset_cols = [col for col in df.columns if 'TOTAL_ASSETS' in col]
    liab_cols = [col for col in df.columns if 'TOTAL_LIAB' in col]
    
    # Calculate Debt to Equity Ratio (Total Liabilities / (Total Assets - Total Liabilities))
    for liab_col in liab_cols:
        for asset_col in asset_cols:
            if (('_' not in liab_col and '_' not in asset_col) or 
                ('_' in liab_col and '_' in asset_col and liab_col.split('_')[-1] == asset_col.split('_')[-1])):
                
                suffix = ""
                if '_' in liab_col:
                    suffix = f"_{liab_col.split('_')[-1]}"
                
                de_ratio_col = f"DEBT_TO_EQUITY_RATIO{suffix}"
                # Calculate equity as assets minus liabilities
                df[f"EQUITY{suffix}"] = df[asset_col] - df[liab_col]
                # Calculate the ratio, handling division by zero
                df[de_ratio_col] = df.apply(
                    lambda row: row[liab_col] / row[f"EQUITY{suffix}"] 
                    if row[f"EQUITY{suffix}"] != 0 else np.nan, 
                    axis=1
                )
    
    return df


def calculate_valuation_ratios(df: pd.DataFrame) -> pd.DataFrame:
    """Calculate valuation ratios"""
    # For valuation ratios, we'd need market data like market cap, enterprise value, etc.
    # If present, we can calculate metrics like P/E ratio, EV/EBITDA, etc.
    
    # We already get some of these directly (IQ_PE_RATIO, IQ_MARKETCAP)
    # Could calculate EV/EBITDA if we have enterprise value
    
    # Add more here in the future
    return df


def calculate_trend_analysis(df: pd.DataFrame, metrics: List[str], 
                            periods: int = 3) -> Dict[str, Dict[str, float]]:
    """
    Calculate trend analysis for specified metrics.
    
    Args:
        df: DataFrame with time series financial data
        metrics: List of metric columns to analyze
        periods: Number of periods to use for trend calculation
        
    Returns:
        Dictionary with trend analysis results
    """
    if 'Year' not in df.columns or df.empty:
        return {}
    
    # Sort by year
    df = df.sort_values('Year')
    
    results = {}
    for metric in metrics:
        if metric not in df.columns:
            continue
            
        # Calculate year-over-year growth rates
        df[f"{metric}_YOY"] = df[metric].pct_change() * 100
        
        # Calculate CAGR if we have enough data
        if len(df) >= 2:
            # Get first and last non-NaN values
            first_value = df[metric].first_valid_index()
            last_value = df[metric].last_valid_index()
            
            if first_value is not None and last_value is not None:
                start_val = df.loc[first_value, metric]
                end_val = df.loc[last_value, metric]
                n_years = df.loc[last_value, 'Year'] - df.loc[first_value, 'Year']
                
                if n_years > 0 and start_val > 0:
                    cagr = (end_val / start_val) ** (1 / n_years) - 1
                    cagr_pct = cagr * 100
                else:
                    cagr_pct = np.nan
            else:
                cagr_pct = np.nan
        else:
            cagr_pct = np.nan
        
        # Calculate moving average if we have enough data
        if len(df) >= periods:
            df[f"{metric}_MA{periods}"] = df[metric].rolling(window=periods).mean()
        
        # Store results
        results[metric] = {
            'latest': df[metric].iloc[-1] if not df[metric].empty else np.nan,
            'avg': df[metric].mean(),
            'min': df[metric].min(),
            'max': df[metric].max(),
            'cagr': cagr_pct,
            'recent_trend': df[f"{metric}_YOY"].iloc[-periods:].mean() if len(df) >= periods else np.nan
        }
    
    return results


def add_moving_averages(df: pd.DataFrame, columns: List[str], 
                        periods: List[int] = [3, 5]) -> pd.DataFrame:
    """
    Add moving averages for specified columns.
    
    Args:
        df: DataFrame with time series data
        columns: List of columns to calculate moving averages for
        periods: List of periods to use for moving averages
        
    Returns:
        DataFrame with moving averages added
    """
    result_df = df.copy()
    
    # Sort by year if it exists
    if 'Year' in result_df.columns:
        result_df = result_df.sort_values('Year')
    
    for col in columns:
        if col in result_df.columns:
            for period in periods:
                if len(result_df) >= period:
                    ma_col = f"{col}_MA{period}"
                    result_df[ma_col] = result_df[col].rolling(window=period).mean()
    
    return result_df
