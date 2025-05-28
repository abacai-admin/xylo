import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date
import os
import tempfile
import numpy as np
from logic.api_handler import fetch_company_by_ticker
from logic.pptx_generator import generate_presentation
from logic.financial_analysis import calculate_financial_ratios, calculate_trend_analysis, add_moving_averages

# Set page config for better layout
st.set_page_config(
    page_title="Financial Slide Builder",
    page_icon="📊",
    layout="wide"
)

# Initialize session state for slides if it doesn't exist
if 'slides' not in st.session_state:
    st.session_state.slides = []

# Initialize session state for slide counter
if 'slide_counter' not in st.session_state:
    st.session_state.slide_counter = 0

def initialize_slide():
    """Initialize a new slide with default values"""
    current_year = date.today().year
    return {
        "id": f"slide_{st.session_state.slide_counter}",
        "title": "",
        "ticker": "",
        "ticker2": "",  # Added second ticker for comparison
        "years": 5,
        "metrics": [],
        "chart_type": "table",
        "chart_data": None,
        "chart_data2": None,  # Added second dataset for comparison
        "selected_columns": [],
        "is_comparison": False,  # Flag to indicate if this is a comparison slide
        "enable_ratios": True,  # Enable financial ratio calculations
        "enable_trend_analysis": False,  # Enable trend analysis
        "moving_average_periods": [3],  # Periods for moving averages
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

def add_slide():
    """Add a new slide to the presentation"""
    st.session_state.slides.append(initialize_slide())
    st.session_state.slide_counter += 1
    st.rerun()

def remove_slide(slide_id):
    """Remove a slide from the presentation"""
    st.session_state.slides = [s for s in st.session_state.slides if s['id'] != slide_id]

def add_bullet_point(slide_idx):
    """Add a bullet point to a slide"""
    if 'new_bullet' in st.session_state and st.session_state.new_bullet:
        st.session_state.slides[slide_idx]['content'].append(st.session_state.new_bullet)
        st.session_state.new_bullet = ""

def display_company_metrics(slide_idx, ticker, years, is_second_company=False):
    """Fetch and display company metrics based on ticker and years"""
    if not ticker:
        return None
    
    try:
        # Fetch data from CIQ API
        with st.spinner(f"Fetching data for {ticker}..."):
            data = fetch_company_by_ticker(ticker, years)
            
        if not data.empty:
            # Check if ratios should be calculated
            if st.session_state.slides[slide_idx]['enable_ratios']:
                with st.spinner(f"Calculating financial ratios for {ticker}..."):
                    data = calculate_financial_ratios(data)
            
            # Check if trend analysis should be performed
            if st.session_state.slides[slide_idx]['enable_trend_analysis']:
                # Only store the trend analysis results (doesn't modify data)
                numeric_cols = data.select_dtypes(include=['float64', 'int64']).columns.tolist()
                if 'Year' in numeric_cols:
                    numeric_cols.remove('Year')
                
                # Store trend analysis for potential display
                trend_results = calculate_trend_analysis(
                    data, 
                    numeric_cols,
                    periods=3
                )
                
                # Store trend analysis in session state
                if is_second_company:
                    st.session_state.slides[slide_idx]['trend_analysis2'] = trend_results
                else:
                    st.session_state.slides[slide_idx]['trend_analysis'] = trend_results
            
            # Add moving averages if configured
            if st.session_state.slides[slide_idx].get('moving_average_periods'):
                numeric_cols = data.select_dtypes(include=['float64', 'int64']).columns.tolist()
                if 'Year' in numeric_cols:
                    numeric_cols.remove('Year')
                
                data = add_moving_averages(
                    data,
                    numeric_cols,
                    st.session_state.slides[slide_idx]['moving_average_periods']
                )
            
            # Store the processed data
            if is_second_company:
                st.session_state.slides[slide_idx]['chart_data2'] = data.to_dict()
                st.session_state.slides[slide_idx]['is_comparison'] = True
            else:
                st.session_state.slides[slide_idx]['chart_data'] = data.to_dict()
            
            return data
        else:
            st.warning(f"No data available for {ticker} for the selected time period.")
            return None
            
    except Exception as e:
        st.error(f"Error fetching data for {ticker}: {str(e)}")
        return None
        
def merge_company_data(data1, data2, ticker1, ticker2):
    """Merge data from two companies for comparison"""
    if data1 is None or data2 is None:
        return None
        
    try:
        # Create copies to avoid modifying the original data
        df1 = data1.copy()
        df2 = data2.copy()
        
        # First, make sure both dataframes have a Year column for merging
        if 'Year' not in df1.columns and 'Date' in df1.columns:
            df1['Year'] = pd.to_datetime(df1['Date']).dt.year
        if 'Year' not in df2.columns and 'Date' in df2.columns:
            df2['Year'] = pd.to_datetime(df2['Date']).dt.year
            
        # Add company identifier to column names for distinction (excluding Year and Date)
        df1_cols = {}
        df2_cols = {}
        
        for col in df1.columns:
            if col not in ['Year', 'Date']:
                df1_cols[col] = f"{col}_{ticker1}"
            else:
                df1_cols[col] = col
                
        for col in df2.columns:
            if col not in ['Year', 'Date']:
                df2_cols[col] = f"{col}_{ticker2}"
            else:
                df2_cols[col] = col
        
        # Rename columns with company identifiers
        df1 = df1.rename(columns=df1_cols)
        df2 = df2.rename(columns=df2_cols)
            
        # Merge the dataframes on Year
        merged_df = pd.merge(df1, df2, on='Year', how='outer')
        
        return merged_df
    except Exception as e:
        st.error(f"Error merging company data: {str(e)}")
        return None

def render_chart(chart_type, data, ticker, slide_id=None, is_comparison=False, second_ticker=None):
    """Render the appropriate chart based on chart type and return selected columns"""
    if data is None or data.empty:
        return None, []
        
    try:
        # Create a copy of the data to manipulate
        df = data.copy()
        
        # If 'Date' column exists, convert to just the year for plotting
        if 'Date' in df.columns:
            df['Year'] = pd.to_datetime(df['Date']).dt.year
        
        # Select numeric columns for plotting
        numeric_cols = df.select_dtypes(include=['float64', 'int64']).columns.tolist()
        if 'Year' in numeric_cols:
            numeric_cols.remove('Year')
        
        # Remove ticker and company metadata columns
        for col in ['Ticker', 'Company']:
            if col in numeric_cols:
                numeric_cols.remove(col)
        
        # For comparison data, filter out the company-specific columns for current metrics
        if is_comparison and second_ticker:
            # Create a list of the base metric names without company suffixes
            base_metrics = set()
            for col in numeric_cols:
                # Split by underscore to get the base name and company suffix
                parts = col.split('_')
                if len(parts) > 1 and (parts[-1] == ticker or parts[-1] == second_ticker):
                    # The base name is everything except the last part (company suffix)
                    base_name = '_'.join(parts[:-1])
                    base_metrics.add(base_name)
                elif col not in ['Year', 'Date', 'Ticker', 'Company']:
                    # If no company suffix, add the column as is
                    base_metrics.add(col)
            
            # Filter for metrics that exist for both companies
            available_metrics = []
            for metric in base_metrics:
                col1 = f"{metric}_{ticker}"
                col2 = f"{metric}_{second_ticker}"
                
                # If either direct match or prefixed version exists for both companies
                if (col1 in df.columns or any(c.startswith(f"{metric}_") and c.endswith(f"_{ticker}") for c in df.columns)) and \
                   (col2 in df.columns or any(c.startswith(f"{metric}_") and c.endswith(f"_{second_ticker}") for c in df.columns)):
                    available_metrics.append(metric)
            
            # Use these base metrics for selection instead of the raw columns
            numeric_cols = sorted(available_metrics) if available_metrics else numeric_cols
        
        # Filter columns to include ratio metrics
        ratio_cols = [col for col in numeric_cols if any(ratio in col for ratio in ['MARGIN', 'RATIO', 'ROA', 'ROE'])]
        financial_cols = [col for col in numeric_cols if col not in ratio_cols]
        
        # Find columns with moving averages
        ma_cols = [col for col in numeric_cols if 'MA' in col and any(col.endswith(f"MA{p}") for p in [2, 3, 5])]
        
        # Use any selected metrics or default to all numeric columns
        metrics_to_plot = numeric_cols[:3] if len(numeric_cols) > 3 else numeric_cols  # Default to first 3 metrics to avoid overcrowding
        
        # Get previously selected columns if any
        default_selection = []
        if slide_id and 'slides' in st.session_state:
            for slide in st.session_state.slides:
                if slide['id'] == slide_id and 'selected_columns' in slide:
                    default_selection = slide['selected_columns']
                    break
        
        if chart_type == 'table':
            # Let user select which columns to include in the table
            st.subheader("Configure Table")
            selected_cols = st.multiselect(
                "Select columns to include in table", 
                numeric_cols, 
                default=default_selection if default_selection else numeric_cols[:5],
                key=f"table_columns_{slide_id if slide_id else 'temp'}"
            )
            
            # Display as a formatted table with currency formatting
            formatted_df = df.copy()
            for col in numeric_cols:
                if col in formatted_df.columns:
                    formatted_df[col] = formatted_df[col].apply(lambda x: f"${x:,.2f}M" if pd.notna(x) else '')
            
            display_cols = ['Year'] if 'Year' in formatted_df.columns else []
            display_cols.extend(selected_cols)
            
            if display_cols:
                # Filter to only selected columns plus Year
                display_df = formatted_df[display_cols] if all(col in formatted_df.columns for col in display_cols) else formatted_df
                st.dataframe(display_df, use_container_width=True)
            else:
                st.dataframe(formatted_df, use_container_width=True)
                
            return formatted_df, selected_cols
            
        elif chart_type == 'pie':
            # Let user select which metric to display in pie chart
            st.subheader("Configure Pie Chart")
            
            # Let user select which year to display
            available_years = sorted(df['Year'].unique().tolist()) if 'Year' in df.columns else []
            if not available_years:
                st.warning("No year data available for pie chart.")
                return df, []
                
            selected_year = st.selectbox(
                "Select year to display", 
                available_years,
                index=len(available_years)-1,  # Default to most recent year
                key=f"pie_year_{slide_id if slide_id else 'temp'}"
            )
            
            # Let user choose from financial columns or ratio columns
            chart_category = st.radio(
                "Select category to display",
                ["Financial Metrics", "Ratio Metrics"],
                key=f"pie_category_{slide_id if slide_id else 'temp'}"
            )
            
            cols_to_select_from = ratio_cols if chart_category == "Ratio Metrics" else financial_cols
            
            # Let user select metrics to display
            selected_metrics = st.multiselect(
                "Select metrics to include in pie chart", 
                cols_to_select_from, 
                default=default_selection if default_selection else cols_to_select_from[:3],
                key=f"pie_metrics_{slide_id if slide_id else 'temp'}"
            )
            
            if not selected_metrics:
                st.warning("Please select at least one metric to display")
                return df, []
            
            # Filter data for selected year
            year_data = df[df['Year'] == selected_year]
            
            if year_data.empty:
                st.warning(f"No data available for {selected_year}")
                return df, selected_metrics
            
            # Create data for the pie chart
            pie_data = []
            for metric in selected_metrics:
                if metric in year_data.columns:
                    value = year_data[metric].iloc[0]
                    if pd.notna(value) and value != 0:  # Only include non-zero values
                        pie_data.append({"Metric": metric, "Value": abs(value)})  # Use absolute value for pie chart
            
            if not pie_data:
                st.warning("No valid data for selected metrics and year.")
                return df, selected_metrics
            
            # Create pie chart
            pie_df = pd.DataFrame(pie_data)
            fig = px.pie(
                pie_df, 
                values='Value', 
                names='Metric',
                title=f"{ticker} - {chart_category} for {selected_year}",
                hole=0.4,  # Make it a donut chart for better appearance
            )
            
            # Improve formatting
            fig.update_traces(
                textposition='inside',
                textinfo='percent+label',
                marker=dict(line=dict(color='#FFFFFF', width=2))
            )
            
            st.plotly_chart(fig, use_container_width=True, key=f"pie_chart_{slide_id if slide_id else 'temp'}")
            return df, selected_metrics
            
        elif chart_type == 'area':
            # Let user select which metrics to display
            st.subheader("Configure Area Chart")
            selected_metrics = st.multiselect(
                "Select metrics to display", 
                numeric_cols, 
                default=default_selection if default_selection else metrics_to_plot[:3],
                key=f"area_metrics_{slide_id if slide_id else 'temp'}"
            )
            
            # Choose between stacked or overlay
            area_mode = st.radio(
                "Area Chart Mode",
                ["Stacked", "Overlay"],
                key=f"area_mode_{slide_id if slide_id else 'temp'}"
            )
            
            if not selected_metrics:
                st.warning("Please select at least one metric to display")
                return df, []
            
            # Create area chart
            if area_mode == "Stacked":
                fig = px.area(
                    df, 
                    x='Year', 
                    y=selected_metrics,
                    title=f"{ticker} - Stacked Financial Metrics Over Time",
                    labels={"value": "USD (Millions)", "variable": "Metric"},
                )
            else:  # Overlay mode
                fig = px.area(
                    df, 
                    x='Year', 
                    y=selected_metrics,
                    title=f"{ticker} - Financial Metrics Over Time",
                    labels={"value": "USD (Millions)", "variable": "Metric"},
                    groupnorm='fraction' if len(selected_metrics) > 1 else None  # Normalize if multiple metrics
                )
                fig.update_layout(hovermode='x unified')
            
            # Improve chart formatting
            fig.update_layout(
                xaxis_title="Year",
                yaxis_title="Amount (USD Millions)",
                legend_title="Metrics"
            )
            fig.update_yaxes(tickprefix="$", ticksuffix="M")
            
            st.plotly_chart(fig, use_container_width=True, key=f"line_chart_{slide_id if slide_id else 'temp'}")
            return df, selected_metrics
            
        elif chart_type == 'line':
            # Let user select which metrics to display
            st.subheader("Configure Line Chart")
            
            # Let user choose whether to include moving averages
            include_ma = st.checkbox(
                "Include Moving Averages", 
                value=False,
                key=f"line_ma_toggle_{slide_id if slide_id else 'temp'}"
            )
            
            # Choose moving average periods if enabled
            ma_periods = []
            if include_ma:
                ma_periods = st.multiselect(
                    "Select Moving Average Periods",
                    [3, 5, 10],
                    default=[3],
                    key=f"line_ma_periods_{slide_id if slide_id else 'temp'}"
                )
                
                # Update session state with moving average periods
                if slide_id and 'slides' in st.session_state:
                    for idx, slide in enumerate(st.session_state.slides):
                        if slide['id'] == slide_id:
                            st.session_state.slides[idx]['moving_average_periods'] = ma_periods
                            break
            
            # Filter metrics to exclude existing MA columns when choosing base metrics
            base_metrics = [col for col in numeric_cols if not any(f"MA{p}" in col for p in [3, 5, 10])]
            
            selected_metrics = st.multiselect(
                "Select metrics to display", 
                base_metrics, 
                default=default_selection if default_selection and all(metric in base_metrics for metric in default_selection) else base_metrics[:1],
                key=f"line_metrics_{slide_id if slide_id else 'temp'}"
            )  
            
            if not selected_metrics:
                st.warning("Please select at least one metric to display")
                return df, []
            
            # Find any existing MA columns for selected metrics
            all_plot_columns = selected_metrics.copy()
            if include_ma:
                for metric in selected_metrics:
                    for period in ma_periods:
                        ma_col = f"{metric}_MA{period}"
                        if ma_col in df.columns:
                            all_plot_columns.append(ma_col)
            
            # Create the figure
            fig = go.Figure()
            
            # Add base metric lines
            for metric in selected_metrics:
                fig.add_trace(go.Scatter(
                    x=df['Year'],
                    y=df[metric],
                    mode='lines+markers',
                    name=metric,
                    line=dict(width=3)
                ))
                
                # Add moving average lines
                if include_ma:
                    for period in ma_periods:
                        ma_col = f"{metric}_MA{period}"
                        if ma_col in df.columns:
                            fig.add_trace(go.Scatter(
                                x=df['Year'],
                                y=df[ma_col],
                                mode='lines',
                                name=f"{metric} {period}Y MA",
                                line=dict(dash='dash', width=2)
                            ))
            
            # Improve chart formatting
            fig.update_layout(
                title=f"{ticker} - Financial Metrics Over Time",
                xaxis_title="Year",
                yaxis_title="Amount (USD Millions)",
                legend_title="Metrics",
                hovermode='x unified'
            )
            fig.update_yaxes(tickprefix="$", ticksuffix="M")
            
            st.plotly_chart(fig, use_container_width=True, key=f"line_chart_{slide_id if slide_id else 'temp'}")
            return df, selected_metrics
            
        elif chart_type == 'bar_chart' or chart_type == 'bar':
            # Let user select which metrics to display
            st.subheader("Configure Bar Chart")
            selected_metrics = st.multiselect(
                "Select metrics to display", 
                numeric_cols, 
                default=default_selection if default_selection else metrics_to_plot[:1],
                key=f"bar_metrics_{slide_id if slide_id else 'temp'}"
            )  
            
            if not selected_metrics:
                st.warning("Please select at least one metric to display")
                return df, []
            
            # Different handling for comparison vs. single company
            if is_comparison and second_ticker:
                # For comparison, we need to prepare the data differently
                chart_data = []
                
                for metric in selected_metrics:
                    # Look for columns that match this metric for both companies
                    for year in df['Year'].unique():
                        year_data = df[df['Year'] == year]
                        
                        # Look for columns for first company
                        col1 = f"{metric}_{ticker}"
                        if col1 in df.columns:
                            value = year_data[col1].iloc[0] if not year_data[col1].isna().all() else None
                            if pd.notna(value):
                                chart_data.append({
                                    'Year': year,
                                    'Metric': metric,
                                    'Company': ticker,
                                    'Value': value
                                })
                        else:
                            pass  # Column not found
                        
                        # Look for columns for second company
                        col2 = f"{metric}_{second_ticker}"
                        if col2 in df.columns:
                            value = year_data[col2].iloc[0] if not year_data[col2].isna().all() else None
                            if pd.notna(value):
                                chart_data.append({
                                    'Year': year,
                                    'Metric': metric,
                                    'Company': second_ticker,
                                    'Value': value
                                })
                        else:
                            pass  # Column not found           
                # Create DataFrame for plotting
                plot_df = pd.DataFrame(chart_data)
                
                if not plot_df.empty:
                    # For comparison, use a grouped bar chart with companies as groups
                    fig = px.bar(plot_df, x='Year', y='Value', color='Company', barmode='group',
                              facet_col='Metric' if len(selected_metrics) > 1 else None,
                              title=f"{ticker} vs {second_ticker} - Financial Metrics Comparison",
                              labels={"Value": "USD (Millions)"})
                    
                    # Improve chart formatting
                    fig.update_layout(
                        xaxis_title="Year",
                        yaxis_title="Amount (USD Millions)",
                        legend_title="Company",
                        hovermode='x unified'
                    )
                    fig.update_yaxes(tickprefix="$", ticksuffix="M")
                    fig.update_traces(texttemplate='%{y:.1f}', textposition='outside')
                    
                    st.plotly_chart(fig, use_container_width=True, key=f"bar_comp_chart_{slide_id if slide_id else 'temp'}")
                else:
                    st.warning("No valid comparison data available for the selected metrics.")
            else:
                # Regular single-company bar chart
                fig = px.bar(df, x='Year', y=selected_metrics,
                          title=f"{ticker} - Financial Metrics by Year",
                          barmode='group',
                          labels={"value": "USD (Millions)", "variable": "Metric"})
                
                # Improve chart formatting
                fig.update_layout(
                    xaxis_title="Year",
                    yaxis_title="Amount (USD Millions)",
                    legend_title="Metrics",
                    hovermode='x unified'
                )
                fig.update_yaxes(tickprefix="$", ticksuffix="M")
                fig.update_traces(texttemplate='%{y:.1f}', textposition='outside')
                
                st.plotly_chart(fig, use_container_width=True, key=f"bar_chart_{slide_id if slide_id else 'temp'}")
            
            return df, selected_metrics
            
    except Exception as e:
        st.error(f"Error rendering chart: {str(e)}")
        return data, []
    
    return data, []

def main():
    st.title("📊 Financial Data Slide Builder")
    
    # Add a new slide button
    if st.button("➕ Add New Financial Slide", use_container_width=True, type="primary"):
        add_slide()
    
    # Display existing slides
    for idx, slide in enumerate(st.session_state.slides):
        with st.expander(f"Slide {idx + 1}: {slide['title'] or 'Financial Data'}", expanded=True):
            col1, col2 = st.columns([5, 1])
            
            with col1:
                # Slide title and company info
                st.text_input("Slide Title", 
                            key=f"title_{slide['id']}",
                            value=slide['title'],
                            on_change=lambda idx=idx: update_slide_title(idx, f"title_{slide['id']}"))
                
                # Company and data selection
                st.markdown("### Primary Company")
                col1a, col1b = st.columns(2)
                with col1a:
                    ticker = st.text_input("Primary Company Ticker", 
                                          key=f"ticker_{slide['id']}",
                                          value=slide.get('ticker', ''),
                                          placeholder="e.g., AAPL",
                                          help="Enter the primary company's stock ticker symbol",
                                          on_change=lambda idx=idx: update_slide_field(idx, 'ticker', f"ticker_{slide['id']}"))
                
                with col1b:
                    # Chart Type Selector
                    chart_type = st.selectbox(
                        "Select Chart Type", 
                        ["table", "line", "bar_chart", "pie", "area"], 
                        index=["table", "line", "bar_chart", "pie", "area"].index(slide.get('chart_type', 'table')),
                        key=f"chart_type_{slide['id']}",
                        on_change=lambda: update_slide_field(idx, 'chart_type', f"chart_type_{slide['id']}"),
                        help="Select how to visualize the financial data"
                    )
                    
                    # Analysis Options
                    st.write("**Analysis Options:**")
                    
                    # Use horizontal layout without nested columns
                    enable_ratios = st.checkbox(
                        "Enable Financial Ratios", 
                        value=slide.get('enable_ratios', True),
                        key=f"enable_ratios_{slide['id']}",
                        help="Calculate financial ratios like profit margins, ROA, etc."
                    )
                    if enable_ratios != slide.get('enable_ratios', True):
                        st.session_state.slides[idx]['enable_ratios'] = enable_ratios
                        st.rerun()
                    
                    enable_trend = st.checkbox(
                        "Enable Trend Analysis", 
                        value=slide.get('enable_trend_analysis', False),
                        key=f"enable_trend_{slide['id']}",
                        help="Show trend analysis with CAGR and growth metrics"
                    )
                    if enable_trend != slide.get('enable_trend_analysis', False):
                        st.session_state.slides[idx]['enable_trend_analysis'] = enable_trend
                        st.rerun()
                
                years = st.slider("Years to Display", 
                                min_value=1, 
                                max_value=10, 
                                value=slide.get('years', 5),
                                key=f"years_{slide['id']}",
                                help="Number of years of historical data to display",
                                on_change=lambda idx=idx: update_slide_field(idx, 'years', f"years_{slide['id']}"))
                
                # Option to add comparison company
                st.markdown("### Comparison Company (Optional)")
                ticker2 = st.text_input("Comparison Company Ticker", 
                                      key=f"ticker2_{slide['id']}",
                                      value=slide.get('ticker2', ''),
                                      placeholder="e.g., MSFT",
                                      help="Enter a second company to compare with the primary company",
                                      on_change=lambda idx=idx: update_slide_field(idx, 'ticker2', f"ticker2_{slide['id']}"))
                
                # Fetch and display data - two buttons for fetching primary and comparison
                fetch_col1, fetch_col2 = st.columns(2)
                with fetch_col1:
                    if ticker and st.button("🔍 Fetch Primary Company", key=f"fetch_{slide['id']}"):
                        data = display_company_metrics(idx, ticker, years)
                        if data is not None:
                            st.session_state.slides[idx]['chart_data'] = data.to_dict()
                            st.success(f"Successfully fetched data for {ticker}")
                
                with fetch_col2:
                    if ticker2 and st.button("🔍 Fetch Comparison Company", key=f"fetch2_{slide['id']}"):
                        data2 = display_company_metrics(idx, ticker2, years, is_second_company=True)
                        if data2 is not None:
                            st.session_state.slides[idx]['chart_data2'] = data2.to_dict()
                            st.success(f"Successfully fetched data for {ticker2}")
                
                # Display data if available
                if 'chart_data' in slide and slide['chart_data']:
                    st.markdown("---")
                    
                    # Chart type selection - simplified to just Table and Bar Chart as requested
                    chart_options = ["Table", "Bar Chart"]
                    chart_type = st.radio("Visualization Type",
                                         chart_options,
                                         key=f"chart_viz_type_{slide['id']}",
                                         horizontal=True,
                                         index=chart_options.index(slide.get('chart_type', 'Table').replace('_', ' ').title()) 
                                         if slide.get('chart_type') in ['table', 'bar_chart'] else 0)
                    
                    # Convert to lowercase for internal use
                    chart_type = chart_type.lower().replace(' ', '_')
                    st.session_state.slides[idx]['chart_type'] = chart_type
                    
                    # Create DataFrame from stored data
                    data = pd.DataFrame(slide['chart_data'])
                    
                    # Check if we have comparison data
                    is_comparison = False
                    merged_data = None
                    ticker2 = slide.get('ticker2', '')
                    
                    if 'chart_data2' in slide and slide['chart_data2'] and ticker2:
                        is_comparison = True
                        data2 = pd.DataFrame(slide['chart_data2'])
                        st.info(f"Showing comparison between {ticker} and {ticker2}")
                        
                        # Merge the data for comparison
                        try:
                            # Merge the data
                            merged_data = merge_company_data(data, data2, ticker, ticker2)
                            
                            if merged_data is not None:
                                # Title reflects comparison
                                st.subheader(f"{ticker} vs {ticker2} Financial Comparison")
                                _, selected_columns = render_chart(chart_type, merged_data, ticker, slide['id'], 
                                                                is_comparison=True, second_ticker=ticker2)
                            else:
                                st.error("Failed to merge company data for comparison.")
                                _, selected_columns = render_chart(chart_type, data, ticker, slide['id'])
                        except Exception as e:
                            st.error(f"Error preparing comparison data: {str(e)}")
                            _, selected_columns = render_chart(chart_type, data, ticker, slide['id'])
                    else:
                        # Regular single company visualization
                        st.subheader(f"{ticker} Financial Data")
                        _, selected_columns = render_chart(chart_type, data, ticker, slide['id'])
                    
                    # Store selected columns for the slide
                    st.session_state.slides[idx]['selected_columns'] = selected_columns
                    
                    # For comparison data, we need to store the merged data
                    if is_comparison and merged_data is not None:
                        st.session_state.slides[idx]['merged_data'] = merged_data.to_dict()
                    
                    # Add Export to PowerPoint button
                    if st.button("Export to PowerPoint", key=f"export_pptx_{slide['id']}"):
                        try:
                            # Create temporary file
                            with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp:
                                temp_path = tmp.name
                            
                            # Make a copy of the slide data with all necessary information
                            slide_export_data = st.session_state.slides[idx].copy()
                            
                            # Ensure we have the chart type in the correct format
                            if slide_export_data.get('chart_type') == 'bar':
                                slide_export_data['chart_type'] = 'bar_chart'
                            
                            # Generate the presentation
                            output_path = generate_presentation([slide_export_data], temp_path)
                            
                            # Read the file for download
                            with open(output_path, "rb") as file:
                                bytes_data = file.read()
                            
                            # Provide download button with appropriate filename
                            filename = f"{ticker}"
                            if 'ticker2' in slide and slide['ticker2']:
                                filename += f"_vs_{slide['ticker2']}"
                            filename += f"_{chart_type}_presentation.pptx"
                            
                            st.download_button(
                                label="Download PowerPoint",
                                data=bytes_data,
                                file_name=filename,
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                            )
                            
                            # Clean up temp file after download
                            try:
                                os.remove(temp_path)
                            except:
                                pass
                                
                            st.success("PowerPoint file generated successfully!")
                        except Exception as e:
                            st.error(f"Error generating PowerPoint: {str(e)}")
                    
                    # Display trend analysis if enabled
                    if slide.get('enable_trend_analysis', False) and 'trend_analysis' in slide:
                        st.markdown("---")
                        st.markdown("### Financial Trend Analysis")
                        
                        trend_data = slide['trend_analysis']
                        if trend_data:
                            # Display CAGR and recent trends in a table
                            trend_rows = []
                            for metric, values in trend_data.items():
                                # Skip moving average columns
                                if any(f"MA{p}" in metric for p in [2, 3, 5, 10]):
                                    continue
                                    
                                cagr = values.get('cagr', None)
                                recent_trend = values.get('recent_trend', None)
                                latest = values.get('latest', None)
                                
                                if cagr is not None and not pd.isna(cagr):
                                    trend_rows.append({
                                        "Metric": metric,
                                        "Latest Value": f"${latest:,.2f}M" if pd.notna(latest) else "N/A",
                                        "CAGR": f"{cagr:.2f}%" if pd.notna(cagr) else "N/A",
                                        "Recent Trend": f"{recent_trend:.2f}%" if pd.notna(recent_trend) else "N/A",
                                    })
                            
                            if trend_rows:
                                trend_df = pd.DataFrame(trend_rows)
                                st.dataframe(trend_df, use_container_width=True)
                                
                                # Create a bar chart for CAGR comparison
                                metrics = [row["Metric"] for row in trend_rows]
                                cagr_values = [float(row["CAGR"].replace("%", "")) if "%" in row["CAGR"] else 0 for row in trend_rows]
                                
                                if len(metrics) > 0 and len(cagr_values) > 0:
                                    cagr_df = pd.DataFrame({"Metric": metrics, "CAGR %": cagr_values})
                                    fig = px.bar(
                                        cagr_df, 
                                        x="Metric", 
                                        y="CAGR %",
                                        title=f"{ticker} - Compound Annual Growth Rate (CAGR)",
                                        color="CAGR %",
                                        color_continuous_scale="RdYlGn",  # Red for negative, green for positive
                                    )
                                    fig.update_layout(yaxis_title="CAGR %")
                                    st.plotly_chart(fig, use_container_width=True, key=f"cagr_chart_{slide['id']}")
                            else:
                                st.info("Not enough historical data to calculate trends.")
                        else:
                            st.info("No trend analysis data available. Please refresh the data.")
                    
                    # Display raw data section
                    st.markdown("---")
                    raw_data_toggle = st.checkbox("View Raw Data", key=f"raw_data_toggle_{slide['id']}")
                    if raw_data_toggle:
                        st.markdown("### Raw Financial Data")
                        st.dataframe(data, use_container_width=True)
                        
                        # Option to download the data as CSV
                        csv = data.to_csv(index=False).encode('utf-8')
                        st.download_button(
                            label="Download Data as CSV",
                            data=csv,
                            file_name=f"{ticker}_financial_data.csv",
                            mime="text/csv",
                        )
                
            with col2:
                # Remove slide button
                if st.button("🗑️", 
                           key=f"remove_{slide['id']}",
                           help="Delete this slide"):
                    remove_slide(slide['id'])
                    st.rerun()
    
    # Show empty state if no slides
    if not st.session_state.slides:
        st.info("""
        ### No financial slides yet! 
        
        Click 'Add New Financial Slide' to get started!
        
        With this tool, you can:
        1. Enter a company ticker symbol (e.g., AAPL for Apple)
        2. Select how many years of historical data to retrieve
        3. View the financial data in tables or interactive charts
        4. Generate presentation slides with this data
        """)

def update_slide_title(slide_idx, title_key):
    """Update the title of a slide"""
    if title_key in st.session_state:
        st.session_state.slides[slide_idx]['title'] = st.session_state[title_key]

def update_slide_field(slide_idx, field, field_key):
    """Update a field in a slide"""
    if field_key in st.session_state:
        st.session_state.slides[slide_idx][field] = st.session_state[field_key]

if __name__ == "__main__":
    main()
