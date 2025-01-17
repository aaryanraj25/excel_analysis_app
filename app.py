import pandas as pd
import streamlit as st
import plotly.express as px
from io import BytesIO
import json
import os
from datetime import datetime

# Initialize session state for stored links
if 'stored_links' not in st.session_state:
    # Try to load existing links from file
    try:
        with open('stored_links.json', 'r') as f:
            st.session_state.stored_links = json.load(f)
    except FileNotFoundError:
        st.session_state.stored_links = {}

def save_links():
    """Save stored links to file"""
    with open('stored_links.json', 'w') as f:
        json.dump(st.session_state.stored_links, f)

def load_and_clean_data(files):
    """Load and process uploaded Excel files"""
    dataframes = {}
    for file in files:
        try:
            df = pd.read_excel(file)
            df = process_dataframe(df)
            if df is not None:
                dataframes[file.name] = df
        except Exception as e:
            st.error(f"Error processing {file.name}: {str(e)}")
    return dataframes

def load_from_link(link):
    """Load data from a saved Excel link"""
    try:
        df = pd.read_excel(link)
        return process_dataframe(df)
    except Exception as e:
        st.error(f"Error loading data from link: {str(e)}")
        return None

def process_dataframe(df):
    """Process and validate dataframe"""
    # Ensure required columns exist
    required_columns = ['Account', 'A/C Holder Name', 'State']
    if not all(col in df.columns for col in required_columns):
        st.error(f"Missing required columns. Please ensure the file contains: {', '.join(required_columns)}")
        return None
    
    # Calculate Total and Average columns for numeric columns only
    numeric_cols = df.select_dtypes(include=['float64', 'int64']).columns
    if len(numeric_cols) > 0:
        df['Total'] = df[numeric_cols].sum(axis=1)
        df['Average'] = df[numeric_cols].mean(axis=1)
    else:
        st.warning("No numeric columns found in the data")
        return None
        
    return df

def analyze_data(dataframes):
    """Analyze and display visualizations for the data"""
    try:
        file_names = list(dataframes.keys())
        selected_file = st.selectbox("Select File to Analyze", file_names)
        
        st.header('Overall Statistics')
        combined_stats = create_combined_statistics(dataframes)
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Packets (All Files)", f"{combined_stats['total_packets']:,}")
        with col2:
            st.metric("Total Accounts (All Files)", combined_stats['total_account_holders'])
        with col3:
            st.metric("Average Packets per Month (All Files)", 
                     f"{combined_stats['avg_packets_per_month']:,.0f}")
        
        df = dataframes[selected_file]
        st.header(f'Analysis for {selected_file}')
        stats = calculate_statistics(df)
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("File Total Packets", f"{stats['total_packets']:,}")
        with col2:
            st.metric("File Number of Accounts", len(df))
        with col3:
            st.metric("File Average Packets per Month", 
                     f"{stats['monthly_averages'].mean():,.0f}")
        
        # Monthly trend chart
        trend_chart = create_monthly_trend_chart(df)
        if trend_chart:
            st.plotly_chart(trend_chart, use_container_width=True)
        
        # Distribution charts
        col1, col2 = st.columns(2)
        with col1:
            account_pie = create_account_distribution_pie(df)
            if account_pie:
                st.plotly_chart(account_pie)
        with col2:
            state_pie = create_state_distribution_pie(df)
            if state_pie:
                st.plotly_chart(state_pie)
        
        # Account holder analysis
        st.header('Account Holder Analysis')
        tab1, tab2 = st.tabs(["Bar Chart", "Line Chart"])
        
        with tab1:
            bar_chart = create_account_holder_bar_chart(df)
            if bar_chart:
                st.plotly_chart(bar_chart, use_container_width=True)
        
        with tab2:
            line_chart = create_monthly_holder_comparison(df)
            if line_chart:
                st.plotly_chart(line_chart, use_container_width=True)
        
        # Heatmap
        heatmap = create_heatmap(df)
        if heatmap:
            st.plotly_chart(heatmap)
            
            st.header('Detailed Data View')
            cols = [col for col in df.columns if col not in ['Total', 'Average']] + ['Total', 'Average']
            df_display = df[cols]
            st.dataframe(df_display)
            
            col1, col2 = st.columns(2)
            with col1:
                csv_data = df_display.to_csv(index=False).encode('utf-8')
                st.download_button(
                    "Download current file data as CSV",
                    csv_data,
                    f"analyzed_data_{selected_file}.csv",
                    "text/csv"
                )
            
            with col2:
                combined_df = pd.concat(dataframes.values(), keys=dataframes.keys())
                combined_csv = combined_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    "Download all analyzed data as CSV",
                    combined_csv,
                    "all_analyzed_data.csv",
                    "text/csv"
                )
    except Exception as e:
        st.error(f"An error occurred while processing the data: {str(e)}")

def main():
    st.set_page_config(layout="wide")
    st.title('Novoxis Multi-File Analysis Dashboard')
    
    # Sidebar for adding new links
    with st.sidebar:
        st.header("Manage Data Sources")
        new_link = st.text_input("Add new Excel sheet link")
        link_name = st.text_input("Give this link a name")
        if st.button("Add Link") and new_link and link_name:
            st.session_state.stored_links[link_name] = new_link
            save_links()
            st.success(f"Added link: {link_name}")
        
        # Show stored links
        st.header("Stored Data Sources")
        for name, link in st.session_state.stored_links.items():
            col1, col2 = st.columns([3, 1])
            col1.write(f"{name}: {link}")
            if col2.button("Remove", key=f"remove_{name}"):
                del st.session_state.stored_links[name]
                save_links()
                st.rerun()
    
    # Main content area
    tab1, tab2 = st.tabs(["File Upload", "Saved Links"])
    
    with tab1:
        uploaded_files = st.file_uploader("Choose Excel files", type=['xlsx', 'xls'], accept_multiple_files=True)
        if uploaded_files:
            dataframes = load_and_clean_data(uploaded_files)
            if dataframes:
                analyze_data(dataframes)
    
    with tab2:
        if st.session_state.stored_links:
            selected_links = st.multiselect(
                "Select links to analyze",
                options=list(st.session_state.stored_links.keys())
            )
            
            if selected_links:
                dataframes = {}
                for name in selected_links:
                    df = load_from_link(st.session_state.stored_links[name])
                    if df is not None:
                        dataframes[name] = df
                
                if dataframes:
                    analyze_data(dataframes)
        else:
            st.info("No saved links available. Add links using the sidebar.")

if __name__ == '__main__':
    main()
