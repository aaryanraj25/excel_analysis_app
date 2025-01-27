import pandas as pd
import streamlit as st
import plotly.express as px
from io import BytesIO
import json
import os
from datetime import datetime

# Initialize session state for stored links
if 'stored_links' not in st.session_state:
    try:
        with open('stored_links.json', 'r') as f:
            st.session_state.stored_links = json.load(f)
    except FileNotFoundError:
        st.session_state.stored_links = {}

def save_links():
    """Save stored links to file"""
    with open('stored_links.json', 'w') as f:
        json.dump(st.session_state.stored_links, f)

def detect_file_type(df):
    """Detect whether the file is invoice or packet data"""
    invoice_columns = ['Sr No', 'Invoice No', 'Account Holder Name', 'Customer Name', 'Amount']
    packet_columns = ['Account', 'A/C Holder Name', 'State']
    
    if all(col in df.columns for col in invoice_columns):
        return 'invoice'
    elif all(col in df.columns for col in packet_columns):
        return 'packet'
    else:
        return None

def load_and_clean_data(files):
    """Load and process uploaded Excel files"""
    dataframes = {'packet': {}, 'invoice': {}}
    for file in files:
        try:
            df = pd.read_excel(file)
            file_type = detect_file_type(df)
            if file_type == 'packet':
                processed_df = process_packet_dataframe(df)
                if processed_df is not None:
                    dataframes['packet'][file.name] = processed_df
            elif file_type == 'invoice':
                processed_df = process_invoice_dataframe(df)
                if processed_df is not None:
                    dataframes['invoice'][file.name] = processed_df
            else:
                st.error(f"Unrecognized file format for {file.name}")
        except Exception as e:
            st.error(f"Error processing {file.name}: {str(e)}")
    return dataframes

def process_packet_dataframe(df):
    """Process and validate packet dataframe"""
    required_columns = ['Account', 'A/C Holder Name', 'State']
    if not all(col in df.columns for col in required_columns):
        st.error(f"Missing required columns for packet data. Please ensure the file contains: {', '.join(required_columns)}")
        return None
    
    numeric_cols = df.select_dtypes(include=['float64', 'int64']).columns
    if len(numeric_cols) > 0:
        df['Total'] = df[numeric_cols].sum(axis=1)
        df['Average'] = df[numeric_cols].mean(axis=1)
    else:
        st.warning("No numeric columns found in the packet data")
        return None
        
    return df

def process_invoice_dataframe(df):
    """Process and validate invoice dataframe"""
    required_columns = ['Sr No', 'Invoice No', 'Account Holder Name', 'Customer Name', 'Amount']
    if not all(col in df.columns for col in required_columns):
        st.error(f"Missing required columns for invoice data. Please ensure the file contains: {', '.join(required_columns)}")
        return None
    
    # Convert Amount to numeric
    df['Amount'] = pd.to_numeric(df['Amount'].astype(str).str.replace(r'[^\d.-]', '', regex=True), errors='coerce')
    return df

def load_from_link(link):
    """Load data from a saved link (Google Sheets or Excel)"""
    try:
        if 'docs.google.com/spreadsheets' in link:
            df = load_google_sheet(link)
        else:
            df = pd.read_excel(link)
        
        file_type = detect_file_type(df)
        if file_type == 'packet':
            return process_packet_dataframe(df)
        elif file_type == 'invoice':
            return process_invoice_dataframe(df)
        else:
            st.error("Unrecognized file format in the linked file")
            return None
    except Exception as e:
        st.error(f"Error loading data from link: {str(e)}")
        return None

def load_google_sheet(link):
    """Load data from a public Google Sheets link"""
    try:
        sheet_id = link.split('/d/')[1].split('/')[0]
        url = f'https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv'
        return pd.read_csv(url)
    except Exception as e:
        st.error(f"Error loading data from Google Sheets: {str(e)}")
        return None

# Packet Analysis Functions
def calculate_packet_statistics(df):
    """Calculate statistics for packet data"""
    numeric_df = df.select_dtypes(include=['float64', 'int64']).drop(['Total', 'Average'], axis=1, errors='ignore')
    return {
        'total_packets': numeric_df.sum().sum(),
        'monthly_averages': numeric_df.mean(),
        'account_totals': numeric_df.sum(axis=1),
        'active_months': numeric_df.notna().sum(axis=1)
    }

def create_packet_visualizations(df):
    """Create visualizations for packet data"""
    # Monthly trend chart
    monthly_data = df.select_dtypes(include=['float64', 'int64']).drop(['Total', 'Average'], axis=1, errors='ignore')
    monthly_totals = monthly_data.sum()
    trend_chart = px.line(x=monthly_totals.index, y=monthly_totals.values,
                         title='Monthly Product Packet Trends')
    trend_chart.update_layout(yaxis_title='Number of Packets')
    
    # Account distribution pie chart
    account_pie = px.pie(values=df['Total'], names=df['Account'],
                        title='Distribution of Packets by Account')
    
    # State distribution pie chart
    state_distribution = df.groupby('State')['Total'].sum()
    state_pie = px.pie(values=state_distribution.values,
                      names=state_distribution.index,
                      title='Distribution of Packets by State')
    
    return trend_chart, account_pie, state_pie

# Invoice Analysis Functions
def calculate_invoice_statistics(df):
    """Calculate statistics for invoice data"""
    return {
        'total_amount': df['Amount'].sum(),
        'average_invoice': df['Amount'].mean(),
        'total_invoices': len(df),
        'unique_customers': df['Customer Name'].nunique(),
        'unique_account_holders': df['Account Holder Name'].nunique()
    }

def create_invoice_visualizations(df):
    """Create visualizations for invoice data"""
    # Customer distribution pie chart
    customer_totals = df.groupby('Customer Name')['Amount'].sum()
    customer_pie = px.pie(values=customer_totals.values,
                         names=customer_totals.index,
                         title='Distribution by Customer')
    
    # Monthly invoice trends
    df['Month'] = pd.to_datetime(df['Invoice No'].str[:6], format='%y%m%d', errors='coerce').dt.strftime('%Y-%m')
    monthly_totals = df.groupby('Month')['Amount'].sum()
    monthly_trend = px.line(x=monthly_totals.index, y=monthly_totals.values,
                           title='Monthly Invoice Trends')
    
    # Top customers bar chart
    top_customers = df.groupby('Customer Name')['Amount'].sum().sort_values(ascending=False).head(10)
    top_customers_bar = px.bar(x=top_customers.index, y=top_customers.values,
                              title='Top 10 Customers by Invoice Amount')
    
    return customer_pie, monthly_trend, top_customers_bar

def analyze_packet_data(dataframes):
    """Analyze and display visualizations for packet data"""
    try:
        file_names = list(dataframes.keys())
        selected_file = st.selectbox("Select Packet File to Analyze", file_names)
        
        df = dataframes[selected_file]
        st.header(f'Packet Analysis for {selected_file}')
        
        # Calculate and display statistics
        stats = calculate_packet_statistics(df)
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Packets", f"{stats['total_packets']:,}")
        with col2:
            st.metric("Number of Accounts", len(df))
        with col3:
            st.metric("Average Packets per Month", 
                     f"{stats['monthly_averages'].mean():,.0f}")
        
        # Create and display visualizations
        trend_chart, account_pie, state_pie = create_packet_visualizations(df)
        
        st.plotly_chart(trend_chart, use_container_width=True)
        
        col1, col2 = st.columns(2)
        with col1:
            st.plotly_chart(account_pie)
        with col2:
            st.plotly_chart(state_pie)
            
        # Display detailed data
        st.header('Detailed Data View')
        cols = [col for col in df.columns if col not in ['Total', 'Average']] + ['Total', 'Average']
        df_display = df[cols]
        st.dataframe(df_display)
        
        # Download buttons
        csv_data = df_display.to_csv(index=False).encode('utf-8')
        st.download_button(
            "Download analyzed data as CSV",
            csv_data,
            f"analyzed_packet_data_{selected_file}.csv",
            "text/csv"
        )
        
    except Exception as e:
        st.error(f"Error in packet analysis: {str(e)}")

def analyze_invoice_data(dataframes):
    """Analyze and display visualizations for invoice data"""
    try:
        file_names = list(dataframes.keys())
        selected_file = st.selectbox("Select Invoice File to Analyze", file_names)
        
        df = dataframes[selected_file]
        st.header(f'Invoice Analysis for {selected_file}')
        
        # Calculate and display statistics
        stats = calculate_invoice_statistics(df)
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Amount", f"₹{stats['total_amount']:,.2f}")
        with col2:
            st.metric("Total Invoices", stats['total_invoices'])
        with col3:
            st.metric("Average Invoice Amount", f"₹{stats['average_invoice']:,.2f}")
        
        col4, col5 = st.columns(2)
        with col4:
            st.metric("Unique Customers", stats['unique_customers'])
        with col5:
            st.metric("Unique Account Holders", stats['unique_account_holders'])
        
        # Create and display visualizations
        customer_pie, monthly_trend, top_customers_bar = create_invoice_visualizations(df)
        
        st.plotly_chart(monthly_trend, use_container_width=True)
        st.plotly_chart(top_customers_bar, use_container_width=True)
        st.plotly_chart(customer_pie, use_container_width=True)
        
        # Display detailed data
        st.header('Detailed Invoice Data')
        st.dataframe(df)
        
        # Download button
        csv_data = df.to_csv(index=False).encode('utf-8')
        st.download_button(
            "Download analyzed invoice data as CSV",
            csv_data,
            f"analyzed_invoice_data_{selected_file}.csv",
            "text/csv"
        )
        
    except Exception as e:
        st.error(f"Error in invoice analysis: {str(e)}")

def main():
    st.set_page_config(layout="wide")
    st.title('Novoxis Analysis Dashboard')
    
    # Sidebar for managing data sources
    with st.sidebar:
        st.header("Manage Data Sources")
        new_link = st.text_input("Add new Excel/Google Sheet link")
        link_name = st.text_input("Give this link a name")
        if st.button("Add Link") and new_link and link_name:
            st.session_state.stored_links[link_name] = new_link
            save_links()
            st.success(f"Added link: {link_name}")
        
        # Display stored links
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
            
            # Analyze packet data if available
            if dataframes['packet']:
                st.header("Packet Analysis")
                analyze_packet_data(dataframes['packet'])
            
            # Analyze invoice data if available
            if dataframes['invoice']:
                st.header("Invoice Analysis")
                analyze_invoice_data(dataframes['invoice'])
    
    with tab2:
        if st.session_state.stored_links:
            selected_links = st.multiselect(
                "Select links to analyze",
                options=list(st.session_state.stored_links.keys())
            )
            
            if selected_links:
                dataframes = {'packet': {}, 'invoice': {}}
                for name in selected_links:
                    df = load_from_link(st.session_state.stored_links[name])
                    if df is not None:
                        file_type = detect_file_type(df)
                        if file_type == 'packet':
                            dataframes['packet'][name] = df
                        elif file_type == 'invoice':
                            dataframes['invoice'][name] = df
                
                # Analyze packet data if available
                if dataframes['packet']:
                    st.header("Packet Analysis")
                    analyze_packet_data(dataframes['packet'])
                
                # Analyze invoice data if available
                if dataframes['invoice']:
                    st.header("Invoice Analysis")
                    analyze_invoice_data(dataframes['invoice'])
        else:
            st.info("No saved links available. Add links using the sidebar.")

if __name__ == '__main__':
    main()