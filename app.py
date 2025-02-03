import pandas as pd
import streamlit as st
import plotly.express as px
from io import BytesIO
import json
import os
from datetime import datetime
import requests

# Set page configuration at the very beginning
st.set_page_config(layout="wide", page_title="Novoxis Analysis Dashboard")

# Constants
API_BASE_URL = "https://aaryanraj25-scriptapilinkmanager.web.val.run/links"

# Initialize session state
if 'initialized' not in st.session_state:
    st.session_state.initialized = True
    try:
        response = requests.get(API_BASE_URL)
        if response.status_code == 200:
            links = response.json()
            st.session_state.stored_links = {link['name']: link['url'] for link in links}
        else:
            st.session_state.stored_links = {}
    except Exception as e:
        st.error(f"Error initializing links: {str(e)}")
        st.session_state.stored_links = {}

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

def calculate_packet_statistics(df):
    """Calculate statistics for pieces data"""
    numeric_df = df.select_dtypes(include=['float64', 'int64']).drop(['Total', 'Average'], axis=1, errors='ignore')
    return {
        'total_packets': numeric_df.sum().sum(),
        'monthly_averages': numeric_df.mean(),
        'account_totals': numeric_df.sum(axis=1),
        'active_months': numeric_df.notna().sum(axis=1)
    }
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
        
        trend_chart, account_pie, state_pie, heatmap, bar_chart = create_packet_visualizations(df)
        
        # Display trend chart
        st.plotly_chart(trend_chart, use_container_width=True)
        
        # Display pie charts in two columns
        col1, col2 = st.columns(2)
        with col1:
            st.plotly_chart(account_pie)
        with col2:
            st.plotly_chart(state_pie)
        
        # Display heatmap
        st.plotly_chart(heatmap, use_container_width=True)
        
        # Display bar chart
        st.plotly_chart(bar_chart, use_container_width=True)
            
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

def create_packet_visualizations(df):
    """Create visualizations for pieces data"""
    # Monthly trend chart
    monthly_data = df.select_dtypes(include=['float64', 'int64']).drop(['Total', 'Average'], axis=1, errors='ignore')
    monthly_totals = monthly_data.sum()
    trend_chart = px.line(x=monthly_totals.index, y=monthly_totals.values,
                         title='Monthly Product Pieces Trends')
    trend_chart.update_layout(yaxis_title='Number of Pieces')
    
    # Account distribution pie chart with 2% threshold
    total_pieces = df['Total'].sum()
    account_data = df.groupby('Account')['Total'].sum().reset_index()
    account_data['Percentage'] = (account_data['Total'] / total_pieces) * 100
    
    # Separate accounts with less than 2%
    others_mask = account_data['Percentage'] < 2
    others_sum = account_data[others_mask]['Total'].sum()
    
    # Create new dataframe with consolidated 'Others' category
    pie_data = account_data[~others_mask].copy()
    if others_sum > 0:
        others_row = pd.DataFrame({
            'Account': ['Others'],
            'Total': [others_sum],
            'Percentage': [(others_sum / total_pieces) * 100]
        })
        pie_data = pd.concat([pie_data, others_row])
    
    account_pie = px.pie(pie_data, values='Total', names='Account',
                        title='Distribution of Pieces by Account (< 2% grouped as Others)')
    
    # State distribution pie chart
    state_distribution = df.groupby('State')['Total'].sum()
    state_pie = px.pie(values=state_distribution.values,
                      names=state_distribution.index,
                      title='Distribution of Pieces by State')
    
    # Heatmap of monthly distribution by account
    heatmap_data = monthly_data.T
    heatmap = px.imshow(heatmap_data,
                        title='Monthly Distribution Heatmap by Account',
                        labels=dict(x='Account', y='Month', color='Pieces'),
                        aspect='auto')
    heatmap.update_layout(
        xaxis_title='Account',
        yaxis_title='Month'
    )
    
    # Bar chart based on A/C Holder Name
    holder_data = df.groupby('A/C Holder Name')['Total'].sum().sort_values(ascending=False)
    bar_chart = px.bar(x=holder_data.index, y=holder_data.values,
                      title='Total Pieces by Account Holder',
                      labels={'x': 'Account Holder Name', 'y': 'Total Pieces'})
    bar_chart.update_layout(
        xaxis_tickangle=-45,
        xaxis_title='Account Holder Name',
        yaxis_title='Total Pieces'
    )
    
    return trend_chart, account_pie, state_pie, heatmap, bar_chart

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
    customer_totals = df.groupby('Customer Name')['Amount'].sum()
    customer_pie = px.pie(values=customer_totals.values,
                         names=customer_totals.index,
                         title='Distribution by Customer')
    
    df['Month'] = pd.to_datetime(df['Invoice No'].str[:6], format='%y%m%d', errors='coerce').dt.strftime('%Y-%m')
    monthly_totals = df.groupby('Month')['Amount'].sum()
    monthly_trend = px.line(x=monthly_totals.index, y=monthly_totals.values,
                           title='Monthly Invoice Trends')
    monthly_trend.update_layout(xaxis_title='Month', yaxis_title='Amount')
    
    top_customers = df.groupby('Customer Name')['Amount'].sum().sort_values(ascending=False).head(10)
    top_customers_bar = px.bar(
        data_frame=pd.DataFrame({
            'Customer': top_customers.index,
            'Amount': top_customers.values
        }),
        x='Customer',
        y='Amount',
        title='Top 10 Customers by Invoice Amount'
    )
    top_customers_bar.update_layout(
        xaxis_title='Customer Name',
        yaxis_title='Total Amount',
        xaxis_tickangle=-45
    )
    
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
        
        trend_chart, account_pie, state_pie, heatmap, bar_chart = create_packet_visualizations(df)
        
        # Display trend chart
        st.plotly_chart(trend_chart, use_container_width=True)
        
        # Display pie charts in two columns
        col1, col2 = st.columns(2)
        with col1:
            st.plotly_chart(account_pie)
        with col2:
            st.plotly_chart(state_pie)
        
        # Display heatmap
        st.plotly_chart(heatmap, use_container_width=True)
        
        # Display bar chart
        st.plotly_chart(bar_chart, use_container_width=True)
            
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
    st.title('Novoxis Analysis Dashboard')
    
    # Sidebar for managing data sources
    with st.sidebar:
        st.header("Manage Data Sources")
        new_link = st.text_input("Add new Excel/Google Sheet link")
        link_name = st.text_input("Give this link a name")
        
        if st.button("Add Link") and new_link and link_name:
            try:
                payload = {
                    "name": link_name,
                    "url": new_link
                }
                response = requests.post(API_BASE_URL, json=payload)
                if response.status_code == 200:
                    response = requests.get(API_BASE_URL)
                    if response.status_code == 200:
                        links = response.json()
                        st.session_state.stored_links = {link['name']: link['url'] for link in links}
                    st.success(f"Added link: {link_name}")
                else:
                    st.error(f"Failed to save link: {response.status_code}")
            except Exception as e:
                st.error(f"Error saving link: {str(e)}")
        
        st.header("Stored Data Sources")
        for name, link in st.session_state.stored_links.items():
            st.write(f"{name}: {link}")
        
        if st.button("Refresh Links"):
            try:
                response = requests.get(API_BASE_URL)
                if response.status_code == 200:
                    links = response.json()
                    st.session_state.stored_links = {link['name']: link['url'] for link in links}
                    st.success("Links refreshed successfully")
                else:
                    st.error(f"Failed to refresh links: {response.status_code}")
            except Exception as e:
                st.error(f"Error refreshing links: {str(e)}")

    # Main content area
    tab1, tab2 = st.tabs(["File Upload", "Saved Links"])
    
    with tab1:
        uploaded_files = st.file_uploader("Choose Excel files", type=['xlsx', 'xls'], accept_multiple_files=True)
        if uploaded_files:
            dataframes = load_and_clean_data(uploaded_files)

            # Analyze packet data if available
            if dataframes['packet']:
                if len(dataframes['packet']) > 1:
                    st.header("Combined Packet Analysis Dashboard")
                    create_combined_dashboard(dataframes['packet'])

                st.header("Individual Packet Analysis")
                analyze_packet_data(dataframes['packet'])

            # Analyze invoice data if available
            if dataframes['invoice']:
                st.header("Invoice Analysis")
                analyze_invoice_data(dataframes['invoice'])

    with tab2:
        if st.session_state.stored_links:
            st.header("Analysis from Saved Links")
            
            # Add filter options
            filter_col1, filter_col2 = st.columns(2)
            with filter_col1:
                file_type_filter = st.multiselect(
                    "Filter by File Type",
                    options=["Packet Data", "Invoice Data"],
                    default=["Packet Data", "Invoice Data"]
                )
            
            with filter_col2:
                date_filter = st.date_input(
                    "Filter by Date (Optional)",
                    value=None,
                    help="Filter files by date if available"
                )
            
            selected_links = st.multiselect(
                "Select links to analyze",
                options=list(st.session_state.stored_links.keys())
            )
            
            if selected_links:
                with st.spinner("Loading and analyzing data..."):
                    dataframes = {'packet': {}, 'invoice': {}}
                    
                    # Load data from selected links
                    for name in selected_links:
                        df = load_from_link(st.session_state.stored_links[name])
                        if df is not None:
                            file_type = detect_file_type(df)
                            if file_type == 'packet' and "Packet Data" in file_type_filter:
                                dataframes['packet'][name] = df
                            elif file_type == 'invoice' and "Invoice Data" in file_type_filter:
                                dataframes['invoice'][name] = df
                    
                    # Display analysis based on loaded data
                    if dataframes['packet']:
                        st.header("Packet Data Analysis")
                        
                        # Show combined analysis if multiple packet files
                        if len(dataframes['packet']) > 1:
                            with st.expander("View Combined Packet Analysis", expanded=True):
                                create_combined_dashboard(dataframes['packet'])
                        
                        # Individual packet analysis
                        with st.expander("View Individual Packet Analysis", expanded=False):
                            analyze_packet_data(dataframes['packet'])
                    
                    if dataframes['invoice']:
                        st.header("Invoice Data Analysis")
                        
                        # Show combined invoice analysis if multiple invoice files
                        if len(dataframes['invoice']) > 1:
                            with st.expander("View Combined Invoice Analysis", expanded=True):
                                create_combined_invoice_dashboard(dataframes['invoice'])
                        
                        # Individual invoice analysis
                        with st.expander("View Individual Invoice Analysis", expanded=False):
                            analyze_invoice_data(dataframes['invoice'])
                    
                    # Add comparison tools
                    if dataframes['packet'] or dataframes['invoice']:
                        st.header("Analysis Tools")
                        
                        tool_col1, tool_col2 = st.columns(2)
                        with tool_col1:
                            if st.button("Generate Summary Report"):
                                generate_summary_report(dataframes)
                        
                        with tool_col2:
                            if st.button("Export All Analysis"):
                                export_all_analysis(dataframes)
        else:
            st.info("No saved links available. Add links using the sidebar.")

def create_combined_invoice_dashboard(invoice_dataframes):
    """Create and display combined dashboard for multiple invoice files"""
    try:
        # Combined statistics
        total_amount = sum(df['Amount'].sum() for df in invoice_dataframes.values())
        total_invoices = sum(len(df) for df in invoice_dataframes.values())
        avg_amount = total_amount / total_invoices
        
        # Display metrics
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Amount (All Sources)", f"₹{total_amount:,.2f}")
        with col2:
            st.metric("Total Invoices (All Sources)", total_invoices)
        with col3:
            st.metric("Average Invoice Amount", f"₹{avg_amount:,.2f}")
        
        # Combine all dataframes
        combined_df = pd.concat(invoice_dataframes.values(), keys=invoice_dataframes.keys())
        combined_df = combined_df.reset_index(level=0).rename(columns={'level_0': 'Source'})
        
        # Create visualizations
        # Source comparison
        source_totals = combined_df.groupby('Source')['Amount'].sum()
        source_comparison = px.bar(
            x=source_totals.index,
            y=source_totals.values,
            title='Total Amount Comparison Across Sources'
        )
        st.plotly_chart(source_comparison, use_container_width=True)
        
        # Customer distribution
        col1, col2 = st.columns(2)
        with col1:
            customer_totals = combined_df.groupby('Customer Name')['Amount'].sum()
            customer_pie = px.pie(
                values=customer_totals.values,
                names=customer_totals.index,
                title='Combined Distribution by Customer'
            )
            st.plotly_chart(customer_pie)
        
        with col2:
            monthly_trend = px.line(
                combined_df.groupby(['Source', 'Month'])['Amount'].sum().reset_index(),
                x='Month',
                y='Amount',
                color='Source',
                title='Monthly Trends Comparison'
            )
            st.plotly_chart(monthly_trend)
        
        # Download combined data
        st.subheader("Download Combined Data")
        csv_data = combined_df.to_csv(index=False).encode('utf-8')
        st.download_button(
            "Download Combined Invoice Analysis",
            csv_data,
            "combined_invoice_analysis.csv",
            "text/csv"
        )
        
    except Exception as e:
        st.error(f"Error in combined invoice analysis: {str(e)}")

def generate_summary_report(dataframes):
    """Generate a summary report of all analyzed data"""
    try:
        report = "# Analysis Summary Report\n\n"
        
        if dataframes['packet']:
            report += "## Packet Data Summary\n"
            for name, df in dataframes['packet'].items():
                report += f"\n### {name}\n"
                report += f"- Total Packets: {df['Total'].sum():,}\n"
                report += f"- Number of Accounts: {len(df)}\n"
                report += f"- States Covered: {', '.join(df['State'].unique())}\n"
        
        if dataframes['invoice']:
            report += "\n## Invoice Data Summary\n"
            for name, df in dataframes['invoice'].items():
                report += f"\n### {name}\n"
                report += f"- Total Amount: ₹{df['Amount'].sum():,.2f}\n"
                report += f"- Number of Invoices: {len(df)}\n"
                report += f"- Unique Customers: {df['Customer Name'].nunique()}\n"
        
        # Convert report to PDF or download as markdown
        st.markdown(report)
        st.download_button(
            "Download Summary Report",
            report,
            "analysis_summary_report.md",
            "text/markdown"
        )
        
    except Exception as e:
        st.error(f"Error generating summary report: {str(e)}")

def export_all_analysis(dataframes):
    """Export all analysis data as a zip file"""
    try:
        # Create a BytesIO object to store the zip file
        zip_buffer = BytesIO()
        
        # Create timestamp for filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Export packet data
        if dataframes['packet']:
            for name, df in dataframes['packet'].items():
                df.to_csv(f"packet_analysis_{name}_{timestamp}.csv", index=False)
        
        # Export invoice data
        if dataframes['invoice']:
            for name, df in dataframes['invoice'].items():
                df.to_csv(f"invoice_analysis_{name}_{timestamp}.csv", index=False)
        
        # Create download button
        st.download_button(
            "Download All Analysis Data",
            zip_buffer.getvalue(),
            f"complete_analysis_{timestamp}.zip",
            "application/zip"
        )
        
    except Exception as e:
        st.error(f"Error exporting analysis data: {str(e)}")

if __name__ == '__main__':
    main()