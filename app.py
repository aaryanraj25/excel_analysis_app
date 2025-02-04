import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
from io import BytesIO
from datetime import datetime
import numpy as np
import zipfile

# Initialize session state for stored links
if 'stored_links' not in st.session_state:
    st.session_state.stored_links = {}

# API endpoint (replace with your actual API endpoint)
API_BASE_URL = "https://aaryanraj25-scriptapilinkmanager.web.val.run/links"

def load_and_clean_data(uploaded_files):
    """Load and clean data from uploaded files"""
    dataframes = {'packet': {}, 'invoice': {}}
    
    for uploaded_file in uploaded_files:
        try:
            # Try different Excel engines
            try:
                df = pd.read_excel(uploaded_file, engine='openpyxl')
            except:
                try:
                    df = pd.read_excel(uploaded_file, engine='xlrd')
                except:
                    df = pd.read_excel(uploaded_file, engine='odf')
            
            file_type = detect_file_type(df)
            
            if file_type == 'packet':
                df = clean_packet_data(df)
                dataframes['packet'][uploaded_file.name] = df
            elif file_type == 'invoice':
                df = clean_invoice_data(df)
                dataframes['invoice'][uploaded_file.name] = df
            
        except Exception as e:
            st.error(f"Error processing {uploaded_file.name}: {str(e)}")
    
    return dataframes

def load_from_url(url):
    """Load data from URL with multiple engine attempts"""
    try:
        response = requests.get(url)
        excel_data = BytesIO(response.content)

        # Try different Excel engines
        try:
            df = pd.read_excel(excel_data, engine='openpyxl')
        except:
            try:
                df = pd.read_excel(excel_data, engine='xlrd')
            except:
                df = pd.read_excel(excel_data, engine='odf')
        return df
    except Exception as e:
        st.error(f"Error loading data from URL: {str(e)}")
        return None

def detect_file_type(df):
    """Detect whether the file is a packet or invoice file"""
    if 'Account' in df.columns and 'Total' in df.columns:
        return 'packet'
    elif 'Invoice Number' in df.columns and 'Amount' in df.columns:
        return 'invoice'
    else:
        raise ValueError("Unknown file type")

def clean_packet_data(df):
    """Clean and preprocess packet data"""
    # Remove any empty rows and columns
    df = df.dropna(how='all').dropna(axis=1, how='all')

    # Ensure required columns exist
    required_cols = ['Account', 'State', 'A/C Holder Name', 'Total']
    if not all(col in df.columns for col in required_cols):
        raise ValueError("Missing required columns in packet data")

    # Convert numeric columns to appropriate type
    numeric_cols = df.select_dtypes(include=['float64', 'int64']).columns
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce')

    return df

def clean_invoice_data(df):
    """Clean and preprocess invoice data"""
    # Remove any empty rows and columns
    df = df.dropna(how='all').dropna(axis=1, how='all')

    # Ensure required columns exist
    required_cols = ['Invoice Number', 'Customer Name', 'Amount', 'Date']
    if not all(col in df.columns for col in required_cols):
        raise ValueError("Missing required columns in invoice data")

    # Convert date column to datetime
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')

    # Convert amount to numeric
    df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce')

    return df

def calculate_packet_statistics(df):
    """Calculate statistics for packet data"""
    stats = {
        'total_packets': df['Total'].sum(),
        'monthly_averages': df.select_dtypes(include=['float64', 'int64']).mean(),
        'state_distribution': df.groupby('State')['Total'].sum(),
        'account_distribution': df.groupby('Account')['Total'].sum()
    }
    return stats

def create_packet_visualizations(df):
    """Create visualizations for packet data"""
    # Monthly trend
    monthly_cols = df.select_dtypes(include=['float64', 'int64']).columns
    monthly_cols = [col for col in monthly_cols if col not in ['Total', 'Average']]

    if monthly_cols:
        trend_data = df[monthly_cols].sum()
        trend_chart = px.line(
            x=trend_data.index,
            y=trend_data.values,
            title='Monthly Distribution Trend'
        )
        trend_chart.update_layout(
            xaxis_title='Month',
            yaxis_title='Total Pieces'
        )
    else:
        trend_chart = None

    # Account distribution pie chart with 2% threshold
    total_pieces = df['Total'].sum()
    account_data = df.groupby('Account')['Total'].sum().reset_index()
    account_data['Percentage'] = (account_data['Total'] / total_pieces) * 100

    others_mask = account_data['Percentage'] < 2
    others_sum = account_data[others_mask]['Total'].sum()
    pie_data = account_data[~others_mask].copy()
    if others_sum > 0:
        others_row = pd.DataFrame({
            'Account': ['Others'],
            'Total': [others_sum],
            'Percentage': [(others_sum / total_pieces) * 100]
        })
        pie_data = pd.concat([pie_data, others_row])

    account_pie = px.pie(
        pie_data,
        values='Total',
        names='Account',
        title='Distribution by Account (< 2% grouped as Others)'
    )

    # State distribution
    state_pie = px.pie(
        values=df.groupby('State')['Total'].sum(),
        names=df.groupby('State')['Total'].sum().index,
        title='Distribution by State'
    )

    # Monthly heatmap
    monthly_data = df[monthly_cols]
    heatmap = px.imshow(
        monthly_data.corr(),
        title='Monthly Distribution Correlation',
        aspect='auto'
    )

    # Account holder bar chart
    bar_chart = px.bar(
        df,
        x='A/C Holder Name',
        y='Total',
        title='Distribution by Account Holder'
    )
    bar_chart.update_layout(xaxis_tickangle=-45)

    return trend_chart, account_pie, state_pie, heatmap, bar_chart

def create_combined_dashboard(packet_dataframes):
    """Create and display combined dashboard for multiple packet files"""
    try:
        # Combined statistics
        total_packets = sum(df['Total'].sum() for df in packet_dataframes.values())
        total_accounts = sum(len(df) for df in packet_dataframes.values())
        avg_packets = total_packets / len(packet_dataframes)

        # Display metrics
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Pieces (All Sources)", f"{total_packets:,}")
        with col2:
            st.metric("Total Accounts (All Sources)", total_accounts)
        with col3:
            st.metric("Average Pieces per Source", f"{avg_packets:,.0f}")

        # Combine all dataframes
        combined_df = pd.concat(packet_dataframes.values(), keys=packet_dataframes.keys())
        combined_df = combined_df.reset_index(level=0).rename(columns={'level_0': 'Source'})

        # Source comparison
        source_totals = combined_df.groupby('Source')['Total'].sum()
        source_comparison = px.bar(
            x=source_totals.index,
            y=source_totals.values,
            title='Total Pieces Comparison Across Sources'
        )
        source_comparison.update_layout(
            xaxis_title='Source',
            yaxis_title='Total Pieces',
            xaxis_tickangle=-45
        )
        st.plotly_chart(source_comparison, use_container_width=True)

        # Account distribution
        account_data = combined_df.groupby('Account')['Total'].sum().reset_index()
        total_pieces = account_data['Total'].sum()
        account_data['Percentage'] = (account_data['Total'] / total_pieces) * 100

        # Group accounts with less than 2%
        others_mask = account_data['Percentage'] < 2
        others_sum = account_data[others_mask]['Total'].sum()
        pie_data = account_data[~others_mask].copy()
        if others_sum > 0:
            others_row = pd.DataFrame({
                'Account': ['Others'],
                'Total': [others_sum],
                'Percentage': [(others_sum / total_pieces) * 100]
            })
            pie_data = pd.concat([pie_data, others_row])

        # Create visualizations
        col1, col2 = st.columns(2)

        with col1:
            account_pie = px.pie(
                pie_data,
                values='Total',
                names='Account',
                title='Combined Distribution of Pieces by Account (< 2% grouped as Others)'
            )
            st.plotly_chart(account_pie)

        with col2:
            state_distribution = combined_df.groupby('State')['Total'].sum()
            state_pie = px.pie(
                values=state_distribution.values,
                names=state_distribution.index,
                title='Combined Distribution of Pieces by State'
            )
            st.plotly_chart(state_pie)
            # Account holder comparison
        holder_comparison = px.bar(
            combined_df.groupby(['Source', 'A/C Holder Name'])['Total'].sum().reset_index(),
            x='A/C Holder Name',
            y='Total',
            color='Source',
            title='Account Holder Comparison Across Sources',
            barmode='group'
        )
        holder_comparison.update_layout(
            xaxis_tickangle=-45,
            xaxis_title='Account Holder Name',
            yaxis_title='Total Pieces'
        )
        st.plotly_chart(holder_comparison, use_container_width=True)
      
        # State comparison
        state_comparison = px.bar(
            combined_df.groupby(['Source', 'State'])['Total'].sum().reset_index(),
            x='State',
            y='Total',
            color='Source',
            title='State-wise Comparison Across Sources',
            barmode='group'
        )
        state_comparison.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(state_comparison, use_container_width=True)
      
        # Monthly trend comparison
        monthly_cols = combined_df.select_dtypes(include=['float64', 'int64']).columns
        monthly_cols = [col for col in monthly_cols if col not in ['Total', 'Average']]
      
        if monthly_cols:
            monthly_data = combined_df.groupby('Source')[monthly_cols].sum()
            monthly_trend = px.line(
                monthly_data.T,
                title='Monthly Trends Comparison Across Sources'
            )
            monthly_trend.update_layout(
                xaxis_title='Month',
                yaxis_title='Total Pieces'
            )
            st.plotly_chart(monthly_trend, use_container_width=True)
      
        # Download combined data
        st.subheader("Download Combined Data")
        csv_data = combined_df.to_csv(index=False).encode('utf-8')
        st.download_button(
            "Download Combined Analysis Data",
            csv_data,
            "combined_packet_analysis.csv",
            "text/csv"
        )
      
    except Exception as e:
        st.error(f"Error in combined analysis: {str(e)}")

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
                combined_df.groupby(['Source', pd.Grouper(key='Date', freq='M')])['Amount'].sum().reset_index(),
                x='Date',
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

def load_from_google_sheets(url):
    """Load data from Google Sheets URL"""
    try:
        # Convert Google Sheets URL to export URL
        if 'edit?usp=sharing' in url:
            url = url.replace('edit?usp=sharing', 'export?format=xlsx')
        elif 'edit' in url:
            url = url.replace('edit', 'export?format=xlsx')

        response = requests.get(url)
        excel_data = BytesIO(response.content)

        try:
            df = pd.read_excel(excel_data, engine='openpyxl')
        except:
            try:
                df = pd.read_excel(excel_data, engine='xlrd')
            except:
                df = pd.read_excel(excel_data, engine='odf')
        return df
    except Exception as e:
        st.error(f"Error loading data from Google Sheets: {str(e)}")
        return None

def main():
    st.title('Novoxis Analysis Dashboard')

    # Initialize session state
    if 'stored_links' not in st.session_state:
        st.session_state.stored_links = {}

    # Sidebar for managing data sources
    with st.sidebar:
        st.header("Manage Data Sources")

        # Add new link section
        new_link = st.text_input("Add new Excel/Google Sheet link")
        link_name = st.text_input("Give this link a name")

        col1, col2 = st.columns(2)
        with col1:
            if st.button("Add Link") and new_link and link_name:
                try:
                    payload = {
                        "name": link_name,
                        "url": new_link
                    }
                    response = requests.post(API_BASE_URL, json=payload)
                    if response.status_code == 200:
                        st.session_state.stored_links[link_name] = new_link
                        st.success(f"Added link: {link_name}")
                    else:
                        st.error(f"Failed to save link: {response.status_code}")
                except Exception as e:
                    st.error(f"Error saving link: {str(e)}")

        with col2:
            if st.button("🔄 Refresh Links"):
                try:
                    response = requests.get(API_BASE_URL)
                    if response.status_code == 200:
                        links = response.json()
                        st.session_state.stored_links = {link['name']: link['url'] for link in links}
                        st.success("Links refreshed successfully!")
                    else:
                        st.error(f"Failed to refresh links: {response.status_code}")
                except Exception as e:
                    st.error(f"Error refreshing links: {str(e)}")

        # Display stored links
        st.header("Stored Data Sources")
        if st.session_state.stored_links:
            for name, link in st.session_state.stored_links.items():
                st.write(f"**{name}**")
                st.write(f"{link}")
                if st.button("🗑️", key=f"delete_{name}"):
                    try:
                        response = requests.delete(f"{API_BASE_URL}/{name}")
                        if response.status_code == 200:
                            del st.session_state.stored_links[name]
                            st.success(f"Deleted: {name}")
                            st.rerun()
                        else:
                            st.error(f"Failed to delete link: {response.status_code}")
                    except Exception as e:
                        st.error(f"Error deleting link: {str(e)}")
        else:
            st.info("No stored links. Add links using the form above.")

    # Main content area with tabs
    tab1, tab2 = st.tabs(["File Upload", "Saved Links"])

    with tab1:
        uploaded_files = st.file_uploader("Choose Excel files", type=['xlsx', 'xls'], accept_multiple_files=True)
        if uploaded_files:
            dataframes = load_and_clean_data(uploaded_files)
            # Rest of your tab1 code...

    with tab2:
        if st.session_state.stored_links:
            st.header("Analysis from Saved Links")

            selected_links = st.multiselect(
                "Select links to analyze",
                options=list(st.session_state.stored_links.keys())
            )

            if selected_links:
                with st.spinner("Loading and analyzing data..."):
                    dataframes = {'packet': {}, 'invoice': {}}

                    for name in selected_links:
                        try:
                            df = load_from_google_sheets(st.session_state.stored_links[name])
                            if df is not None:
                                file_type = detect_file_type(df)

                                if file_type == 'packet':
                                    df = clean_packet_data(df)
                                    dataframes['packet'][name] = df
                                elif file_type == 'invoice':
                                    df = clean_invoice_data(df)
                                    dataframes['invoice'][name] = df

                        except Exception as e:
                            st.error(f"Error loading {name}: {str(e)}")

                    # Display analysis based on loaded data
                    if dataframes['packet']:
                        st.header("Pieces Data Analysis")
                        if len(dataframes['packet']) > 1:
                            create_combined_dashboard(dataframes['packet'])

                        for name, df in dataframes['packet'].items():
                            with st.expander(f"Analysis for {name}", expanded=False):
                                stats = calculate_packet_statistics(df)
                                trend_chart, account_pie, state_pie, heatmap, bar_chart = create_packet_visualizations(df)

                                st.write(f"Total Pieces: {stats['total_packets']:,}")
                                if trend_chart:
                                    st.plotly_chart(trend_chart, use_container_width=True)
                                st.plotly_chart(account_pie, use_container_width=True)
                                st.plotly_chart(state_pie, use_container_width=True)
                                st.plotly_chart(heatmap, use_container_width=True)
                                st.plotly_chart(bar_chart, use_container_width=True)

                    if dataframes['invoice']:
                        st.header("Invoice Data Analysis")
                        if len(dataframes['invoice']) > 1:
                            create_combined_invoice_dashboard(dataframes['invoice'])

                        for name, df in dataframes['invoice'].items():
                            with st.expander(f"Analysis for {name}", expanded=False):
                                st.write(f"Total Amount: ₹{df['Amount'].sum():,.2f}")
                                st.write(f"Number of Invoices: {len(df)}")
        else:
            st.info("No saved links available. Add links using the sidebar.")

if __name__ == "__main__":
    main()