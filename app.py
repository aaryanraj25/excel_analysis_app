import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

# Original Packet Analysis Functions
def load_and_clean_packet_data(file):
    df = pd.read_excel(file)
    df.set_index('Account', inplace=True, drop=False)
    
    # Calculate Total and Average columns
    numeric_cols = df.select_dtypes(include=['float64', 'int64']).columns
    df['Total'] = df[numeric_cols].sum(axis=1)
    df['Average'] = df[numeric_cols].mean(axis=1)
    
    return df

def calculate_packet_statistics(df):
    numeric_df = df.select_dtypes(include=['float64', 'int64']).drop(['Total', 'Average'], axis=1, errors='ignore')
    return {
        'total_packets': numeric_df.sum().sum(),
        'monthly_averages': numeric_df.mean(),
        'account_totals': numeric_df.sum(axis=1),
        'active_months': numeric_df.notna().sum(axis=1)
    }

# New Invoice Analysis Functions
def load_and_clean_invoice_data(file):
    df = pd.read_excel(file)
    # Ensure required columns exist
    required_cols = ['Sr No', 'Invoice No', 'Account Holder Name', 'Customer Name', 'Amount']
    if not all(col in df.columns for col in required_cols):
        st.error("Excel file must contain: Sr No, Invoice No, Account Holder Name, Customer Name, Amount")
        return None
    
    # Convert Amount to numeric, handling any currency symbols or commas
    df['Amount'] = pd.to_numeric(df['Amount'].astype(str).str.replace(r'[^\d.-]', '', regex=True), errors='coerce')
    
    return df

def calculate_invoice_statistics(df):
    return {
        'total_amount': df['Amount'].sum(),
        'average_invoice': df['Amount'].mean(),
        'total_invoices': len(df),
        'unique_customers': df['Customer Name'].nunique(),
        'unique_account_holders': df['Account Holder Name'].nunique()
    }

def create_invoice_visualizations(df):
    # Top 10 customers by amount
    customer_totals = df.groupby('Customer Name')['Amount'].sum().sort_values(ascending=False)
    fig_customers = px.bar(customer_totals.head(10), 
                          title='Top 10 Customers by Invoice Amount',
                          labels={'value': 'Total Amount', 'Customer Name': 'Customer'})
    
    # Account holder distribution
    account_totals = df.groupby('Account Holder Name')['Amount'].sum()
    fig_accounts = px.pie(values=account_totals.values, 
                         names=account_totals.index, 
                         title='Distribution by Account Holder')
    
    # Invoice amount histogram
    fig_histogram = px.histogram(df, x='Amount', 
                                title='Distribution of Invoice Amounts',
                                nbins=30)
    
    # Customer-Account Holder relationship heatmap
    pivot_table = pd.crosstab(df['Customer Name'], df['Account Holder Name'], values=df['Amount'], aggfunc='sum')
    fig_heatmap = px.imshow(pivot_table,
                           title='Customer-Account Holder Relationship Map',
                           labels=dict(x='Account Holder', y='Customer', color='Amount'))
    
    return fig_customers, fig_accounts, fig_histogram, fig_heatmap

# Original Packet Analysis Visualizations
def create_monthly_trend_chart(df):
    monthly_data = df.select_dtypes(include=['float64', 'int64']).drop(['Total', 'Average'], axis=1, errors='ignore')
    monthly_totals = monthly_data.sum()
    fig = px.line(x=monthly_totals.index, y=monthly_totals.values, 
                  title='Monthly Product Packet Trends')
    fig.update_layout(yaxis_title='Number of Packets')
    return fig

def create_account_distribution_pie(df):
    fig = px.pie(values=df['Total'], names=df.index, 
                 title='Distribution of Packets by Account')
    return fig

def create_monthly_comparison_bar(df):
    monthly_data = df.select_dtypes(include=['float64', 'int64']).drop(['Total', 'Average'], axis=1, errors='ignore')
    fig = px.bar(monthly_data, title='Monthly Packet Comparison by Account', barmode='group')
    fig.update_layout(yaxis_title='Number of Packets')
    return fig

def create_heatmap(df):
    numeric_data = df.select_dtypes(include=['float64', 'int64']).drop(['Total', 'Average'], axis=1, errors='ignore')
    fig = px.imshow(numeric_data.T, title='Product Packet Quantity Heatmap', aspect='auto')
    return fig

def main():
    st.title('Novoxis Analysis Dashboard')
    
    # Add a tab selection for different analyses
    analysis_type = st.radio("Select Analysis Type", ["Packet Analysis", "Invoice Analysis"])
    
    if analysis_type == "Packet Analysis":
        uploaded_file = st.file_uploader("Choose Packet Data Excel file", type=['xlsx', 'xls'], key='packet')
        
        if uploaded_file:
            df = load_and_clean_packet_data(uploaded_file)
            stats = calculate_packet_statistics(df)
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Packets", f"{stats['total_packets']:,}")
            with col2:
                st.metric("Number of Accounts", len(df))
            with col3:
                st.metric("Average Packets per Month", f"{stats['monthly_averages'].mean():,.0f}")
            
            st.plotly_chart(create_monthly_trend_chart(df))
            st.plotly_chart(create_account_distribution_pie(df))
            st.plotly_chart(create_monthly_comparison_bar(df))
            st.plotly_chart(create_heatmap(df))
            
            st.header('Detailed Data View')
            cols = [col for col in df.columns if col not in ['Total', 'Average']] + ['Total', 'Average']
            df = df[cols]
            st.dataframe(df)
            
            st.download_button("Download analyzed packet data as CSV",
                             df.to_csv(), "analyzed_packet_data.csv", "text/csv")
    
    else:  # Invoice Analysis
        uploaded_file = st.file_uploader("Choose Invoice Data Excel file", type=['xlsx', 'xls'], key='invoice')
        
        if uploaded_file:
            df = load_and_clean_invoice_data(uploaded_file)
            if df is not None:
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
                
                # Display visualizations
                fig_customers, fig_accounts, fig_histogram, fig_heatmap = create_invoice_visualizations(df)
                st.plotly_chart(fig_customers)
                st.plotly_chart(fig_accounts)
                st.plotly_chart(fig_histogram)
                st.plotly_chart(fig_heatmap)
                
                st.header('Detailed Invoice Data')
                st.dataframe(df)
                
                st.download_button("Download analyzed invoice data as CSV",
                                 df.to_csv(index=False), 
                                 "analyzed_invoice_data.csv", 
                                 "text/csv")

if __name__ == '__main__':
    main()