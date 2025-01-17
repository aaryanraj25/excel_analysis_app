import pandas as pd
import streamlit as st
import plotly.express as px
from io import BytesIO

def load_and_clean_data(files):
    dataframes = {}
    for file in files:
        df = pd.read_excel(file)
        # Use the correct column name 'A/C Holder Name'
        df.set_index('A/C Holder Name', inplace=True, drop=False)
        
        # Calculate Total and Average columns
        numeric_cols = df.select_dtypes(include=['float64', 'int64']).columns
        df['Total'] = df[numeric_cols].sum(axis=1)
        df['Average'] = df[numeric_cols].mean(axis=1)
        
        dataframes[file.name] = df
    return dataframes

def calculate_statistics(df):
    numeric_df = df.select_dtypes(include=['float64', 'int64']).drop(['Total', 'Average'], axis=1, errors='ignore')
    return {
        'total_packets': numeric_df.sum().sum(),
        'monthly_averages': numeric_df.mean(),
        'account_holder_totals': numeric_df.sum(axis=1),
        'active_months': numeric_df.notna().sum(axis=1)
    }

def create_monthly_trend_chart(df):
    monthly_data = df.select_dtypes(include=['float64', 'int64']).drop(['Total', 'Average'], axis=1, errors='ignore')
    monthly_totals = monthly_data.sum()
    fig = px.line(x=monthly_totals.index, y=monthly_totals.values, 
                  title='Monthly Product Packet Trends')
    fig.update_layout(yaxis_title='Number of Packets')
    return fig

def create_account_holder_distribution_pie(df):
    fig = px.pie(values=df['Total'], names=df.index, 
                 title='Distribution of Packets by A/C Holder Name')
    fig.update_layout(
        showlegend=True,
        legend_title='A/C Holder Name'
    )
    return fig

def create_account_holder_comparison_bar(df):
    # Get data excluding Total and Average columns
    monthly_data = df.select_dtypes(include=['float64', 'int64']).drop(['Total', 'Average'], axis=1, errors='ignore')
    
    # Create separate plots for each month
    figs = []
    for month in monthly_data.columns:
        month_data = monthly_data[month].reset_index()
        fig = px.bar(month_data, 
                     x='A/C Holder Name',
                     y=month,
                     title=f'A/C Holder Distribution for {month}',
                     labels={'A/C Holder Name': 'A/C Holder Name', 
                            month: 'Number of Packets'})
        fig.update_layout(
            showlegend=False,
            xaxis_tickangle=45,
            height=400
        )
        figs.append(fig)
    
    return figs

def create_heatmap(df):
    numeric_data = df.select_dtypes(include=['float64', 'int64']).drop(['Total', 'Average'], axis=1, errors='ignore')
    fig = px.imshow(numeric_data.T, 
                    title='Product Packet Quantity Heatmap', 
                    labels={'y': 'Month', 'x': 'A/C Holder Name'},
                    aspect='auto')
    return fig

def create_combined_statistics(dataframes):
    combined_stats = {
        'total_packets': 0,
        'total_account_holders': 0,
        'avg_packets_per_month': 0
    }
    
    for df in dataframes.values():
        stats = calculate_statistics(df)
        combined_stats['total_packets'] += stats['total_packets']
        combined_stats['total_account_holders'] += len(df)
        combined_stats['avg_packets_per_month'] += stats['monthly_averages'].mean()
    
    if dataframes:
        combined_stats['avg_packets_per_month'] /= len(dataframes)
    
    return combined_stats

def main():
    st.title('Novoxis Multi-File Analysis Dashboard')
    
    # Multiple file upload
    uploaded_files = st.file_uploader("Choose Excel files", type=['xlsx', 'xls'], accept_multiple_files=True)
    
    if uploaded_files:
        dataframes = load_and_clean_data(uploaded_files)
        file_names = list(dataframes.keys())
        selected_file = st.selectbox("Select File to Analyze", file_names)
        
        # Display combined statistics
        st.header('Overall Statistics')
        combined_stats = create_combined_statistics(dataframes)
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Packets (All Files)", f"{combined_stats['total_packets']:,}")
        with col2:
            st.metric("Total A/C Holders (All Files)", combined_stats['total_account_holders'])
        with col3:
            st.metric("Average Packets per Month (All Files)", 
                     f"{combined_stats['avg_packets_per_month']:,.0f}")
        
        # Display individual file analysis
        st.header(f'Analysis for {selected_file}')
        df = dataframes[selected_file]
        stats = calculate_statistics(df)
        
        # Individual file metrics
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("File Total Packets", f"{stats['total_packets']:,}")
        with col2:
            st.metric("File Number of A/C Holders", len(df))
        with col3:
            st.metric("File Average Packets per Month", 
                     f"{stats['monthly_averages'].mean():,.0f}")
        
        # Charts
        st.plotly_chart(create_monthly_trend_chart(df))
        st.plotly_chart(create_account_holder_distribution_pie(df))
        
        # Monthly A/C Holder Distribution Charts
        st.header('Monthly A/C Holder Distribution')
        account_holder_charts = create_account_holder_comparison_bar(df)
        for fig in account_holder_charts:
            st.plotly_chart(fig)
            
        st.plotly_chart(create_heatmap(df))
        
        # Detailed Data View
        st.header('Detailed Data View')
        cols = [col for col in df.columns if col not in ['Total', 'Average']] + ['Total', 'Average']
        df = df[cols]
        st.dataframe(df)
        
        # Download buttons
        st.download_button(
            "Download current file data as CSV",
            df.to_csv(),
            f"analyzed_data_{selected_file}.csv",
            "text/csv"
        )
        
        combined_df = pd.concat(dataframes.values(), keys=dataframes.keys())
        st.download_button(
            "Download all analyzed data as CSV",
            combined_df.to_csv(),
            "all_analyzed_data.csv",
            "text/csv"
        )

if __name__ == '__main__':
    main()
