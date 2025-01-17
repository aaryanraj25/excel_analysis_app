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

def create_account_distribution_pie(df):
    # Using Account column for distribution
    fig = px.pie(values=df['Total'], names=df['Account'], 
                 title='Distribution of Packets by Account')
    fig.update_layout(
        showlegend=True,
        legend_title='Account'
    )
    return fig

def create_monthly_holder_comparison(df):
    """
    Create a single plot with months on x-axis and A/C Holder Name as variables,
    all unselected by default
    """
    # Get data excluding Total and Average columns
    monthly_data = df.select_dtypes(include=['float64', 'int64']).drop(['Total', 'Average'], axis=1, errors='ignore')
    
    # Melt the dataframe to get it in the right format for plotting
    melted_data = monthly_data.reset_index()
    melted_data = pd.melt(melted_data, 
                         id_vars=['A/C Holder Name'], 
                         var_name='Month', 
                         value_name='Packets')
    
    # Create the figure
    fig = px.line(melted_data, 
                  x='Month', 
                  y='Packets', 
                  color='A/C Holder Name',
                  title='Monthly Distribution by A/C Holder Name',
                  labels={'Month': 'Month',
                         'Packets': 'Number of Packets',
                         'A/C Holder Name': 'A/C Holder Name'})
    
    # Update layout to hide all traces by default
    fig.update_layout(
        showlegend=True,
        legend_title='A/C Holder Name',
        height=500,
        xaxis_title='Month',
        yaxis_title='Number of Packets'
    )
    
    # Set all traces to invisible by default
    for trace in fig.data:
        trace.visible = "legendonly"
    
    return fig

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
    
    # Set page configuration
    st.set_page_config(
        page_title="Novoxis Analysis Dashboard",
        page_icon="📊",
        layout="wide"
    )
    
    # Multiple file upload
    uploaded_files = st.file_uploader("Choose Excel files", type=['xlsx', 'xls'], accept_multiple_files=True)
    
    if uploaded_files:
        try:
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
            
            # Account-based distribution
            st.header('Account Distribution')
            st.plotly_chart(create_account_distribution_pie(df))
            
            # A/C Holder Name monthly comparison
            st.header('Monthly A/C Holder Distribution')
            st.plotly_chart(create_monthly_holder_comparison(df))
            
            st.plotly_chart(create_heatmap(df))
            
            # Detailed Data View
            st.header('Detailed Data View')
            cols = [col for col in df.columns if col not in ['Total', 'Average']] + ['Total', 'Average']
            df = df[cols]
            st.dataframe(df)
            
            # Download buttons
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    "Download current file data as CSV",
                    df.to_csv(),
                    f"analyzed_data_{selected_file}.csv",
                    "text/csv"
                )
            
            with col2:
                combined_df = pd.concat(dataframes.values(), keys=dataframes.keys())
                st.download_button(
                    "Download all analyzed data as CSV",
                    combined_df.to_csv(),
                    "all_analyzed_data.csv",
                    "text/csv"
                )
                
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
            st.error("Please check your Excel file format and try again.")

if __name__ == '__main__':
    main()
