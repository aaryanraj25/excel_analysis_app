import pandas as pd
import streamlit as st
import plotly.express as px
from io import BytesIO

def load_and_clean_data(files):
    dataframes = {}
    for file in files:
        df = pd.read_excel(file)
        
        # Ensure required columns exist
        required_columns = ['Account', 'A/C Holder Name', 'State']
        if not all(col in df.columns for col in required_columns):
            st.error(f"Missing required columns in {file.name}. Please ensure the file contains: {', '.join(required_columns)}")
            return None
        
        # Calculate Total and Average columns for numeric columns only
        numeric_cols = df.select_dtypes(include=['float64', 'int64']).columns
        if len(numeric_cols) > 0:
            df['Total'] = df[numeric_cols].sum(axis=1)
            df['Average'] = df[numeric_cols].mean(axis=1)
        else:
            st.warning(f"No numeric columns found in {file.name}")
            return None
            
        dataframes[file.name] = df
    return dataframes

def create_combined_statistics(dataframes):
    """Calculate combined statistics across all files"""
    total_packets = 0
    total_account_holders = 0
    all_monthly_averages = []
    
    for df in dataframes.values():
        numeric_df = df.select_dtypes(include=['float64', 'int64']).drop(['Total', 'Average'], axis=1, errors='ignore')
        total_packets += numeric_df.sum().sum()
        total_account_holders += len(df)
        all_monthly_averages.append(numeric_df.mean().mean())
    
    return {
        'total_packets': total_packets,
        'total_account_holders': total_account_holders,
        'avg_packets_per_month': sum(all_monthly_averages) / len(all_monthly_averages) if all_monthly_averages else 0
    }

def calculate_statistics(df):
    numeric_df = df.select_dtypes(include=['float64', 'int64']).drop(['Total', 'Average'], axis=1, errors='ignore')
    return {
        'total_packets': numeric_df.sum().sum(),
        'monthly_averages': numeric_df.mean(),
        'account_totals': numeric_df.sum(axis=1),
        'active_months': numeric_df.notna().sum(axis=1)
    }

def create_monthly_trend_chart(df):
    monthly_data = df.select_dtypes(include=['float64', 'int64']).drop(['Total', 'Average'], axis=1, errors='ignore')
    if monthly_data.empty:
        return None
    monthly_totals = monthly_data.sum()
    fig = px.line(x=monthly_totals.index, y=monthly_totals.values, 
                  title='Monthly Product Packet Trends')
    fig.update_layout(yaxis_title='Number of Packets')
    return fig

def create_account_distribution_pie(df):
    if 'Total' not in df.columns or 'Account' not in df.columns:
        return None
    fig = px.pie(values=df['Total'], names=df['Account'], 
                 title='Distribution of Packets by Account')
    fig.update_layout(showlegend=True, legend_title='Account')
    return fig

def create_state_distribution_pie(df):
    if 'Total' not in df.columns or 'State' not in df.columns:
        return None
    state_distribution = df.groupby('State')['Total'].sum()
    fig = px.pie(values=state_distribution.values, 
                 names=state_distribution.index,
                 title='Distribution of Packets by State')
    fig.update_layout(showlegend=True, legend_title='State')
    return fig

def create_monthly_holder_comparison(df):
    if 'AC Holder Name' not in df.columns:
        return None
        
    monthly_data = df.select_dtypes(include=['float64', 'int64']).drop(['Total', 'Average'], axis=1, errors='ignore')
    if monthly_data.empty:
        return None
        
    try:
        melted_data = monthly_data.reset_index()
        melted_data = pd.melt(melted_data, 
                             id_vars=['AC Holder Name'], 
                             var_name='Month', 
                             value_name='Packets')
        
        fig = px.line(melted_data, 
                      x='Month', 
                      y='Packets', 
                      color='AC Holder Name',
                      title='Monthly Distribution by A/C Holder Name')
        
        fig.update_layout(
            showlegend=True,
            legend=dict(yanchor="top", y=1, xanchor="left", x=1.02),
            legend_title='AC Holder Name',
            height=600,
            margin=dict(r=150),
            xaxis_title='Month',
            yaxis_title='Number of Packets'
        )
        
        for trace in fig.data:
            trace.visible = "legendonly"
            
        return fig
    except Exception as e:
        st.error(f"Error creating monthly holder comparison: {str(e)}")
        return None

def create_heatmap(df):
    numeric_data = df.select_dtypes(include=['float64', 'int64']).drop(['Total', 'Average'], axis=1, errors='ignore')
    if numeric_data.empty:
        return None
    fig = px.imshow(numeric_data.T, 
                    title='Product Packet Quantity Heatmap', 
                    labels={'y': 'Month', 'x': 'Account'},
                    aspect='auto')
    return fig

def main():
    st.set_page_config(layout="wide")
    st.title('Novoxis Multi-File Analysis Dashboard')
    
    uploaded_files = st.file_uploader("Choose Excel files", type=['xlsx', 'xls'], accept_multiple_files=True)
    
    if uploaded_files:
        with st.spinner('Loading and processing files...'):
            dataframes = load_and_clean_data(uploaded_files)
            
        if not dataframes:
            st.error("No valid data found in the uploaded files. Please check the file format and contents.")
            return
            
        file_names = list(dataframes.keys())
        selected_file = st.selectbox("Select File to Analyze", file_names)
        
        try:
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
            
            # Charts
            trend_chart = create_monthly_trend_chart(df)
            if trend_chart:
                st.plotly_chart(trend_chart)
            
            col1, col2 = st.columns(2)
            with col1:
                account_pie = create_account_distribution_pie(df)
                if account_pie:
                    st.plotly_chart(account_pie)
            with col2:
                state_pie = create_state_distribution_pie(df)
                if state_pie:
                    st.plotly_chart(state_pie)
            
            st.header('Monthly A/C Holder Distribution')
            monthly_chart = create_monthly_holder_comparison(df)
            if monthly_chart:
                st.plotly_chart(monthly_chart, use_container_width=True)
            
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

if __name__ == '__main__':
    main()
