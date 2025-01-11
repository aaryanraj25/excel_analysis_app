import pandas as pd
import streamlit as st
import plotly.express as px

def load_and_clean_data(file):
    df = pd.read_excel(file)
    df.set_index('Account', inplace=True, drop=False)
    
    # Calculate Total and Average columns
    numeric_cols = df.select_dtypes(include=['float64', 'int64']).columns
    df['Total'] = df[numeric_cols].sum(axis=1)
    df['Average'] = df[numeric_cols].mean(axis=1)
    
    return df

def calculate_statistics(df):
    # Exclude Total and Average columns from numeric calculations
    numeric_df = df.select_dtypes(include=['float64', 'int64']).drop(['Total', 'Average'], axis=1, errors='ignore')
    return {
        'total_packets': numeric_df.sum().sum(),
        'monthly_averages': numeric_df.mean(),
        'account_totals': numeric_df.sum(axis=1),
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
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'])
    
    if uploaded_file:
        df = load_and_clean_data(uploaded_file)
        stats = calculate_statistics(df)
        
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
        # Reorder columns to show Total and Average at the end
        cols = [col for col in df.columns if col not in ['Total', 'Average']] + ['Total', 'Average']
        df = df[cols]
        st.dataframe(df)
        
        st.download_button("Download analyzed data as CSV",
                         df.to_csv(), "analyzed_data.csv", "text/csv")

if __name__ == '__main__':
    main()
