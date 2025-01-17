# [Previous imports and other functions remain the same until create_monthly_holder_comparison]

def create_monthly_holder_comparison(df):
    # Debug information
    st.write("Available columns:", df.columns.tolist())
    
    # Check for possible variations of the account holder column name
    possible_names = ['A/C Holder Name', 'AC Holder Name', 'Account Holder Name', 
                     'Holder Name', 'Name', 'Account Holder']
    
    holder_column = None
    for name in possible_names:
        if name in df.columns:
            holder_column = name
            break
    
    if not holder_column:
        st.error(f"Could not find account holder column. Available columns: {', '.join(df.columns)}")
        return None
        
    monthly_data = df.select_dtypes(include=['float64', 'int64']).drop(['Total', 'Average'], axis=1, errors='ignore')
    if monthly_data.empty:
        st.error("No numeric data found for monthly comparison")
        return None
        
    try:
        # Include the correct holder column name in the melted data
        melted_data = pd.concat([df[holder_column], monthly_data], axis=1)
        melted_data = pd.melt(melted_data, 
                             id_vars=[holder_column], 
                             var_name='Month', 
                             value_name='Packets')
        
        fig = px.line(melted_data, 
                      x='Month', 
                      y='Packets', 
                      color=holder_column,  # Use the found column name
                      title='Monthly Distribution by Account Holder')
        
        fig.update_layout(
            showlegend=True,
            legend=dict(yanchor="top", y=1, xanchor="left", x=1.02),
            legend_title='Account Holder',
            height=600,
            margin=dict(r=150),
            xaxis_title='Month',
            yaxis_title='Number of Packets'
        )
        
        for trace in fig.data:
            trace.visible = "legendonly"
            
        return fig
    except Exception as e:
        st.error(f"Error in monthly holder comparison: {str(e)}\nData shape: {df.shape}")
        st.write("Data sample:", df.head())
        return None

# [Rest of the code remains the same]
