import pandas as pd
import numpy as np
import plotly.express as px
from tabulate import tabulate

def read_all_excel_sheets(file_path):
    """
    Reads all sheets from an Excel file into a dictionary of DataFrames.
    
    Args:
        file_path (str): Path to Excel file.
        
    Returns:
        dict: Dictionary where keys are sheet names and values are DataFrames.
    """
    sheets_dict = pd.read_excel(
        file_path,
        sheet_name=None,       # Read all sheets
        engine='openpyxl',     # Required engine for .xlsx files
        dtype=str,             # Read all as strings to preserve formatting
        na_values=['', 'NA']   # Optional NA values
    )
    return sheets_dict
def clean_strings(df):
    """
    Cleans all string (object) columns in a DataFrame.
    
    Operations performed:
    1. Strips leading and trailing whitespace from string values.
    2. Converts all string values to uppercase.
    3. Fills missing values (NaN) with a placeholder '-'.
    
    Args:
        df (pd.DataFrame): Input DataFrame containing columns of various data types.
        
    Returns:
        pd.DataFrame: DataFrame with cleaned string columns. Non-string columns are left unchanged.
    
    Example:
        >>> df = pd.DataFrame({'Name': [' Alice ', None], 'Age': [25, 30]})
        >>> clean_strings(df)
           Name  Age
        0  ALICE   25
        1      -   30
    """
    for col in df.select_dtypes(include='object').columns:
        df[col] = df[col].str.strip().str.upper().fillna('-')
    return df

def data_quality(df: pd.DataFrame) -> pd.DataFrame:
    """Return a summary of data quality for each column in a DataFrame."""
    return pd.DataFrame({
        "Data Type": df.dtypes.astype(str),
        "Missing Values": df.isna().sum(),
        "% Missing": (df.isna().mean() * 100).round(2),
        "Unique Values": df.nunique()
    }).sort_values(by="% Missing", ascending=False)
def df_overview(file_path, sheet_name):
    """
    Provides an advanced interactive overview of an Excel sheet using Plotly Express
    
    Features:
    - Basic metadata summary
    - Data quality assessment
    - Statistical overview
    - Interactive visualizations using Plotly Express
    - Sample data display
    """
    try:
        # Read data
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
        
        # 1. Metadata
        metadata = {
            "Sheet Name": sheet_name,
            "Rows": df.shape[0],
            "Columns": df.shape[1],
            "Memory Usage": f"{df.memory_usage(deep=True).sum() / 1024**2:.2f} MB",
            "Duplicate Rows": df.duplicated().sum()
        }
        print(f"\n{'-'*60}\nDATA OVERVIEW: {sheet_name}\n{'-'*60}")
        print(tabulate([[k, v] for k, v in metadata.items()], 
                      headers=['Metadata', 'Value'], tablefmt='pretty'))
        
        # 2. Data Quality
        dq = data_quality(df)
        print("\nDATA QUALITY ASSESSMENT")
        print(tabulate(dq, headers='keys', tablefmt='psql', showindex=True))
        
        # 3. Sample values
        sample_values = {}
        for col in df.columns:
            sample_values[col] = df[col].dropna().sample(min(3, len(df[col]))).values if not df[col].dropna().empty else []
        print("\nVALUE SAMPLE")
        print(tabulate(pd.DataFrame(sample_values), headers='keys', tablefmt='psql'))
        
        # 4. Visualizations using Plotly Express

        # Missing values
        fig_missing = px.bar(
            dq.reset_index(),
            x='index',
            y='% Missing',
            color='% Missing',
            color_continuous_scale='Viridis',
            title=f'Missing Values by Column: {sheet_name}',
            labels={'index': 'Column', '% Missing': 'Percentage Missing'}
        )
        fig_missing.update_layout(xaxis_tickangle=-45)
        fig_missing.show()

        # Numerical distributions
        num_cols = df.select_dtypes(include=np.number).columns
        for col in num_cols:
            fig_num = px.histogram(
                df,
                x=col,
                nbins=50,
                marginal='box',
                title=f'Distribution of {col}'
            )
            fig_num.show()

        # Categorical distributions
        cat_cols = df.select_dtypes(include=['object', 'category']).columns
        for col in cat_cols:
            top_10 = df[col].value_counts().nlargest(10).reset_index()
            top_10.columns = ['Value', 'Count']
            fig_cat = px.bar(
                top_10,
                x='Value',
                y='Count',
                title=f'Top 10 Values of {col}',
                color='Count',
                color_continuous_scale='Bluered'
            )
            fig_cat.show()

        # Correlation matrix
        if len(num_cols) > 1:
            corr = df[num_cols].corr().round(2)
            fig_corr = px.imshow(
                corr,
                text_auto=True,
                title='Correlation Matrix',
                color_continuous_scale='RdBu_r',
                zmin=-1,
                zmax=1
            )
            fig_corr.update_layout(height=600)
            fig_corr.show()

        # 5. Sample data
        print(f"\nDATA PREVIEW (First 5 rows)")
        print(tabulate(df.head(), headers='keys', tablefmt='psql'))
        
    except Exception as e:
        print(f"Error processing data: {e}")
