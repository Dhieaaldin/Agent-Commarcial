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


def df_overview(file_path, sheet_name):
    """
    Provides an advanced overview of an Excel sheet with metadata, data quality,
    value samples, and interactive visualizations.
    
    Args:
        file_path (str): Path to Excel file.
        sheet_name (str): Sheet name to analyze.
    """
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')

        # --- Metadata ---
        metadata = {
            "Sheet Name": sheet_name,
            "Rows": df.shape[0],
            "Columns": df.shape[1],
            "Memory Usage": f"{df.memory_usage(deep=True).sum() / 1024**2:.2f} MB",
            "Duplicate Rows": df.duplicated().sum()
        }

        # --- Data Quality ---
        dq = pd.DataFrame({
            'Data Type': df.dtypes.astype(str),
            'Missing Values': df.isnull().sum(),
            '% Missing': (df.isnull().mean() * 100).round(1),
            'Unique Values': df.nunique()
        })

        # --- Sample Values ---
        sample_values = {}
        for col in df.columns:
            if df[col].dtype == 'object':
                sample_values[col] = df[col].dropna().sample(min(3, len(df[col]))).values
            else:
                if not df[col].empty:
                    sample_values[col] = df[col].sample(min(3, len(df[col]))).values
                else:
                    sample_values[col] = []

        # --- Print Metadata & Data Quality ---
        print(f"\n{'-'*60}\nDATA OVERVIEW: {sheet_name}\n{'-'*60}")
        print(tabulate([[k, v] for k, v in metadata.items()],
                       headers=['Metadata', 'Value'], tablefmt='pretty'))
        
        print("\n\n" + '-'*60 + "\nDATA QUALITY ASSESSMENT\n" + '-'*60)
        print(tabulate(dq, headers='keys', tablefmt='psql', showindex=True))

        print("\n\n" + '-'*60 + "\nVALUE SAMPLE\n" + '-'*60)
        print(tabulate(pd.DataFrame(sample_values), headers='keys', tablefmt='psql'))

        # --- Missing Values Plot ---
        missing_fig = px.bar(
            dq.reset_index(),
            x='index',
            y='% Missing',
            title=f'Missing Values by Column: {sheet_name}',
            labels={'index': 'Column', '% Missing': 'Percentage Missing'},
            color='% Missing',
            color_continuous_scale='Viridis'
        )
        missing_fig.update_layout(xaxis_tickangle=-45)
        missing_fig.show()

    except Exception as e:
        print(f"Error processing sheet '{sheet_name}': {e}")
