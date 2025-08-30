import pandas as pd
import numpy as np
from datetime import datetime, timedelta

def create_sample_data():
    """Create sample data for testing the Excel visualizer."""
    
    # Set random seed for reproducibility
    np.random.seed(42)
    
    # Generate sample data
    n_samples = 100
    
    # Date range
    start_date = datetime(2023, 1, 1)
    dates = [start_date + timedelta(days=i) for i in range(n_samples)]
    
    # Sample data
    data = {
        'Date': dates,
        'Sales': np.random.normal(1000, 200, n_samples),
        'Profit': np.random.normal(200, 50, n_samples),
        'Customers': np.random.poisson(50, n_samples),
        'Region': np.random.choice(['North', 'South', 'East', 'West'], n_samples),
        'Product_Category': np.random.choice(['Electronics', 'Clothing', 'Books', 'Home'], n_samples),
        'Rating': np.random.uniform(1, 5, n_samples),
        'Units_Sold': np.random.randint(10, 1000, n_samples)
    }
    
    # Create DataFrame
    df = pd.DataFrame(data)
    
    # Round numeric columns
    df['Sales'] = df['Sales'].round(2)
    df['Profit'] = df['Profit'].round(2)
    df['Rating'] = df['Rating'].round(2)
    
    # Save to Excel
    filename = 'sample_sales_data.xlsx'
    df.to_excel(filename, index=False)
    
    print(f"Sample data created and saved to {filename}")
    print(f"Data shape: {df.shape}")
    print("\nFirst few rows:")
    print(df.head())
    print("\nData types:")
    print(df.dtypes)
    print("\nSummary statistics:")
    print(df.describe())
    
    return filename

if __name__ == "__main__":
    create_sample_data()
