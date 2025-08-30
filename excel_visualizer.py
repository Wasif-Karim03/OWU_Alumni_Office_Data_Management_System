import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import os
from pathlib import Path

class ExcelVisualizer:
    def __init__(self):
        """Initialize the Excel Visualizer with default settings."""
        self.data = None
        self.file_path = None
        
        # Set style for matplotlib
        plt.style.use('default')
        sns.set_palette("husl")
        
    def read_excel(self, file_path):
        """
        Read an Excel file and load the data.
        
        Args:
            file_path (str): Path to the Excel file
        """
        try:
            self.file_path = file_path
            self.data = pd.read_excel(file_path)
            print(f"Successfully loaded data from {file_path}")
            print(f"Data shape: {self.data.shape}")
            print(f"Columns: {list(self.data.columns)}")
            print("\nFirst few rows:")
            print(self.data.head())
            return True
        except Exception as e:
            print(f"Error reading Excel file: {e}")
            return False
    
    def get_data_info(self):
        """Display information about the loaded data."""
        if self.data is None:
            print("No data loaded. Please read an Excel file first.")
            return
        
        print("\n=== DATA INFORMATION ===")
        print(f"Shape: {self.data.shape}")
        print(f"Data types:\n{self.data.dtypes}")
        print(f"\nMissing values:\n{self.data.isnull().sum()}")
        print(f"\nNumeric columns: {list(self.data.select_dtypes(include=[np.number]).columns)}")
        print(f"Categorical columns: {list(self.data.select_dtypes(include=['object']).columns)}")
    
    def create_bar_chart(self, x_column, y_column, title="Bar Chart", figsize=(10, 6)):
        """Create a bar chart using matplotlib."""
        if self.data is None:
            print("No data loaded. Please read an Excel file first.")
            return
        
        plt.figure(figsize=figsize)
        plt.bar(self.data[x_column], self.data[y_column])
        plt.title(title)
        plt.xlabel(x_column)
        plt.ylabel(y_column)
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.show()
    
    def create_line_chart(self, x_column, y_column, title="Line Chart", figsize=(10, 6)):
        """Create a line chart using matplotlib."""
        if self.data is None:
            print("No data loaded. Please read an Excel file first.")
            return
        
        plt.figure(figsize=figsize)
        plt.plot(self.data[x_column], self.data[y_column], marker='o')
        plt.title(title)
        plt.xlabel(x_column)
        plt.ylabel(y_column)
        plt.xticks(rotation=45)
        plt.grid(True, alpha=0.3)
        plt.tight_layout()
        plt.show()
    
    def create_scatter_plot(self, x_column, y_column, title="Scatter Plot", figsize=(10, 6)):
        """Create a scatter plot using matplotlib."""
        if self.data is None:
            print("No data loaded. Please read an Excel file first.")
            return
        
        plt.figure(figsize=figsize)
        plt.scatter(self.data[x_column], self.data[y_column], alpha=0.6)
        plt.title(title)
        plt.xlabel(x_column)
        plt.ylabel(y_column)
        plt.grid(True, alpha=0.3)
        plt.tight_layout()
        plt.show()
    
    def create_histogram(self, column, bins=20, title="Histogram", figsize=(10, 6)):
        """Create a histogram using matplotlib."""
        if self.data is None:
            print("No data loaded. Please read an Excel file first.")
            return
        
        plt.figure(figsize=figsize)
        plt.hist(self.data[column].dropna(), bins=bins, alpha=0.7, edgecolor='black')
        plt.title(title)
        plt.xlabel(column)
        plt.ylabel('Frequency')
        plt.grid(True, alpha=0.3)
        plt.tight_layout()
        plt.show()
    
    def create_pie_chart(self, column, title="Pie Chart", figsize=(10, 8)):
        """Create a pie chart using matplotlib."""
        if self.data is None:
            print("No data loaded. Please read an Excel file first.")
            return
        
        value_counts = self.data[column].value_counts()
        
        plt.figure(figsize=figsize)
        plt.pie(value_counts.values, labels=value_counts.index, autopct='%1.1f%%', startangle=90)
        plt.title(title)
        plt.axis('equal')
        plt.show()
    
    def create_correlation_heatmap(self, figsize=(10, 8)):
        """Create a correlation heatmap for numeric columns."""
        if self.data is None:
            print("No data loaded. Please read an Excel file first.")
            return
        
        numeric_data = self.data.select_dtypes(include=[np.number])
        if numeric_data.empty:
            print("No numeric columns found for correlation analysis.")
            return
        
        correlation_matrix = numeric_data.corr()
        
        plt.figure(figsize=figsize)
        sns.heatmap(correlation_matrix, annot=True, cmap='coolwarm', center=0, 
                   square=True, linewidths=0.5)
        plt.title('Correlation Heatmap')
        plt.tight_layout()
        plt.show()
    
    def create_box_plot(self, column, by_column=None, title="Box Plot", figsize=(10, 6)):
        """Create a box plot using matplotlib."""
        if self.data is None:
            print("No data loaded. Please read an Excel file first.")
            return
        
        plt.figure(figsize=figsize)
        if by_column:
            self.data.boxplot(column=column, by=by_column)
            plt.title(f'{title} by {by_column}')
        else:
            self.data.boxplot(column=column)
            plt.title(title)
        plt.tight_layout()
        plt.show()
    
    def create_interactive_plotly_chart(self, chart_type='scatter', **kwargs):
        """Create an interactive chart using Plotly."""
        if self.data is None:
            print("No data loaded. Please read an Excel file first.")
            return
        
        if chart_type == 'scatter':
            fig = px.scatter(self.data, **kwargs)
        elif chart_type == 'line':
            fig = px.line(self.data, **kwargs)
        elif chart_type == 'bar':
            fig = px.bar(self.data, **kwargs)
        elif chart_type == 'histogram':
            fig = px.histogram(self.data, **kwargs)
        elif chart_type == 'box':
            fig = px.box(self.data, **kwargs)
        else:
            print(f"Unsupported chart type: {chart_type}")
            return
        
        fig.show()
    
    def create_dashboard(self, numeric_columns=None, categorical_columns=None):
        """Create a comprehensive dashboard with multiple charts."""
        if self.data is None:
            print("No data loaded. Please read an Excel file first.")
            return
        
        if numeric_columns is None:
            numeric_columns = list(self.data.select_dtypes(include=[np.number]).columns)
        if categorical_columns is None:
            categorical_columns = list(self.data.select_dtypes(include=['object']).columns)
        
        # Create subplots
        fig, axes = plt.subplots(2, 2, figsize=(15, 12))
        fig.suptitle('Data Dashboard', fontsize=16)
        
        # Plot 1: Histogram of first numeric column
        if numeric_columns:
            axes[0, 0].hist(self.data[numeric_columns[0]].dropna(), bins=20, alpha=0.7)
            axes[0, 0].set_title(f'Distribution of {numeric_columns[0]}')
            axes[0, 0].set_xlabel(numeric_columns[0])
            axes[0, 0].set_ylabel('Frequency')
        
        # Plot 2: Bar chart of first categorical column
        if categorical_columns:
            value_counts = self.data[categorical_columns[0]].value_counts().head(10)
            axes[0, 1].bar(range(len(value_counts)), value_counts.values)
            axes[0, 1].set_title(f'Top 10 values in {categorical_columns[0]}')
            axes[0, 1].set_xticks(range(len(value_counts)))
            axes[0, 1].set_xticklabels(value_counts.index, rotation=45)
        
        # Plot 3: Scatter plot of first two numeric columns
        if len(numeric_columns) >= 2:
            axes[1, 0].scatter(self.data[numeric_columns[0]], self.data[numeric_columns[1]], alpha=0.6)
            axes[1, 0].set_title(f'{numeric_columns[0]} vs {numeric_columns[1]}')
            axes[1, 0].set_xlabel(numeric_columns[0])
            axes[1, 0].set_ylabel(numeric_columns[1])
        
        # Plot 4: Box plot
        if numeric_columns:
            axes[1, 1].boxplot(self.data[numeric_columns[0]].dropna())
            axes[1, 1].set_title(f'Box Plot of {numeric_columns[0]}')
        
        plt.tight_layout()
        plt.show()
    
    def save_chart(self, filename="chart.png", dpi=300):
        """Save the current matplotlib figure."""
        plt.savefig(filename, dpi=dpi, bbox_inches='tight')
        print(f"Chart saved as {filename}")
    
    def export_data_summary(self, filename="data_summary.txt"):
        """Export a summary of the data to a text file."""
        if self.data is None:
            print("No data loaded. Please read an Excel file first.")
            return
        
        with open(filename, 'w') as f:
            f.write("DATA SUMMARY REPORT\n")
            f.write("=" * 50 + "\n\n")
            f.write(f"File: {self.file_path}\n")
            f.write(f"Shape: {self.data.shape}\n\n")
            f.write("COLUMNS:\n")
            f.write("-" * 20 + "\n")
            for col in self.data.columns:
                f.write(f"{col}: {self.data[col].dtype}\n")
            f.write(f"\nMissing values:\n{self.data.isnull().sum()}\n")
            f.write(f"\nNumeric columns: {list(self.data.select_dtypes(include=[np.number]).columns)}\n")
            f.write(f"Categorical columns: {list(self.data.select_dtypes(include=['object']).columns)}\n")
        
        print(f"Data summary exported to {filename}")

def main():
    """Main function to demonstrate the Excel Visualizer."""
    print("=== EXCEL DATA VISUALIZER ===\n")
    
    # Initialize the visualizer
    viz = ExcelVisualizer()
    
    # Check if there are Excel files in the current directory
    excel_files = list(Path('.').glob('*.xlsx')) + list(Path('.').glob('*.xls'))
    
    if excel_files:
        print("Found Excel files:")
        for i, file in enumerate(excel_files):
            print(f"{i+1}. {file}")
        
        # Try to read the first Excel file
        if excel_files:
            print(f"\nReading {excel_files[0]}...")
            if viz.read_excel(str(excel_files[0])):
                viz.get_data_info()
                
                # Get column names for visualization
                numeric_cols = list(viz.data.select_dtypes(include=[np.number]).columns)
                categorical_cols = list(viz.data.select_dtypes(include=['object']).columns)
                
                print(f"\nNumeric columns: {numeric_cols}")
                print(f"Categorical columns: {categorical_cols}")
                
                # Create some example visualizations if we have data
                if numeric_cols and categorical_cols:
                    print("\nCreating example visualizations...")
                    
                    # Bar chart
                    if len(categorical_cols) > 0 and len(numeric_cols) > 0:
                        viz.create_bar_chart(categorical_cols[0], numeric_cols[0], 
                                           f"{categorical_cols[0]} vs {numeric_cols[0]}")
                    
                    # Histogram
                    if numeric_cols:
                        viz.create_histogram(numeric_cols[0], title=f"Distribution of {numeric_cols[0]}")
                    
                    # Pie chart
                    if categorical_cols:
                        viz.create_pie_chart(categorical_cols[0], title=f"Distribution of {categorical_cols[0]}")
                    
                    # Dashboard
                    viz.create_dashboard()
                    
                    # Export summary
                    viz.export_data_summary()
                else:
                    print("Not enough columns of different types for visualization.")
    else:
        print("No Excel files found in the current directory.")
        print("Please place an Excel file (.xlsx or .xls) in this directory and run the script again.")
        print("\nExample usage:")
        print("1. Place your Excel file in this directory")
        print("2. Run: python excel_visualizer.py")
        print("3. The script will automatically detect and visualize your data")

if __name__ == "__main__":
    main()
