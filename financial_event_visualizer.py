#!/usr/bin/env python3
"""
Financial Event Data Visualizer
Specialized for income/expense snapshots across multiple events
"""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from pathlib import Path

class FinancialEventVisualizer:
    def __init__(self):
        """Initialize the Financial Event Visualizer."""
        self.data = None
        self.file_path = None
        
        # Set style for matplotlib
        plt.style.use('default')
        sns.set_palette("husl")
        
        # Financial color scheme
        self.colors = {
            'income': '#2E8B57',      # Sea Green
            'expenses': '#DC143C',    # Crimson
            'underwritten': '#FF8C00', # Dark Orange
            'profit': '#32CD32',      # Lime Green
            'loss': '#FF0000'         # Red
        }
    
    def read_excel(self, file_path):
        """Read an Excel file and load the financial event data."""
        try:
            self.file_path = file_path
            self.data = pd.read_excel(file_path)
            print(f"Successfully loaded financial data from {file_path}")
            print(f"Data shape: {self.data.shape}")
            print(f"Columns: {list(self.data.columns)}")
            print("\nFirst few rows:")
            print(self.data.head())
            return True
        except Exception as e:
            print(f"Error reading Excel file: {e}")
            return False
    
    def create_sample_financial_data(self):
        """Create sample financial event data similar to the image."""
        events = [
            "07/27/2024 Football Golf Outing",
            "8/02/2024 Clippers - Clipper Stadium", 
            "8/15/2024 Legacy Reception",
            "8/21/2024 Senior Class Welcome Back Toast",
            "8/25/2024 Monnett Club - Columbus Library",
            "9/11/2024 OWU Near You - Toledo",
            "9/11/2024 OWU Near You - Denver",
            "9/12/2024 OWU Near You - Tucson",
            "9/12/2024 OWU Near You - Cleveland",
            "9/12/2024 OWU Near You - Atlanta",
            "9/12/2024 OWU Near You - Chicago",
            "9/13/2024 OWU Near You - Cincinnati"
        ]
        
        # Create sample data
        data = {
            'Event': events,
            'Event Income': [0, 350, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
            'All Incurred Expenses': [0, 500, 501, 275, 0, 114, 14, 14, 14, 14, 14, 14],
            'Underwritten': [0, 0, 0, 0, 0, 0, 700, 0, 0, 0, 0, 0],
            'Profit/Loss': [0, -150, 0, 0, 0, 0, -14, -14, -14, -14, -14, -14]
        }
        
        df = pd.DataFrame(data)
        
        # Save to Excel
        filename = 'financial_events_sample.xlsx'
        df.to_excel(filename, index=False)
        print(f"Sample financial event data created and saved to {filename}")
        return filename
    
    def create_profit_loss_chart(self, figsize=(14, 8)):
        """Create a bar chart showing profit/loss for each event."""
        if self.data is None:
            print("No data loaded. Please read an Excel file first.")
            return
        
        plt.figure(figsize=figsize)
        
        # Create profit/loss bars
        bars = plt.bar(range(len(self.data)), self.data['Profit/Loss'], 
                      color=['red' if x < 0 else 'green' for x in self.data['Profit/Loss']],
                      alpha=0.7)
        
        plt.title('Event Profit/Loss Analysis', fontsize=16, fontweight='bold')
        plt.xlabel('Events', fontsize=12)
        plt.ylabel('Profit/Loss ($)', fontsize=12)
        
        # Set x-axis labels
        plt.xticks(range(len(self.data)), self.data['Event'], rotation=45, ha='right')
        
        # Add value labels on bars
        for i, bar in enumerate(bars):
            height = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2., height,
                    f'${height:,.0f}', ha='center', va='bottom' if height > 0 else 'top')
        
        # Add horizontal line at zero
        plt.axhline(y=0, color='black', linestyle='-', alpha=0.3)
        
        plt.grid(True, alpha=0.3)
        plt.tight_layout()
        plt.show()
    
    def create_income_vs_expenses_chart(self, figsize=(14, 8)):
        """Create a grouped bar chart comparing income vs expenses."""
        if self.data is None:
            print("No data loaded. Please read an Excel file first.")
            return
        
        x = np.arange(len(self.data))
        width = 0.35
        
        fig, ax = plt.subplots(figsize=figsize)
        
        # Create bars
        bars1 = ax.bar(x - width/2, self.data['Event Income'], width, 
                      label='Event Income', color=self.colors['income'], alpha=0.8)
        bars2 = ax.bar(x + width/2, self.data['All Incurred Expenses'], width, 
                      label='Expenses', color=self.colors['expenses'], alpha=0.8)
        
        ax.set_xlabel('Events', fontsize=12)
        ax.set_ylabel('Amount ($)', fontsize=12)
        ax.set_title('Event Income vs Expenses Comparison', fontsize=16, fontweight='bold')
        ax.set_xticks(x)
        ax.set_xticklabels(self.data['Event'], rotation=45, ha='right')
        ax.legend()
        
        # Add value labels
        for bars in [bars1, bars2]:
            for bar in bars:
                height = bar.get_height()
                if height > 0:
                    ax.text(bar.get_x() + bar.get_width()/2., height,
                           f'${height:,.0f}', ha='center', va='bottom')
        
        plt.grid(True, alpha=0.3)
        plt.tight_layout()
        plt.show()
    
    def create_underwriting_analysis(self, figsize=(12, 6)):
        """Create a chart showing underwriting amounts."""
        if self.data is None:
            print("No data loaded. Please read an Excel file first.")
            return
        
        # Filter events with underwriting
        underwritten_events = self.data[self.data['Underwritten'] > 0]
        
        if underwritten_events.empty:
            print("No underwriting data found.")
            return
        
        plt.figure(figsize=figsize)
        bars = plt.bar(range(len(underwritten_events)), underwritten_events['Underwritten'],
                      color=self.colors['underwritten'], alpha=0.8)
        
        plt.title('Event Underwriting Analysis', fontsize=16, fontweight='bold')
        plt.xlabel('Events', fontsize=12)
        plt.ylabel('Underwritten Amount ($)', fontsize=12)
        
        plt.xticks(range(len(underwritten_events)), underwritten_events['Event'], 
                  rotation=45, ha='right')
        
        # Add value labels
        for i, bar in enumerate(bars):
            height = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2., height,
                    f'${height:,.0f}', ha='center', va='bottom')
        
        plt.grid(True, alpha=0.3)
        plt.tight_layout()
        plt.show()
    
    def create_financial_summary_dashboard(self, figsize=(16, 12)):
        """Create a comprehensive financial dashboard."""
        if self.data is None:
            print("No data loaded. Please read an Excel file first.")
            return
        
        fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=figsize)
        fig.suptitle('Financial Events Dashboard', fontsize=18, fontweight='bold')
        
        # 1. Profit/Loss Distribution
        profit_loss = self.data['Profit/Loss']
        colors = ['red' if x < 0 else 'green' for x in profit_loss]
        ax1.bar(range(len(profit_loss)), profit_loss, color=colors, alpha=0.7)
        ax1.set_title('Profit/Loss by Event')
        ax1.set_ylabel('Amount ($)')
        ax1.axhline(y=0, color='black', linestyle='-', alpha=0.3)
        ax1.grid(True, alpha=0.3)
        
        # 2. Income vs Expenses
        x = np.arange(len(self.data))
        width = 0.35
        ax2.bar(x - width/2, self.data['Event Income'], width, 
                label='Income', color=self.colors['income'], alpha=0.8)
        ax2.bar(x + width/2, self.data['All Incurred Expenses'], width, 
                label='Expenses', color=self.colors['expenses'], alpha=0.8)
        ax2.set_title('Income vs Expenses')
        ax2.set_ylabel('Amount ($)')
        ax2.legend()
        ax2.grid(True, alpha=0.3)
        
        # 3. Total Financial Summary
        total_income = self.data['Event Income'].sum()
        total_expenses = self.data['All Incurred Expenses'].sum()
        total_underwritten = self.data['Underwritten'].sum()
        net_result = self.data['Profit/Loss'].sum()
        
        summary_data = ['Total Income', 'Total Expenses', 'Total Underwritten', 'Net Result']
        summary_values = [total_income, total_expenses, total_underwritten, net_result]
        summary_colors = [self.colors['income'], self.colors['expenses'], 
                         self.colors['underwritten'], 
                         self.colors['profit'] if net_result >= 0 else self.colors['loss']]
        
        bars = ax3.bar(summary_data, summary_values, color=summary_colors, alpha=0.8)
        ax3.set_title('Overall Financial Summary')
        ax3.set_ylabel('Amount ($)')
        
        # Add value labels
        for bar in bars:
            height = bar.get_height()
            ax3.text(bar.get_x() + bar.get_width()/2., height,
                    f'${height:,.0f}', ha='center', va='bottom')
        
        # 4. Event Type Analysis (OWU Near You vs Others)
        owu_events = self.data[self.data['Event'].str.contains('OWU Near You')]
        other_events = self.data[~self.data['Event'].str.contains('OWU Near You')]
        
        owu_avg_expense = owu_events['All Incurred Expenses'].mean()
        other_avg_expense = other_events['All Incurred Expenses'].mean()
        
        categories = ['OWU Near You Events', 'Other Events']
        avg_expenses = [owu_avg_expense, other_avg_expense]
        
        ax4.bar(categories, avg_expenses, color=[self.colors['expenses'], self.colors['underwritten']], alpha=0.8)
        ax4.set_title('Average Expenses by Event Type')
        ax4.set_ylabel('Average Amount ($)')
        
        # Add value labels
        for i, v in enumerate(avg_expenses):
            ax4.text(i, v, f'${v:.0f}', ha='center', va='bottom')
        
        plt.tight_layout()
        plt.show()
    
    def create_interactive_financial_chart(self):
        """Create an interactive Plotly chart for financial analysis."""
        if self.data is None:
            print("No data loaded. Please read an Excel file first.")
            return
        
        # Create interactive scatter plot
        fig = px.scatter(self.data, 
                        x='All Incurred Expenses', 
                        y='Event Income',
                        size='Underwritten',
                        color='Profit/Loss',
                        hover_data=['Event'],
                        title='Interactive Financial Analysis: Expenses vs Income',
                        labels={'All Incurred Expenses': 'Expenses ($)', 
                               'Event Income': 'Income ($)',
                               'Underwritten': 'Underwritten Amount',
                               'Profit/Loss': 'Profit/Loss ($)'})
        
        # Add hover template
        fig.update_traces(hovertemplate='<b>%{customdata[0]}</b><br>' +
                                       'Expenses: $%{x:,.0f}<br>' +
                                       'Income: $%{y:,.0f}<br>' +
                                       'Underwritten: $%{marker.size:,.0f}<br>' +
                                       'Profit/Loss: $%{marker.color:,.0f}<extra></extra>')
        
        fig.show()
    
    def export_financial_summary(self, filename="financial_summary_report.txt"):
        """Export a comprehensive financial summary report."""
        if self.data is None:
            print("No data loaded. Please read an Excel file first.")
            return
        
        with open(filename, 'w') as f:
            f.write("FINANCIAL EVENTS SUMMARY REPORT\n")
            f.write("=" * 50 + "\n\n")
            f.write(f"File: {self.file_path}\n")
            f.write(f"Total Events: {len(self.data)}\n\n")
            
            f.write("FINANCIAL SUMMARY:\n")
            f.write("-" * 20 + "\n")
            f.write(f"Total Event Income: ${self.data['Event Income'].sum():,.2f}\n")
            f.write(f"Total Expenses: ${self.data['All Incurred Expenses'].sum():,.2f}\n")
            f.write(f"Total Underwritten: ${self.data['Underwritten'].sum():,.2f}\n")
            f.write(f"Net Result: ${self.data['Profit/Loss'].sum():,.2f}\n\n")
            
            f.write("EVENT ANALYSIS:\n")
            f.write("-" * 20 + "\n")
            for _, row in self.data.iterrows():
                f.write(f"Event: {row['Event']}\n")
                f.write(f"  Income: ${row['Event Income']:,.2f}\n")
                f.write(f"  Expenses: ${row['All Incurred Expenses']:,.2f}\n")
                f.write(f"  Underwritten: ${row['Underwritten']:,.2f}\n")
                f.write(f"  Profit/Loss: ${row['Profit/Loss']:,.2f}\n\n")
            
            # Find best and worst performing events
            best_event = self.data.loc[self.data['Profit/Loss'].idxmax()]
            worst_event = self.data.loc[self.data['Profit/Loss'].idxmin()]
            
            f.write("PERFORMANCE HIGHLIGHTS:\n")
            f.write("-" * 20 + "\n")
            f.write(f"Best Performing Event: {best_event['Event']} (${best_event['Profit/Loss']:,.2f})\n")
            f.write(f"Worst Performing Event: {worst_event['Event']} (${worst_event['Profit/Loss']:,.2f})\n")
        
        print(f"Financial summary exported to {filename}")

def main():
    """Main function to demonstrate the Financial Event Visualizer."""
    print("=== FINANCIAL EVENT DATA VISUALIZER ===\n")
    
    # Initialize the visualizer
    viz = FinancialEventVisualizer()
    
    # Check if there are Excel files in the current directory
    excel_files = list(Path('.').glob('*.xlsx')) + list(Path('.').glob('*.xls'))
    
    if excel_files:
        print("Found Excel files:")
        for i, file in enumerate(excel_files):
            print(f"{i+1}. {file}")
        
        # Try to read the first Excel file and check if it's financial data
        if excel_files:
            print(f"\nReading {excel_files[0]}...")
            if viz.read_excel(str(excel_files[0])):
                # Check if this is financial event data
                required_columns = ['Event Income', 'All Incurred Expenses', 'Underwritten', 'Profit/Loss']
                if all(col in viz.data.columns for col in required_columns):
                    print("\nFinancial event data detected! Creating visualizations...")
                    
                    # Create various financial charts
                    viz.create_profit_loss_chart()
                    viz.create_income_vs_expenses_chart()
                    viz.create_underwriting_analysis()
                    viz.create_financial_summary_dashboard()
                    viz.create_interactive_financial_chart()
                    
                    # Export summary
                    viz.export_financial_summary()
                    
                    print("\nFinancial analysis complete!")
                else:
                    print("\nThis Excel file doesn't contain financial event data.")
                    print("Creating sample financial data instead...")
                    sample_file = viz.create_sample_financial_data()
                    
                    print(f"\nReading sample data from {sample_file}...")
                    if viz.read_excel(sample_file):
                        print("\nCreating sample financial visualizations...")
                        
                        # Create various financial charts
                        viz.create_profit_loss_chart()
                        viz.create_income_vs_expenses_chart()
                        viz.create_underwriting_analysis()
                        viz.create_financial_summary_dashboard()
                        viz.create_interactive_financial_chart()
                        
                        # Export summary
                        viz.export_financial_summary()
                        
                        print("\nSample financial analysis complete!")
                        print("You can now use this with your own financial event data!")
    else:
        print("No Excel files found. Creating sample financial data...")
        sample_file = viz.create_sample_financial_data()
        
        print(f"\nReading sample data from {sample_file}...")
        if viz.read_excel(sample_file):
            print("\nCreating sample financial visualizations...")
            
            # Create various financial charts
            viz.create_profit_loss_chart()
            viz.create_income_vs_expenses_chart()
            viz.create_underwriting_analysis()
            viz.create_financial_summary_dashboard()
            viz.create_interactive_financial_chart()
            
            # Export summary
            viz.export_financial_summary()
            
            print("\nSample financial analysis complete!")
            print("You can now use this with your own financial event data!")

if __name__ == "__main__":
    main()
