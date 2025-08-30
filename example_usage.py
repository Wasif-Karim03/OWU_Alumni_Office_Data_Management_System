#!/usr/bin/env python3
"""
Example usage of the Excel Data Visualizer
This script demonstrates various features and customization options.
"""

from excel_visualizer import ExcelVisualizer
import pandas as pd
import numpy as np

def create_custom_data():
    """Create custom sample data for demonstration."""
    
    # Generate sample data
    np.random.seed(123)
    n_samples = 50
    
    data = {
        'Department': np.random.choice(['Engineering', 'Marketing', 'Sales', 'HR', 'Finance'], n_samples),
        'Salary': np.random.normal(75000, 15000, n_samples),
        'Years_Experience': np.random.uniform(1, 15, n_samples),
        'Performance_Score': np.random.uniform(60, 100, n_samples),
        'Projects_Completed': np.random.poisson(8, n_samples),
        'Satisfaction_Rating': np.random.uniform(1, 5, n_samples)
    }
    
    # Create DataFrame and save to Excel
    df = pd.DataFrame(data)
    df['Salary'] = df['Salary'].round(0)
    df['Years_Experience'] = df['Years_Experience'].round(1)
    df['Performance_Score'] = df['Performance_Score'].round(1)
    df['Satisfaction_Rating'] = df['Satisfaction_Rating'].round(1)
    
    filename = 'employee_data.xlsx'
    df.to_excel(filename, index=False)
    print(f"Custom data created and saved to {filename}")
    return filename

def demonstrate_features():
    """Demonstrate various features of the Excel Visualizer."""
    
    print("=== EXCEL VISUALIZER FEATURE DEMONSTRATION ===\n")
    
    # Initialize the visualizer
    viz = ExcelVisualizer()
    
    # Create and load custom data
    data_file = create_custom_data()
    print(f"\nLoading data from {data_file}...")
    
    if viz.read_excel(data_file):
        # Display data information
        viz.get_data_info()
        
        print("\n" + "="*50)
        print("CREATING VARIOUS CHARTS AND VISUALIZATIONS")
        print("="*50)
        
        # 1. Bar Chart - Average Salary by Department
        print("\n1. Creating Bar Chart: Average Salary by Department")
        dept_salary = viz.data.groupby('Department')['Salary'].mean().sort_values(ascending=False)
        temp_df = pd.DataFrame({'Department': dept_salary.index, 'Avg_Salary': dept_salary.values})
        viz.data = temp_df  # Temporarily replace data for this chart
        viz.create_bar_chart('Department', 'Avg_Salary', 'Average Salary by Department', figsize=(10, 6))
        
        # Restore original data
        viz.read_excel(data_file)
        
        # 2. Scatter Plot - Salary vs Experience
        print("\n2. Creating Scatter Plot: Salary vs Years of Experience")
        viz.create_scatter_plot('Years_Experience', 'Salary', 'Salary vs Years of Experience')
        
        # 3. Histogram - Salary Distribution
        print("\n3. Creating Histogram: Salary Distribution")
        viz.create_histogram('Salary', bins=15, title='Salary Distribution Across All Departments')
        
        # 4. Box Plot - Performance Scores by Department
        print("\n4. Creating Box Plot: Performance Scores by Department")
        viz.create_box_plot('Performance_Score', 'Department', 'Performance Scores by Department')
        
        # 5. Pie Chart - Department Distribution
        print("\n5. Creating Pie Chart: Department Distribution")
        viz.create_pie_chart('Department', 'Employee Distribution by Department')
        
        # 6. Correlation Heatmap
        print("\n6. Creating Correlation Heatmap")
        viz.create_correlation_heatmap()
        
        # 7. Interactive Plotly Chart
        print("\n7. Creating Interactive Scatter Plot")
        viz.create_interactive_plotly_chart('scatter', 
                                         x='Years_Experience', 
                                         y='Salary',
                                         color='Department',
                                         title='Interactive: Salary vs Experience by Department')
        
        # 8. Comprehensive Dashboard
        print("\n8. Creating Comprehensive Dashboard")
        viz.create_dashboard()
        
        # 9. Export Data Summary
        print("\n9. Exporting Data Summary")
        viz.export_data_summary('employee_data_summary.txt')
        
        print("\n" + "="*50)
        print("DEMONSTRATION COMPLETE!")
        print("="*50)
        print("\nCharts have been displayed and data summary exported.")
        print("You can now explore your own Excel files using the same methods!")

def custom_analysis_example():
    """Show how to perform custom analysis and create specific visualizations."""
    
    print("\n=== CUSTOM ANALYSIS EXAMPLE ===\n")
    
    viz = ExcelVisualizer()
    viz.read_excel('employee_data.xlsx')
    
    # Custom analysis: Find top performers
    print("Custom Analysis: Top 10 Performers")
    top_performers = viz.data.nlargest(10, 'Performance_Score')[['Department', 'Salary', 'Performance_Score', 'Years_Experience']]
    print(top_performers)
    
    # Create custom visualization for top performers
    plt.figure(figsize=(12, 6))
    plt.subplot(1, 2, 1)
    plt.bar(range(len(top_performers)), top_performers['Performance_Score'])
    plt.title('Top 10 Performance Scores')
    plt.xlabel('Rank')
    plt.ylabel('Performance Score')
    
    plt.subplot(1, 2, 2)
    plt.scatter(top_performers['Years_Experience'], top_performers['Salary'])
    plt.title('Top Performers: Experience vs Salary')
    plt.xlabel('Years of Experience')
    plt.ylabel('Salary')
    
    plt.tight_layout()
    plt.show()
    
    print("\nCustom analysis complete!")

if __name__ == "__main__":
    # Run the main demonstration
    demonstrate_features()
    
    # Ask if user wants to see custom analysis
    try:
        response = input("\nWould you like to see a custom analysis example? (y/n): ").lower()
        if response in ['y', 'yes']:
            custom_analysis_example()
    except KeyboardInterrupt:
        print("\n\nDemonstration ended by user.")
    
    print("\nThank you for using the Excel Data Visualizer!")
    print("You can now use these techniques with your own Excel files.")
