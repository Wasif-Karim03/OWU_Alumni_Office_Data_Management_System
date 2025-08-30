#!/usr/bin/env python3
"""
Test script for the Excel Visualizer
This script tests all visualization functions without displaying charts.
"""

import matplotlib
matplotlib.use('Agg')  # Use non-interactive backend for testing

from excel_visualizer import ExcelVisualizer
import os

def test_visualizer():
    """Test all visualization functions."""
    print("=== TESTING EXCEL VISUALIZER ===\n")
    
    # Initialize the visualizer
    viz = ExcelVisualizer()
    
    # Test reading Excel file
    if not os.path.exists('sample_sales_data.xlsx'):
        print("Sample data file not found. Please run create_sample_data.py first.")
        return False
    
    print("1. Testing Excel file reading...")
    if viz.read_excel('sample_sales_data.xlsx'):
        print("✓ Excel file read successfully")
    else:
        print("✗ Failed to read Excel file")
        return False
    
    # Test data info
    print("\n2. Testing data information...")
    try:
        viz.get_data_info()
        print("✓ Data info displayed successfully")
    except Exception as e:
        print(f"✗ Error displaying data info: {e}")
        return False
    
    # Test creating charts (without displaying)
    print("\n3. Testing chart creation...")
    
    try:
        # Test bar chart
        viz.create_bar_chart('Region', 'Sales', 'Test Bar Chart')
        print("✓ Bar chart created successfully")
        
        # Test line chart
        viz.create_line_chart('Date', 'Sales', 'Test Line Chart')
        print("✓ Line chart created successfully")
        
        # Test scatter plot
        viz.create_scatter_plot('Sales', 'Profit', 'Test Scatter Plot')
        print("✓ Scatter plot created successfully")
        
        # Test histogram
        viz.create_histogram('Sales', title='Test Histogram')
        print("✓ Histogram created successfully")
        
        # Test pie chart
        viz.create_pie_chart('Region', title='Test Pie Chart')
        print("✓ Pie chart created successfully")
        
        # Test box plot
        viz.create_box_plot('Sales', 'Region', 'Test Box Plot')
        print("✓ Box plot created successfully")
        
        # Test correlation heatmap
        viz.create_correlation_heatmap()
        print("✓ Correlation heatmap created successfully")
        
        # Test dashboard
        viz.create_dashboard()
        print("✓ Dashboard created successfully")
        
        # Test interactive plotly chart
        viz.create_interactive_plotly_chart('scatter', x='Sales', y='Profit')
        print("✓ Interactive plotly chart created successfully")
        
    except Exception as e:
        print(f"✗ Error creating charts: {e}")
        return False
    
    # Test export functions
    print("\n4. Testing export functions...")
    try:
        viz.export_data_summary('test_summary.txt')
        if os.path.exists('test_summary.txt'):
            print("✓ Data summary exported successfully")
            os.remove('test_summary.txt')  # Clean up
        else:
            print("✗ Data summary export failed")
            return False
    except Exception as e:
        print(f"✗ Error exporting data summary: {e}")
        return False
    
    print("\n=== ALL TESTS PASSED! ===")
    print("The Excel Visualizer is working correctly.")
    return True

if __name__ == "__main__":
    success = test_visualizer()
    if success:
        print("\n🎉 Ready to use! You can now:")
        print("1. Place your Excel files in this directory")
        print("2. Run: python excel_visualizer.py")
        print("3. Or use the functions in your own scripts")
    else:
        print("\n❌ Some tests failed. Please check the error messages above.")
