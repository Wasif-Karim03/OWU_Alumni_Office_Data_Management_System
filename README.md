# Excel Data Visualizer

A powerful Python application that can read Excel files and create various types of data visualizations including charts, graphs, and interactive dashboards.

## Features

- **Excel File Reading**: Supports both `.xlsx` and `.xls` formats
- **Multiple Chart Types**: Bar charts, line charts, scatter plots, histograms, pie charts, box plots
- **Interactive Visualizations**: Plotly-based interactive charts
- **Data Analysis**: Correlation heatmaps, data summaries, and statistics
- **Dashboard Creation**: Multi-panel dashboards with various chart types
- **Export Capabilities**: Save charts as images and export data summaries

## Installation

### Prerequisites
- Python 3.7 or higher
- pip (Python package installer)

### Step 1: Install Dependencies
```bash
pip install -r requirements.txt
```

### Step 2: Verify Installation
```bash
python -c "import pandas, matplotlib, seaborn, plotly; print('All packages installed successfully!')"
```

## Quick Start

### 1. Generate Sample Data (Optional)
If you don't have an Excel file to test with, create sample data:
```bash
python create_sample_data.py
```

### 2. Run the Visualizer
```bash
python excel_visualizer.py
```

The script will automatically:
- Detect Excel files in the current directory
- Load the first Excel file found
- Display data information
- Create example visualizations
- Generate a comprehensive dashboard

## Usage Examples

### Basic Usage
```python
from excel_visualizer import ExcelVisualizer

# Initialize the visualizer
viz = ExcelVisualizer()

# Read an Excel file
viz.read_excel('your_data.xlsx')

# Get information about the data
viz.get_data_info()

# Create various charts
viz.create_bar_chart('Category', 'Value', 'My Bar Chart')
viz.create_histogram('Numeric_Column', title='Distribution')
viz.create_scatter_plot('X_Column', 'Y_Column')
```

### Advanced Usage
```python
# Create a correlation heatmap
viz.create_correlation_heatmap()

# Create a comprehensive dashboard
viz.create_dashboard()

# Create interactive Plotly charts
viz.create_interactive_plotly_chart('scatter', x='X_Column', y='Y_Column')

# Export data summary
viz.export_data_summary('my_summary.txt')
```

## Chart Types Available

### 1. **Bar Charts** (`create_bar_chart`)
- Perfect for categorical vs numerical data
- Customizable titles and axis labels
- Automatic rotation for long labels

### 2. **Line Charts** (`create_line_chart`)
- Ideal for time series data
- Includes markers and grid lines
- Great for trends and patterns

### 3. **Scatter Plots** (`create_scatter_plot`)
- Shows relationships between two variables
- Alpha transparency for overlapping points
- Grid lines for better readability

### 4. **Histograms** (`create_histogram`)
- Displays data distribution
- Configurable bin count
- Handles missing data automatically

### 5. **Pie Charts** (`create_pie_chart`)
- Shows proportions of categorical data
- Percentage labels
- Professional appearance

### 6. **Box Plots** (`create_box_plot`)
- Statistical summary of numerical data
- Shows median, quartiles, and outliers
- Can be grouped by categorical variables

### 7. **Correlation Heatmaps** (`create_correlation_heatmap`)
- Matrix visualization of correlations
- Color-coded for easy interpretation
- Perfect for identifying relationships

### 8. **Interactive Charts** (`create_interactive_plotly_chart`)
- Hover information
- Zoom and pan capabilities
- Export to HTML

## Data Requirements

The visualizer works best with data that has:
- **Numeric columns**: For quantitative analysis and most chart types
- **Categorical columns**: For grouping and comparison
- **Clean data**: Minimal missing values for best results

## File Structure

```
excel_visualizer/
├── excel_visualizer.py      # Main application
├── create_sample_data.py    # Sample data generator
├── requirements.txt         # Python dependencies
├── README.md               # This file
└── sample_sales_data.xlsx  # Sample data (after running create_sample_data.py)
```

## Customization

### Chart Styling
```python
# Custom figure sizes
viz.create_bar_chart('Category', 'Value', figsize=(12, 8))

# Custom titles and labels
viz.create_scatter_plot('X', 'Y', title='Custom Title')
```

### Saving Charts
```python
# Save with custom filename and resolution
viz.save_chart('my_chart.png', dpi=300)
```

## Troubleshooting

### Common Issues

1. **"No module named 'pandas'"**
   - Solution: Run `pip install -r requirements.txt`

2. **Excel file not found**
   - Ensure the Excel file is in the same directory as the script
   - Check file extensions (.xlsx or .xls)

3. **Charts not displaying**
   - Make sure you're running in an environment that supports GUI (not headless)
   - For Jupyter notebooks, add `%matplotlib inline`

4. **Memory issues with large files**
   - Consider sampling your data first
   - Use `nrows` parameter in `pd.read_excel()` for large files

### Performance Tips

- For large datasets (>10,000 rows), consider sampling
- Use specific column selection when reading Excel files
- Close matplotlib figures when done to free memory

## Advanced Features

### Custom Dashboard Creation
```python
# Create dashboard with specific columns
viz.create_dashboard(
    numeric_columns=['Sales', 'Profit'],
    categorical_columns=['Region', 'Product_Category']
)
```

### Data Export
```python
# Export comprehensive data summary
viz.export_data_summary('detailed_report.txt')
```

## Contributing

Feel free to extend this application with:
- Additional chart types
- More export formats
- Interactive features
- Performance optimizations

## License

This project is open source and available under the MIT License.

## Support

For issues or questions:
1. Check the troubleshooting section
2. Verify your Python environment and dependencies
3. Ensure your Excel file format is supported

---

**Happy Data Visualization! 📊📈**
