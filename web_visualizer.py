#!/usr/bin/env python3
"""
Web-based Excel Data Visualizer
Flask application with modern UI for uploading Excel files and viewing visualizations
"""

from flask import Flask, render_template, request, jsonify, send_file, session, redirect, url_for
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.utils
import json
import os
import uuid
from werkzeug.utils import secure_filename
import numpy as np
from pathlib import Path
import base64
import io
from datetime import datetime
import openpyxl

app = Flask(__name__)
app.secret_key = 'owu_alumni_secret_key_2025'  # Secret key for sessions
app.config['SECRET_KEY'] = 'your-secret-key-here'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Ensure upload directory exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Allowed file extensions
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv'}

def allowed_file(filename):
    """Check if file extension is allowed."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def analyze_data(df):
    """Analyze the uploaded data and determine the best visualization approach."""
    # Convert data types to strings for JSON serialization
    data_types_dict = {}
    for col, dtype in df.dtypes.items():
        data_types_dict[str(col)] = str(dtype)
    
    # Convert missing values to regular Python types
    missing_values_dict = {}
    for col, count in df.isnull().sum().items():
        missing_values_dict[str(col)] = int(count)
    
    # Safely convert sample data to JSON-serializable format
    sample_data = []
    for idx, row in df.head(10).iterrows():
        row_dict = {}
        for col in df.columns:
            value = row[col]
            if pd.isna(value):
                row_dict[str(col)] = None
            elif isinstance(value, (np.integer, np.floating)):
                row_dict[str(col)] = float(value) if isinstance(value, np.floating) else int(value)
            else:
                row_dict[str(col)] = str(value)
        sample_data.append(row_dict)
    
    analysis = {
        'shape': [int(df.shape[0]), int(df.shape[1])],
        'columns': [str(col) for col in df.columns],
        'data_types': data_types_dict,
        'numeric_columns': [str(col) for col in df.select_dtypes(include=[np.number]).columns],
        'categorical_columns': [str(col) for col in df.select_dtypes(include=['object']).columns],
        'datetime_columns': [str(col) for col in df.select_dtypes(include=['datetime64']).columns],
        'missing_values': missing_values_dict,
        'sample_data': sample_data
    }
    
    # Determine if this is financial data
    financial_keywords = ['income', 'expense', 'profit', 'loss', 'revenue', 'cost', 'amount', 'price', 'sales']
    is_financial = any(keyword in ' '.join(df.columns).lower() for keyword in financial_keywords)
    analysis['is_financial'] = bool(is_financial)
    
    return analysis

def create_visualizations(df, analysis):
    """Create various visualizations based on the data analysis."""
    charts = {}
    
    try:
        # 1. Basic Data Overview Chart
        if analysis['numeric_columns']:
            # Create a summary chart of numeric columns
            numeric_summary = df[analysis['numeric_columns']].describe()
            fig = px.bar(x=numeric_summary.columns, y=numeric_summary.loc['mean'],
                        title='Average Values by Column',
                        labels={'x': 'Columns', 'y': 'Average Value'})
            charts['numeric_summary'] = json.dumps(fig, cls=plotly.utils.PlotlyJSONEncoder)
        
        # 2. Correlation Heatmap (if multiple numeric columns)
        if len(analysis['numeric_columns']) > 1:
            correlation_matrix = df[analysis['numeric_columns']].corr()
            fig = px.imshow(correlation_matrix,
                           title='Correlation Heatmap',
                           color_continuous_scale='RdBu',
                           aspect='auto')
            charts['correlation'] = json.dumps(fig, cls=plotly.utils.PlotlyJSONEncoder)
        
        # 3. Distribution Charts for Numeric Columns
        if analysis['numeric_columns']:
            for col in analysis['numeric_columns'][:3]:  # Limit to first 3 columns
                fig = px.histogram(df, x=col, title=f'Distribution of {col}')
                charts[f'distribution_{col}'] = json.dumps(fig, cls=plotly.utils.PlotlyJSONEncoder)
        
        # 4. Categorical Analysis
        if analysis['categorical_columns']:
            for col in analysis['categorical_columns'][:2]:  # Limit to first 2 columns
                value_counts = df[col].value_counts().head(10)
                fig = px.bar(x=value_counts.values, y=value_counts.index,
                            orientation='h', title=f'Top 10 Values in {col}')
                charts[f'categorical_{col}'] = json.dumps(fig, cls=plotly.utils.PlotlyJSONEncoder)
        
        # 5. Financial-specific charts (if applicable)
        if analysis['is_financial']:
            # Look for common financial column patterns
            income_cols = [col for col in df.columns if 'income' in col.lower() or 'revenue' in col.lower()]
            expense_cols = [col for col in df.columns if 'expense' in col.lower() or 'cost' in col.lower()]
            
            if income_cols and expense_cols:
                # Income vs Expenses comparison
                fig = px.scatter(df, x=expense_cols[0], y=income_cols[0],
                               title=f'{income_cols[0]} vs {expense_cols[0]}',
                               labels={expense_cols[0]: 'Expenses', income_cols[0]: 'Income'})
                charts['income_vs_expenses'] = json.dumps(fig, cls=plotly.utils.PlotlyJSONEncoder)
        
        # 6. Time Series (if datetime columns exist)
        if analysis['datetime_columns']:
            for col in analysis['datetime_columns'][:1]:  # Limit to first datetime column
                if analysis['numeric_columns']:
                    # Create time series plot with first numeric column
                    numeric_col = analysis['numeric_columns'][0]
                    fig = px.line(df, x=col, y=numeric_col,
                                title=f'{numeric_col} Over Time')
                    charts['time_series'] = json.dumps(fig, cls=plotly.utils.PlotlyJSONEncoder)
        
        # 7. Scatter Plot Matrix (if multiple numeric columns)
        if len(analysis['numeric_columns']) >= 2:
            fig = px.scatter_matrix(df[analysis['numeric_columns'][:3]],
                                  title='Scatter Plot Matrix')
            charts['scatter_matrix'] = json.dumps(fig, cls=plotly.utils.PlotlyJSONEncoder)
        
    except Exception as e:
        print(f"Error creating visualizations: {e}")
        charts['error'] = str(e)
    
    return charts

def require_auth(f):
    """Decorator to require authentication for protected routes."""
    def decorated_function(*args, **kwargs):
        if not session.get('authenticated'):
            return redirect(url_for('signin'))
        return f(*args, **kwargs)
    decorated_function.__name__ = f.__name__
    return decorated_function

@app.route('/')
def landing():
    """Landing page with sign-in options."""
    return render_template('landing.html')

@app.route('/visualizer')
@require_auth
def index():
    """Main page with file upload form."""
    return render_template('index.html')

@app.route('/signin', methods=['GET', 'POST'])
def signin():
    """Sign-in page for user authentication."""
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        
        # Check credentials
        if username == 'alumniadmin' and password == 'alumniadmin2025':
            session['authenticated'] = True
            session['username'] = username
            return redirect(url_for('index'))
        else:
            return render_template('signin.html', error='Invalid username or password')
    
    return render_template('signin.html')

@app.route('/logout')
def logout():
    """Logout user and clear session."""
    session.clear()
    return redirect(url_for('landing'))

@app.route('/demo')
def demo():
    """Demo page - redirect to main visualizer."""
    return render_template('index.html')

@app.route('/about')
def about():
    """About page with Ohio Wesleyan University information."""
    return jsonify({
        'university': 'Ohio Wesleyan University',
        'location': 'Delaware, Ohio',
        'tool': 'Excel Data Visualizer',
        'description': 'Professional data analysis tool for the Ohio Wesleyan University community'
    })

@app.route('/data-management')
@require_auth
def data_management():
    """Data management page for manually adding and editing event data."""
    return render_template('data_management.html')

@app.route('/api/files', methods=['GET'])
def get_files():
    """Get list of all event files."""
    try:
        files = []
        if os.path.exists('event_data'):
            for filename in os.listdir('event_data'):
                if filename.endswith('.json'):
                    file_path = os.path.join('event_data', filename)
                    with open(file_path, 'r') as f:
                        data = json.load(f)
                        files.append({
                            'id': filename.replace('.json', ''),
                            'name': data.get('name', filename),
                            'event_count': len(data.get('events', [])),
                            'created_date': data.get('created_date', ''),
                            'last_modified': data.get('last_modified', '')
                        })
        return jsonify({'success': True, 'files': files})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/files', methods=['POST'])
def create_file():
    """Create a new event file."""
    try:
        data = request.json
        file_name = data.get('name', 'New Event File')
        
        # Create event_data directory if it doesn't exist
        os.makedirs('event_data', exist_ok=True)
        
        # Generate unique file ID
        file_id = f"owu_events_{int(datetime.now().timestamp())}"
        
        # Create file structure
        file_data = {
            'id': file_id,
            'name': file_name,
            'created_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'last_modified': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'events': []
        }
        
        # Save file
        file_path = os.path.join('event_data', f"{file_id}.json")
        with open(file_path, 'w') as f:
            json.dump(file_data, f, indent=2)
        
        return jsonify({'success': True, 'file_id': file_id, 'message': 'File created successfully'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/files/<file_id>', methods=['GET'])
def get_file(file_id):
    """Get specific event file data."""
    try:
        file_path = os.path.join('event_data', f"{file_id}.json")
        if os.path.exists(file_path):
            with open(file_path, 'r') as f:
                data = json.load(f)
            return jsonify({'success': True, 'data': data})
        else:
            return jsonify({'error': 'File not found'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/files/<file_id>/events', methods=['POST'])
def add_event(file_id):
    """Add a new event to a file."""
    try:
        data = request.json
        file_path = os.path.join('event_data', f"{file_id}.json")
        
        if not os.path.exists(file_path):
            return jsonify({'error': 'File not found'}), 404
        
        with open(file_path, 'r') as f:
            file_data = json.load(f)
        
        # Create new event
        new_event = {
            'id': f"event_{len(file_data['events']) + 1}",
            'date': data.get('date', ''),
            'name': data.get('name', ''),
            'income': float(data.get('income', 0)),
            'expenses': float(data.get('expenses', 0)),
            'underwritten': float(data.get('underwritten', 0)),
            'profit_loss': float(data.get('income', 0)) - float(data.get('expenses', 0)) + float(data.get('underwritten', 0)),
            'created_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
        
        file_data['events'].append(new_event)
        file_data['last_modified'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        # Save updated file
        with open(file_path, 'w') as f:
            json.dump(file_data, f, indent=2)
        
        return jsonify({'success': True, 'event': new_event, 'message': 'Event added successfully'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/files/<file_id>/events/<event_id>', methods=['PUT'])
def update_event(file_id, event_id):
    """Update an existing event."""
    try:
        data = request.json
        file_path = os.path.join('event_data', f"{file_id}.json")
        
        if not os.path.exists(file_path):
            return jsonify({'error': 'File not found'}), 404
        
        with open(file_path, 'r') as f:
            file_data = json.load(f)
        
        # Find and update event
        for event in file_data['events']:
            if event['id'] == event_id:
                event['date'] = data.get('date', event['date'])
                event['name'] = data.get('name', event['name'])
                event['income'] = float(data.get('income', event['income']))
                event['expenses'] = float(data.get('expenses', event['expenses']))
                event['underwritten'] = float(data.get('underwritten', event['underwritten']))
                event['profit_loss'] = event['income'] - event['expenses'] + event['underwritten']
                event['last_modified'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                break
        
        file_data['last_modified'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        # Save updated file
        with open(file_path, 'w') as f:
            json.dump(file_data, f, indent=2)
        
        return jsonify({'success': True, 'message': 'Event updated successfully'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/files/<file_id>/events/<event_id>', methods=['DELETE'])
def delete_event(file_id, event_id):
    """Delete an event from a file."""
    try:
        file_path = os.path.join('event_data', f"{file_id}.json")
        
        if not os.path.exists(file_path):
            return jsonify({'error': 'File not found'}), 404
        
        with open(file_path, 'r') as f:
            file_data = json.load(f)
        
        # Remove event
        file_data['events'] = [event for event in file_data['events'] if event['id'] != event_id]
        file_data['last_modified'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        # Save updated file
        with open(file_path, 'w') as f:
            json.dump(file_data, f, indent=2)
        
        return jsonify({'success': True, 'message': 'Event deleted successfully'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/files/<file_id>', methods=['DELETE'])
def delete_file(file_id):
    """Delete an entire event file."""
    try:
        file_path = os.path.join('event_data', f"{file_id}.json")
        
        if os.path.exists(file_path):
            os.remove(file_path)
            return jsonify({'success': True, 'message': 'File deleted successfully'})
        else:
            return jsonify({'error': 'File not found'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/files/<file_id>/export', methods=['GET'])
def export_file(file_id):
    """Export event file as Excel."""
    try:
        file_path = os.path.join('event_data', f"{file_id}.json")
        
        if not os.path.exists(file_path):
            return jsonify({'error': 'File not found'}), 404
        
        with open(file_path, 'r') as f:
            file_data = json.load(f)
        
        # Create DataFrame for export
        events_data = []
        for event in file_data['events']:
            events_data.append({
                'Date': event['date'],
                'Event Name': event['name'],
                'Event Income': event['income'],
                'All Incurred Expenses': event['expenses'],
                'Underwritten': event['underwritten'],
                'Profit/Loss': event['profit_loss']
            })
        
        df = pd.DataFrame(events_data)
        
        # Create Excel file
        excel_path = os.path.join('event_data', f"{file_id}_export.xlsx")
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Event Financial Data', index=False)
            
            # Get the workbook and worksheet
            workbook = writer.book
            worksheet = writer.sheets['Event Financial Data']
            
            # Format headers
            for col in range(1, len(df.columns) + 1):
                cell = worksheet.cell(row=1, column=col)
                cell.font = cell.font.copy(bold=True)
                cell.fill = openpyxl.styles.PatternFill(start_color='D00000', end_color='D00000', fill_type='solid')
                cell.font = cell.font.copy(color='FFFFFF')
        
        return send_file(excel_path, as_attachment=True, download_name=f"{file_data['name']}.xlsx")
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/files/<file_id>/visualize', methods=['GET'])
def visualize_file(file_id):
    """Get file data for visualization."""
    try:
        file_path = os.path.join('event_data', f"{file_id}.json")
        
        if not os.path.exists(file_path):
            return jsonify({'error': 'File not found'}), 404
        
        with open(file_path, 'r') as f:
            file_data = json.load(f)
        
        # Convert to DataFrame for analysis
        events_data = []
        for event in file_data['events']:
            events_data.append({
                'Date': event['date'],
                'Event Name': event['name'],
                'Event Income': event['income'],
                'All Incurred Expenses': event['expenses'],
                'Underwritten': event['underwritten'],
                'Profit/Loss': event['profit_loss']
            })
        
        df = pd.DataFrame(events_data)
        
        # Analyze the data
        analysis = analyze_data(df)
        
        return jsonify({
            'success': True,
            'file_name': file_data['name'],
            'analysis': {
                'shape': analysis['shape'],
                'data': df.to_dict('records')
            },
            'charts': {},
            'data': df.to_dict('records')
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/visualize-managed-data/<file_id>')
@require_auth
def visualize_managed_data_page(file_id):
    """Page to visualize managed data."""
    try:
        file_path = os.path.join('event_data', f"{file_id}.json")
        
        if not os.path.exists(file_path):
            return "File not found", 404
        
        with open(file_path, 'r') as f:
            file_data = json.load(f)
        
        return render_template('visualize_managed_data.html', 
                             file_id=file_id, 
                             file_name=file_data['name'])
    except Exception as e:
        return f"Error: {str(e)}", 500

@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload and create visualizations."""
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        if file and allowed_file(file.filename):
            # Generate unique filename
            filename = secure_filename(file.filename)
            unique_filename = f"{uuid.uuid4()}_{filename}"
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
            
            # Save file
            file.save(filepath)
            
            # Read file based on extension
            file_extension = filename.rsplit('.', 1)[1].lower()
            if file_extension in ['xlsx', 'xls']:
                df = pd.read_excel(filepath)
            elif file_extension == 'csv':
                # Handle CSV files (standard format)
                df = pd.read_csv(filepath)
                
                # Clean up the data
                df = df.replace(['n/a', 'n/a (AG)', '(n/a)', ''], np.nan)
                
                # Convert currency strings to numeric values
                for col in df.columns:
                    try:
                        if df[col].dtype == 'object':
                            # Convert to string first, then clean and convert to numeric
                            df[col] = df[col].astype(str).str.replace('$', '').str.replace(',', '')
                            # Try to convert to numeric, keeping NaN for non-numeric values
                            df[col] = pd.to_numeric(df[col], errors='coerce')
                    except Exception as e:
                        print(f"Warning: Could not convert column {col}: {e}")
                        continue
            
            # Analyze data
            analysis = analyze_data(df)
            
            # Debug: Print data info
            print(f"DataFrame shape: {df.shape}")
            print(f"DataFrame columns: {list(df.columns)}")
            print(f"DataFrame head:\n{df.head()}")
            
            # Create visualizations
            charts = create_visualizations(df, analysis)
            
            # Clean up uploaded file
            os.remove(filepath)
            
            return jsonify({
                'success': True,
                'analysis': analysis,
                'charts': charts,
                'message': 'File processed successfully!'
            })
        
        else:
            return jsonify({'error': 'Invalid file type. Please upload a data file (.xlsx, .xls, or .csv)'}), 400
    
    except Exception as e:
        return jsonify({'error': f'Error processing file: {str(e)}'}), 500

@app.route('/sample-data')
def get_sample_data():
    """Generate sample data for demonstration."""
    try:
        # Create sample financial data for Ohio Wesleyan University events
        events = [
            "07/27/2024 OWU Football Golf Outing",
            "8/02/2024 OWU Clippers Home Game", 
            "8/15/2024 OWU Legacy Reception",
            "8/21/2024 OWU Senior Class Welcome Back",
            "8/25/2024 OWU Monnett Club Event",
            "9/11/2024 OWU Near You - Toledo",
            "9/11/2024 OWU Near You - Denver",
            "9/12/2024 OWU Near You - Tucson",
            "9/12/2024 OWU Near You - Cleveland",
            "9/12/2024 OWU Near You - Atlanta",
            "9/12/2024 OWU Near You - Chicago",
            "9/13/2024 OWU Near You - Cincinnati"
        ]
        
        data = {
            'Event': events,
            'Event Income': [0, 350, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
            'All Incurred Expenses': [0, 500, 501, 275, 0, 114, 14, 14, 14, 14, 14, 14],
            'Underwritten': [0, 0, 0, 0, 0, 0, 700, 0, 0, 0, 0, 0],
            'Profit/Loss': [0, -150, 0, 0, 0, 0, -14, -14, -14, -14, -14, -14]
        }
        
        df = pd.DataFrame(data)
        
        # Analyze sample data
        analysis = analyze_data(df)
        
        # Create visualizations
        charts = create_visualizations(df, analysis)
        
        return jsonify({
            'success': True,
            'analysis': analysis,
            'charts': charts,
            'message': 'Sample data loaded successfully!'
        })
    
    except Exception as e:
        return jsonify({'error': f'Error generating sample data: {str(e)}'}), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5001)
