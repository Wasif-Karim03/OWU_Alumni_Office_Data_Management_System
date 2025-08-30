#!/usr/bin/env python3
"""
Startup script for the Excel Data Visualizer Web Application
"""

import os
import sys
import webbrowser
import time
from pathlib import Path

def check_dependencies():
    """Check if all required packages are installed."""
    required_packages = ['flask', 'pandas', 'plotly', 'numpy']
    missing_packages = []
    
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            missing_packages.append(package)
    
    if missing_packages:
        print(f"❌ Missing packages: {', '.join(missing_packages)}")
        print("Please install them using: pip install -r web_requirements.txt")
        return False
    
    print("✅ All required packages are installed")
    return True

def check_files():
    """Check if all required files exist."""
    required_files = [
        'web_visualizer.py',
        'templates/index.html'
    ]
    
    missing_files = []
    for file in required_files:
        if not Path(file).exists():
            missing_files.append(file)
    
    if missing_files:
        print(f"❌ Missing files: {', '.join(missing_files)}")
        return False
    
    print("✅ All required files are present")
    return True

def start_server():
    """Start the Flask server."""
    try:
        print("🚀 Starting Excel Data Visualizer Web Application...")
        print("📁 Working directory:", os.getcwd())
        
        # Import and run the Flask app
        from web_visualizer import app
        
        print("🌐 Server will be available at: http://localhost:5001")
        print("📱 You can also access it from other devices on your network")
        print("⏹️  Press Ctrl+C to stop the server")
        print("\n" + "="*50)
        
        # Open browser after a short delay
        def open_browser():
            time.sleep(2)
            webbrowser.open('http://localhost:5001')
        
        import threading
        browser_thread = threading.Thread(target=open_browser)
        browser_thread.daemon = True
        browser_thread.start()
        
        # Start the Flask app
        app.run(debug=False, host='0.0.0.0', port=5001)
        
    except Exception as e:
        print(f"❌ Error starting server: {e}")
        print("\nTroubleshooting tips:")
        print("1. Make sure you're in the correct directory")
        print("2. Check that all files are present")
        print("3. Verify all dependencies are installed")
        return False
    
    return True

def main():
    """Main function."""
    print("🎯 Excel Data Visualizer Web Application")
    print("="*50)
    
    # Check dependencies
    if not check_dependencies():
        sys.exit(1)
    
    # Check files
    if not check_files():
        sys.exit(1)
    
    # Start server
    start_server()

if __name__ == "__main__":
    main()
