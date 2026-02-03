#!/usr/bin/env python3
"""
Main entry point for the Guest Service Recovery Report Generator.
Run with: python3 run.py
"""

import sys
import os

# Add the project root to the Python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app.app import app

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    debug_mode = os.environ.get('FLASK_ENV') == 'development'
    print("Starting Guest Service Recovery Report Generator...")
    print(f"Visit http://localhost:{port} in your browser")
    app.run(debug=debug_mode, host='0.0.0.0', port=port)
