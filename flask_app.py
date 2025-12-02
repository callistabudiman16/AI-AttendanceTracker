from flask import Flask, render_template, request, redirect, url_for, flash, session, send_file, jsonify
import pandas as pd
import qrcode
from PIL import Image
import io
import os
import re
from datetime import datetime
from difflib import SequenceMatcher
import requests
from werkzeug.utils import secure_filename

# Import all the helper functions from the original app
# We'll need to adapt these for Flask

try:
    import google.generativeai as genai
    GEMINI_AVAILABLE = True
except ImportError:
    GEMINI_AVAILABLE = False

try:
    import msal
    MSAL_AVAILABLE = True
except ImportError:
    MSAL_AVAILABLE = False

app = Flask(__name__)
app.secret_key = os.urandom(24)  # Secret key for sessions
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ROSTER_FILE'] = 'roster_attendance.xlsx'

# Create uploads directory if it doesn't exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs('templates', exist_ok=True)
os.makedirs('static', exist_ok=True)

# Initialize session defaults
def init_session():
    if 'roster_loaded' not in session:
        session['roster_loaded'] = False
    if 'min_confidence' not in session:
        session['min_confidence'] = 0.75
    if 'use_gemini' not in session:
        session['use_gemini'] = False
    if 'gemini_api_key' not in session:
        session['gemini_api_key'] = os.getenv('GEMINI_API_KEY', '')
    if 'late_threshold_minutes' not in session:
        session['late_threshold_minutes'] = 15
    if 'class_start_time' not in session:
        session['class_start_time'] = '09:00'
    if 'onedrive_connected' not in session:
        session['onedrive_connected'] = False
    if 'onedrive_client_id' not in session:
        session['onedrive_client_id'] = os.getenv('ONEDRIVE_CLIENT_ID', '')
    if 'onedrive_access_token' not in session:
        session['onedrive_access_token'] = None
    if 'onedrive_file_id' not in session:
        session['onedrive_file_id'] = None
    if 'onedrive_file_path' not in session:
        session['onedrive_file_path'] = 'roster_attendance.xlsx'

# Load all helper functions from app.py - we'll copy them here
# For now, let's create a minimal working version and import the functions

# Copy the essential functions (simplified for Flask)
def normalize_name_for_roster(name):
    """Convert 'Last, First' or 'First, Last' to 'First Last' for matching"""
    if ',' in name:
        parts = [p.strip() for p in name.split(',')]
        if len(parts) == 2:
            return f"{parts[1]} {parts[0]}"
    return name

def extract_name_components(name):
    """Extract first, middle, last name components from various formats"""
    name = name.strip()
    components = {'first': '', 'middle': '', 'last': '', 'middle_initial': ''}
    
    if ',' in name:
        parts = [p.strip() for p in name.split(',')]
        if len(parts) == 2:
            last_part = parts[0].strip()
            first_middle = parts[1].strip().split()
            components['last'] = last_part
            if len(first_middle) >= 1:
                components['first'] = first_middle[0]
            if len(first_middle) >= 2:
                middle = first_middle[1]
                if len(middle) == 1 or (len(middle) == 2 and middle.endswith('.')):
                    components['middle_initial'] = middle.replace('.', '')
                else:
                    components['middle'] = middle
    else:
        parts = name.split()
        if len(parts) >= 2:
            components['first'] = parts[0]
            components['last'] = parts[-1]
            if len(parts) > 2:
                middle_parts = parts[1:-1]
                for mp in middle_parts:
                    if len(mp) == 1 or (len(mp) == 2 and mp.endswith('.')):
                        components['middle_initial'] = mp.replace('.', '')
                    else:
                        if components['middle']:
                            components['middle'] += ' ' + mp
                        else:
                            components['middle'] = mp
    
    return components

def calculate_similarity(str1, str2):
    """Calculate similarity ratio between two strings"""
    return SequenceMatcher(None, str1.lower().strip(), str2.lower().strip()).ratio()

def load_roster():
    """Load roster from file"""
    try:
        if os.path.exists(app.config['ROSTER_FILE']):
            df = pd.read_excel(app.config['ROSTER_FILE'], engine='openpyxl')
            return df
        return None
    except Exception as e:
        flash(f'Error loading roster: {str(e)}', 'error')
        return None

def save_roster(roster_df):
    """Save roster to file"""
    try:
        roster_df.to_excel(app.config['ROSTER_FILE'], index=False, engine='openpyxl')
        return True
    except Exception as e:
        flash(f'Error saving roster: {str(e)}', 'error')
        return False

@app.route('/')
def index():
    init_session()
    roster_df = load_roster()
    return render_template('index.html', 
                         roster=roster_df,
                         roster_loaded=roster_df is not None)

@app.route('/upload_roster', methods=['POST'])
def upload_roster():
    if 'roster_file' not in request.files:
        flash('No file selected', 'error')
        return redirect(url_for('index'))
    
    file = request.files['roster_file']
    if file.filename == '':
        flash('No file selected', 'error')
        return redirect(url_for('index'))
    
    try:
        if file.filename.endswith('.csv'):
            roster_df = pd.read_csv(file)
        else:
            roster_df = pd.read_excel(file)
        
        if save_roster(roster_df):
            session['roster_loaded'] = True
            flash(f'Roster loaded successfully: {len(roster_df)} students', 'success')
        else:
            flash('Failed to save roster', 'error')
    except Exception as e:
        flash(f'Error loading roster: {str(e)}', 'error')
    
    return redirect(url_for('index'))

@app.route('/update_settings', methods=['POST'])
def update_settings():
    session['late_threshold_minutes'] = int(request.form.get('late_threshold', 15))
    session['class_start_time'] = request.form.get('class_start_time', '09:00')
    session['min_confidence'] = float(request.form.get('min_confidence', 0.75))
    session['use_gemini'] = 'use_gemini' in request.form
    session['gemini_api_key'] = request.form.get('gemini_api_key', '')
    flash('Settings updated', 'success')
    return redirect(url_for('index'))

@app.route('/checkin')
def checkin():
    init_session()
    roster_df = load_roster()
    if roster_df is None:
        flash('Please upload a roster file first', 'warning')
        return redirect(url_for('index'))
    return render_template('checkin.html', roster=roster_df)

@app.route('/zoom')
def zoom():
    init_session()
    roster_df = load_roster()
    if roster_df is None:
        flash('Please upload a roster file first', 'warning')
        return redirect(url_for('index'))
    return render_template('zoom.html', roster=roster_df)

@app.route('/view_roster')
def view_roster():
    init_session()
    roster_df = load_roster()
    if roster_df is None:
        flash('Please upload a roster file first', 'warning')
        return redirect(url_for('index'))
    
    # Find date columns
    date_columns = [col for col in roster_df.columns 
                   if re.match(r'\d{4}-\d{2}-\d{2}', str(col))]
    
    roster_with_total = roster_df.copy()
    if date_columns:
        roster_with_total['Total Points'] = roster_with_total[date_columns].sum(axis=1)
    
    return render_template('view_roster.html', 
                         roster=roster_df,
                         roster_with_total=roster_with_total,
                         date_columns=date_columns)

@app.route('/download_roster')
def download_roster():
    roster_df = load_roster()
    if roster_df is None:
        flash('No roster available', 'error')
        return redirect(url_for('index'))
    
    output = io.BytesIO()
    roster_df.to_excel(output, index=False, engine='openpyxl')
    output.seek(0)
    
    filename = f"attendance_roster_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(output, 
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    as_attachment=True,
                    download_name=filename)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
