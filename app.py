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
import base64
import time

# Import optional dependencies
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
app.secret_key = os.urandom(24)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ROSTER_FILE'] = 'roster_attendance.xlsx'

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs('templates', exist_ok=True)
os.makedirs('static', exist_ok=True)

# Helper functions (copied from Streamlit version)
def normalize_name_for_roster(name):
    """Convert 'Last, First' or 'First, Last' to 'First Last' for matching"""
    if ',' in name:
        parts = [p.strip() for p in name.split(',')]
        if len(parts) == 2:
            return f"{parts[1]} {parts[0]}"
    return name

def normalize_name_for_zoom(name):
    """Convert 'First Last' to 'Last,First' (roster format) or 'Last, First' if needed"""
    if ',' not in name:
        parts = name.split()
        if len(parts) >= 2:
            # Convert to "Last,First" format (matching roster format)
            last_name = parts[-1]
            first_names = ' '.join(parts[:-1])
            return f"{last_name},{first_names}"  # No space after comma to match roster
    return name

def extract_name_components(name):
    """Extract first, middle, last name components from various formats"""
    # Normalize the name: remove extra whitespace, handle both comma formats
    name = ' '.join(name.strip().split())  # Normalize whitespace
    
    # Handle both "Last,First" (no space) and "Last, First" (with space) formats
    if ',' in name:
        # Split on comma, but handle both formats
        comma_idx = name.find(',')
        last_part = name[:comma_idx].strip()
        first_part = name[comma_idx+1:].strip()
        components = {'first': '', 'middle': '', 'last': last_part, 'middle_initial': ''}
        
        if first_part:
            first_middle = first_part.split()
            if len(first_middle) >= 1:
                components['first'] = first_middle[0].strip()
            if len(first_middle) >= 2:
                middle = first_middle[1].strip()
                if len(middle) == 1 or (len(middle) == 2 and middle.endswith('.')):
                    components['middle_initial'] = middle.replace('.', '').strip()
                else:
                    components['middle'] = middle
                if len(first_middle) > 2:
                    components['middle'] = ' '.join([m.strip() for m in first_middle[1:]])
    else:
        parts = [p.strip() for p in name.split() if p.strip()]
        if len(parts) >= 2:
            components = {'first': parts[0], 'middle': '', 'last': parts[-1], 'middle_initial': ''}
            if len(parts) > 2:
                middle_parts = parts[1:-1]
                for mp in middle_parts:
                    mp = mp.strip()
                    if not mp:
                        continue
                    if len(mp) == 1 or (len(mp) == 2 and mp.endswith('.')):
                        components['middle_initial'] = mp.replace('.', '')
                    else:
                        if components['middle']:
                            components['middle'] += ' ' + mp
                        else:
                            components['middle'] = mp
        else:
            components = {'first': '', 'middle': '', 'last': '', 'middle_initial': ''}
            if len(parts) == 1:
                components['last'] = parts[0]
    
    # Clean up empty strings
    for key in components:
        components[key] = components[key].strip() if components[key] else ''
    
    return components

def get_all_name_variations(name):
    """Get all possible name format variations for flexible matching"""
    variations = [name]
    components = extract_name_components(name)
    first = components['first']
    middle = components['middle']
    middle_initial = components['middle_initial']
    last = components['last']
    
    if not first or not last:
        if ',' in name:
            parts = [p.strip() for p in name.split(',')]
            if len(parts) == 2:
                variations.append(f"{parts[1]} {parts[0]}")
        else:
            parts = name.split()
            if len(parts) >= 2:
                variations.append(f"{parts[-1]}, {' '.join(parts[:-1])}")
    else:
        if middle:
            variations.append(f"{last}, {first} {middle}")  # With space
            variations.append(f"{last},{first} {middle}")  # No space (roster format)
        if middle_initial:
            variations.append(f"{last}, {first} {middle_initial}")  # With space
            variations.append(f"{last},{first} {middle_initial}")  # No space (roster format)
            variations.append(f"{last}, {first} {middle_initial}.")  # With space and period
            variations.append(f"{last},{first} {middle_initial}.")  # No space and period
        variations.append(f"{last}, {first}")  # With space
        variations.append(f"{last},{first}")  # No space (roster format)
        if middle:
            variations.append(f"{first} {middle} {last}")
        if middle_initial:
            variations.append(f"{first} {middle_initial} {last}")
            variations.append(f"{first} {middle_initial}. {last}")
        variations.append(f"{first} {last}")
        if middle and len(middle) > 0:
            mi = middle[0].upper()
            variations.append(f"{first} {mi} {last}")
            variations.append(f"{first} {mi}. {last}")
            variations.append(f"{last}, {first} {mi}")  # With space
            variations.append(f"{last},{first} {mi}")  # No space (roster format)
            variations.append(f"{last}, {first} {mi}.")  # With space and period
            variations.append(f"{last},{first} {mi}.")  # No space and period
    
    seen = set()
    unique_variations = []
    for v in variations:
        v_lower = v.lower().strip()
        if v_lower and v_lower not in seen:
            seen.add(v_lower)
            unique_variations.append(v)
    return unique_variations

def match_name_with_components(check_in_name, roster_name):
    """Advanced matching that handles partial names and middle initials"""
    check_components = extract_name_components(check_in_name)
    roster_components = extract_name_components(roster_name)
    
    # Must have first and last names to match
    if not check_components['first'] or not check_components['last']:
        return 0.0
    if not roster_components['first'] or not roster_components['last']:
        return 0.0
    
    check_first = check_components['first'].lower().strip()
    check_last = check_components['last'].lower().strip()
    roster_first = roster_components['first'].lower().strip()
    roster_last = roster_components['last'].lower().strip()
    
    # Try normal matching: check first = roster first, check last = roster last
    normal_match = (check_first == roster_first and check_last == roster_last)
    
    # Try swapped matching: check first = roster last, check last = roster first
    # This handles cases like "Warren S. Kyle" (First Last) vs "Warren, Kyle Stephen" (Last, First)
    swapped_match = (check_first == roster_last and check_last == roster_first)
    
    # If neither matches, return 0
    if not normal_match and not swapped_match:
        return 0.0
    
    # Use swapped match result if normal match failed
    using_swapped = swapped_match and not normal_match
    
    # If we get here, first and last names match - now check middle names
    check_middle = check_components.get('middle', '') or check_components.get('middle_initial', '')
    roster_middle = roster_components.get('middle', '') or roster_components.get('middle_initial', '')
    
    # If one or both don't have middle names, still a good match (0.95 confidence)
    if not check_middle or not roster_middle:
        return 0.95
    
    # Both have middle names - check for exact match
    check_middle_lower = check_middle.lower().strip()
    roster_middle_lower = roster_middle.lower().strip()
    
    if check_middle_lower == roster_middle_lower:
        return 1.0
    
    # Check if one middle name is a prefix of the other (e.g., "Ann" matches "Ann Montuya")
    if check_middle_lower.startswith(roster_middle_lower) or roster_middle_lower.startswith(check_middle_lower):
        return 0.96
    
    # Check if the first word of either middle name matches the other
    check_middle_first_word = check_middle_lower.split()[0] if check_middle_lower else ''
    roster_middle_first_word = roster_middle_lower.split()[0] if roster_middle_lower else ''
    
    if check_middle_first_word == roster_middle_first_word:
        return 0.94
    
    # Check middle initial match
    check_initial = check_components.get('middle_initial', '')
    roster_initial = roster_components.get('middle_initial', '')
    
    # Extract initial from middle name if not explicitly set
    if not check_initial and check_middle:
        check_initial = check_middle[0].upper() if check_middle else ''
    if not roster_initial and roster_middle:
        roster_initial = roster_middle[0].upper() if roster_middle else ''
    
    if check_initial and roster_initial and check_initial.lower() == roster_initial.lower():
        return 0.92
    
    # First and last match, but middle names don't match at all - still return high confidence
    # because first+last match is very strong
    return 0.90

def calculate_similarity(str1, str2):
    """Calculate similarity ratio between two strings"""
    return SequenceMatcher(None, str1.lower().strip(), str2.lower().strip()).ratio()

def find_student_with_gemini(student_name, roster_df, name_col):
    """Use Gemini API for name matching"""
    if 'gemini_api_key' not in session or not session.get('gemini_api_key'):
        return None
    
    try:
        genai.configure(api_key=session['gemini_api_key'])
        model = genai.GenerativeModel('gemini-pro')
        roster_names = roster_df[name_col].astype(str).tolist()
        
        prompt = f"""You are matching student attendance names to a roster. The roster contains full names in "Last Name, First Name Middle Name" format.

Student name from attendance: "{student_name}"
Roster names (full names): {roster_names}

Match the student name to the best roster entry. Important:
1. The roster has FULL names (e.g., "Budiman, Natasha Callista")
2. Attendance might have PARTIAL names or variations:
   - "Natasha Budiman" (first + last, missing middle)
   - "Budiman, Natasha" (last + first, missing middle)
   - "Natasha C Budiman" (first + middle initial + last)
   - "Budiman, Natasha C" (last + first + middle initial)
   - "Natasha Callista Budiman" (all parts, different format)

Match based on:
- First name and last name must match
- Middle name/initial can be missing in attendance (still a match)
- Middle initial must match if present in both
- Different name formats (comma position, order) are acceptable

Return ONLY the exact roster name from the list that matches, or "NO_MATCH" if no good match exists.
Do not include any explanation, just return the roster name exactly as it appears in the list, or NO_MATCH."""
        
        response = model.generate_content(prompt)
        matched_name = response.text.strip().strip('"\'`').strip()
        
        if matched_name.upper() == "NO_MATCH" or not matched_name:
            return None
        
        matches = roster_df[roster_df[name_col].str.lower().str.strip() == matched_name.lower().strip()]
        if not matches.empty:
            return matches.index[0], 0.95, matched_name
        
        matches = roster_df[roster_df[name_col].str.lower().str.contains(matched_name.lower(), na=False)]
        if not matches.empty:
            return matches.index[0], 0.95, matches.iloc[0][name_col]
        
    except Exception as e:
        pass
    return None

def find_student_in_roster(student_name, roster_df, use_gemini=False, debug=False):
    """Find student in roster by matching names"""
    name_variations = get_all_name_variations(student_name)
    name_variations.extend([normalize_name_for_roster(student_name), normalize_name_for_zoom(student_name)])
    
    seen = set()
    unique_variations = []
    for v in name_variations:
        v_lower = v.lower().strip()
        if v_lower and v_lower not in seen:
            seen.add(v_lower)
            unique_variations.append(v)
    name_variations = unique_variations
    
    name_col = None
    # First, try to find column with 'name' or 'student' in the name
    # But exclude columns like 'Unnamed' which contain 'name' but aren't the name column
    for col in roster_df.columns:
        col_str = str(col).lower().strip()
        # Check for name column - exclude 'unnamed' columns
        if (('name' in col_str or 'student' in col_str) and 
            'unnamed' not in col_str and 
            col_str not in ['id', 'email', 'major', 'level']):
            name_col = col
            break
    
    # If not found, check all columns and score them based on how name-like they are
    if name_col is None:
        best_score = -1
        best_col = None
        
        for col_idx, col in enumerate(roster_df.columns):
            if roster_df[col].dtype != 'object':
                continue  # Skip non-text columns
            
            # Skip date columns (MM.DD format)
            if re.match(r'^\d{1,2}\.\d{1,2}$', str(col)):
                continue
            
            # Get sample values from this column
            sample_values = roster_df[col].dropna().head(20)
            if len(sample_values) == 0:
                continue
            
            # Score how name-like this column is
            score = 0
            name_like_count = 0
            
            for val in sample_values:
                val_str = str(val).strip()
                # Skip if it's clearly not a name
                if not val_str or val_str.lower() in ['nan', 'none', '']:
                    continue
                
                # Check for name-like patterns
                has_comma = ',' in val_str
                has_spaces = ' ' in val_str
                has_letters = any(c.isalpha() for c in val_str)
                is_numeric = val_str.replace('.', '').replace('-', '').isdigit()
                is_date_like = re.match(r'^\d{1,2}\.\d{1,2}$', val_str) or re.match(r'^\d{1,2}/\d{1,2}', val_str)
                has_multiple_words = len(val_str.split()) >= 2
                
                # Score based on name-like characteristics
                if has_letters and not is_numeric and not is_date_like:
                    score += 1
                    if has_comma or has_multiple_words:
                        score += 2  # Names often have commas (Last, First) or multiple words
                        name_like_count += 1
                    elif has_spaces and has_letters:
                        score += 1
                        name_like_count += 1
            
            # Calculate ratio of name-like values
            if len(sample_values) > 0:
                name_ratio = name_like_count / len(sample_values)
                score = score * name_ratio  # Weight score by ratio of name-like values
                
                # Bonus for column C (index 2) - user specified this is the name column
                if col_idx == 2:
                    score += 10
                
                if score > best_score:
                    best_score = score
                    best_col = col
        
        if best_col is not None:
            name_col = best_col
    
    # Last resort: try column C (index 2) if it exists, then first column
    if name_col is None:
        if len(roster_df.columns) > 2:
            name_col = roster_df.columns[2]  # Column C (index 2)
        else:
            name_col = roster_df.columns[0] if len(roster_df.columns) > 0 else None
    
    if name_col is None:
        return None, 0.0, None
    
    best_match = None
    best_confidence = 0.0
    best_matched_name = None
    
    # First, try exact matches on all variations (highest priority)
    for name_var in name_variations:
        # Normalize the variation to match roster format (remove spaces after comma)
        name_var_normalized = name_var.replace(', ', ',').replace(' ,', ',').lower().strip()
        # Check each roster entry for exact match (normalizing spaces after comma)
        for idx, row in roster_df.iterrows():
            roster_name = str(row[name_col]).strip()
            roster_name_normalized = roster_name.replace(', ', ',').replace(' ,', ',').lower().strip()
            if roster_name_normalized == name_var_normalized:
                return idx, 1.0, roster_name
    
    # Second, try component-based matching (handles partial names, middle names, etc.)
    # This is more reliable than fuzzy matching for names with middle name variations
    # Component matching works for ALL students by checking each roster entry systematically
    # Try component matching on both the original name AND normalized variations
    component_match_found = False
    component_high_confidence_match = False  # Track if we found a high-confidence component match (>= 0.90)
    component_best_confidence = 0.0  # Track best component match separately
    
    # Create a list of name variations to try for component matching
    # Include the original name and key normalized variations
    names_to_try_set = set([student_name])
    zoom_normalized = normalize_name_for_zoom(student_name)
    if zoom_normalized and zoom_normalized != student_name:
        names_to_try_set.add(zoom_normalized)  # Convert "First Last" to "Last,First"
    roster_normalized = normalize_name_for_roster(student_name)
    if roster_normalized and roster_normalized != student_name:
        names_to_try_set.add(roster_normalized)  # Convert "Last, First" to "First Last"
    names_to_try = list(names_to_try_set)
    
    for idx, row in roster_df.iterrows():
        roster_name = str(row[name_col]).strip()
        
        # Try component matching with all name variations
        for try_name in names_to_try:
            if not try_name:
                continue
            component_confidence = match_name_with_components(try_name, roster_name)
            
            # Component matching takes priority - track all component matches
            # This ensures component matching works for ALL students, not just specific cases
            if component_confidence > 0.0:  # Any component match is better than no match
                component_match_found = True
                if component_confidence > component_best_confidence:
                    component_best_confidence = component_confidence
                    best_confidence = component_confidence
                    best_match = idx
                    best_matched_name = roster_name
                    # Mark high-confidence matches - these definitely take priority
                    if component_confidence >= 0.90:
                        component_high_confidence_match = True
                    # If we found a high-confidence match, we can break early
                    if component_confidence >= 0.90:
                        break
        
        # If we found a high-confidence match, no need to check more roster entries
        if component_high_confidence_match and component_best_confidence >= 0.90:
            break
    
    # If component matching found a match (even low confidence), prefer it over fuzzy matching
    # Only use fuzzy matching if component matching found nothing (0.0 for all entries)
    if component_match_found and component_best_confidence > 0.0:
        # Component matching found something - don't let fuzzy matching override unless it's significantly better
        # For component matches >= 0.90, never use fuzzy matching
        # For component matches < 0.90, only use fuzzy if it's much better (unlikely but possible)
        pass  # Keep the component match we found
    
    # Third, try fuzzy matching on all variations (only if component matching didn't find a high-confidence match)
    # Skip fuzzy matching if component matching already found a high-confidence match (>= 0.90)
    # This ensures component matching takes priority for ALL students with good matches
    # Only run fuzzy matching if:
    # 1. Component matching found nothing (component_best_confidence == 0.0), OR
    # 2. Component matching found a low-confidence match (< 0.90) and we want to try fuzzy as fallback
    if not component_high_confidence_match:
        # Only use fuzzy matching if it can potentially improve on component matching
        # If component matching found something, only accept fuzzy matches that are significantly better
        fuzzy_threshold = component_best_confidence if component_match_found else 0.0
        
        for name_var in name_variations:
            name_var_normalized = name_var.replace(', ', ',').replace(' ,', ',')
            for idx, row in roster_df.iterrows():
                roster_name = str(row[name_col]).strip()
                roster_name_normalized = roster_name.replace(', ', ',').replace(' ,', ',')
                similarity = calculate_similarity(name_var_normalized, roster_name_normalized)
                # Only update if fuzzy matching gives a significantly better score
                # This prevents fuzzy matching from overriding good component matches
                if similarity > best_confidence and similarity > fuzzy_threshold:
                    # For component matches < 0.90, only use fuzzy if it's >= 0.90 (much better)
                    if not component_match_found or (component_match_found and similarity >= 0.90):
                        best_confidence = similarity
                        best_match = idx
                        best_matched_name = roster_name
    
    # Fourth, try Gemini if enabled and confidence is low
    if use_gemini and GEMINI_AVAILABLE:
        gemini_match = find_student_with_gemini(student_name, roster_df, name_col)
        if gemini_match:
            gemini_idx, gemini_conf, gemini_matched = gemini_match
            if gemini_conf > best_confidence:
                return gemini_match
    
    # Return match if we found one (accept all matches regardless of confidence)
    if best_match is not None:
        return best_match, best_confidence, best_matched_name
    return None, best_confidence, None

def update_roster_with_attendance(roster_df, student_name, points, date_str, use_gemini=False):
    """Update roster with attendance points"""
    idx, confidence, matched_name = find_student_in_roster(student_name, roster_df, use_gemini)
    if idx is None:
        return roster_df, False, confidence, matched_name
    
    if date_str not in roster_df.columns:
        roster_df[date_str] = 0.0
    
    current_points = roster_df.loc[idx, date_str]
    if pd.isna(current_points) or current_points == 0:
        roster_df.loc[idx, date_str] = points
    else:
        roster_df.loc[idx, date_str] = max(current_points, points)
    
    return roster_df, True, confidence, matched_name

def parse_duration(duration_str):
    """Parse duration string to minutes"""
    # Handle NaN, None, or empty values
    if pd.isna(duration_str) or duration_str is None or str(duration_str).strip().lower() in ['nan', '', 'none']:
        return None
    
    # Convert to string and handle numeric values
    duration_str = str(duration_str).strip()
    
    # If it's already a number (like 171.0), return it as minutes
    try:
        numeric_value = float(duration_str)
        # If it's a reasonable duration value (less than 24 hours = 1440 minutes)
        # assume it's already in minutes
        if 0 <= numeric_value <= 1440:
            return numeric_value
    except (ValueError, TypeError):
        pass
    
    # Clean the string for parsing time formats
    duration_str = re.sub(r'[^\d:.]', '', duration_str)
    
    if not duration_str or duration_str.lower() == 'nan':
        return None
    
    # Handle time format with colons (HH:MM:SS or MM:SS)
    if ':' in duration_str:
        parts = duration_str.split(':')
        try:
            if len(parts) == 3:
                hours, minutes, seconds = map(int, parts)
                return hours * 60 + minutes + seconds / 60
            elif len(parts) == 2:
                minutes, seconds = map(int, parts)
                return minutes + seconds / 60
            elif len(parts) == 1:
                return int(parts[0])
        except (ValueError, TypeError):
            return None
    else:
        # Try to parse as number (assume minutes)
        try:
            return float(duration_str)
        except (ValueError, TypeError):
            return None
    
    return None

def load_roster():
    """Load roster from file"""
    try:
        roster_file_path = app.config['ROSTER_FILE']
        if os.path.exists(roster_file_path):
            # Read Excel file - if first column is 'Unnamed: 0', it's likely an index column
            roster_df = pd.read_excel(roster_file_path, engine='openpyxl')
            
            # If first column is 'Unnamed: 0' and looks like an index (consecutive numbers), drop it
            if len(roster_df.columns) > 0 and 'Unnamed: 0' in roster_df.columns:
                first_col = roster_df.columns[0]
                if first_col == 'Unnamed: 0':
                    # Check if it looks like an index (starts from 0 or 1 and increments)
                    try:
                        first_val = roster_df[first_col].iloc[0]
                        second_val = roster_df[first_col].iloc[1] if len(roster_df) > 1 else None
                        if (isinstance(first_val, (int, float)) and 
                            (second_val is None or isinstance(second_val, (int, float))) and
                            (first_val == 0 or first_val == 1)):
                            # Looks like an index column, drop it
                            roster_df = roster_df.drop(columns=[first_col])
                    except:
                        pass  # If we can't check, keep the column
            
            if roster_df is not None and len(roster_df) > 0:
                return roster_df
            else:
                flash(f'Roster file is empty or invalid', 'warning')
                return None
        else:
            # File doesn't exist - this is okay if roster hasn't been uploaded yet
            return None
    except Exception as e:
        flash(f'Error loading roster: {str(e)}', 'error')
        import traceback
        print(f"Error loading roster: {traceback.format_exc()}")
        return None

def save_roster(roster_df):
    """Save roster to file with better error handling for permission issues"""
    import time
    import os
    
    roster_file = app.config['ROSTER_FILE']
    print(f"Saving roster to: {roster_file}")
    print(f"Roster DataFrame shape: {roster_df.shape}")
    print(f"Roster columns: {list(roster_df.columns)}")
    
    # Calculate Total Points column before saving
    # Find all date columns (MM.DD format or Month.Day format)
    date_columns = []
    non_date_columns = ['Unnamed: 0', 'No.', 'ID', 'Name', 'Major', 'Level', 'Total Points']
    
    for col in roster_df.columns:
        col_str = str(col).strip()
        col_lower = col_str.lower()
        
        # Skip known non-date columns
        if col_str in non_date_columns or col_lower in [c.lower() for c in non_date_columns]:
            continue
        
        # Match MM.DD format (e.g., 10.23, 11.4, 1.5)
        if re.match(r'^\d{1,2}\.\d{1,2}$', col_str):
            date_columns.append(col)
        # Match Month.Day format (e.g., Oct.23, Nov.4, R,Oct.23, T,Oct.21)
        elif re.search(r'(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\.\d{1,2}', col_lower):
            date_columns.append(col)
        # Match date-like patterns with prefixes (R,Oct.23, T,Oct.21, etc.)
        elif re.match(r'^[A-Z],(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\.\d{1,2}', col_lower):
            date_columns.append(col)
    
    # Calculate Total Points from date columns
    if date_columns:
        # Only sum numeric date columns (excluding non-date columns)
        numeric_date_cols = [col for col in date_columns if roster_df[col].dtype in ['int64', 'float64']]
        if numeric_date_cols:
            # Calculate total points, replacing NaN with 0 for calculation
            roster_df['Total Points'] = roster_df[numeric_date_cols].fillna(0).sum(axis=1)
            print(f"Calculated Total Points for {len(roster_df)} students from {len(numeric_date_cols)} date columns")
    
    # Check if file exists and is potentially locked
    if os.path.exists(roster_file):
        # Try to check if file is locked (Windows)
        if os.name == 'nt':
            try:
                # Try to open the file in append mode to check if it's locked
                test_file = open(roster_file, 'r+b')
                test_file.close()
            except (PermissionError, IOError) as lock_error:
                error_msg = (
                    f'Cannot save roster: The file "{os.path.basename(roster_file)}" is currently open in another program.\n\n'
                    f'Please close the file in Excel or any other program and try again.\n\n'
                    f'File location: {roster_file}'
                )
                print(f"ERROR: File is locked - {error_msg}")
                flash(error_msg, 'error')
                return False
    
    # Try saving with retry logic
    max_retries = 3
    retry_delay = 0.5
    
    for attempt in range(max_retries):
        try:
            # Save to Excel
            roster_df.to_excel(roster_file, index=False, engine='openpyxl')
            
            # Verify the file was created/updated
            if os.path.exists(roster_file):
                file_size = os.path.getsize(roster_file)
                print(f"Roster file saved successfully. File size: {file_size} bytes")
                return True
            else:
                print(f"ERROR: Roster file was not created at {roster_file}")
                if attempt < max_retries - 1:
                    time.sleep(retry_delay)
                    continue
                return False
                
        except PermissionError as pe:
            error_msg = (
                f'Permission denied: Cannot save roster file "{os.path.basename(roster_file)}".\n\n'
                f'Possible reasons:\n'
                f'1. The file is open in Microsoft Excel or another program - Please close it and try again\n'
                f'2. You do not have write permissions to this file\n'
                f'3. Another process is using this file\n\n'
                f'File location: {roster_file}\n\n'
                f'Please close the file and try again.'
            )
            print(f"ERROR: Permission denied (attempt {attempt + 1}/{max_retries}) - {error_msg}")
            if attempt < max_retries - 1:
                time.sleep(retry_delay)
                continue
            flash(error_msg, 'error')
            return False
            
        except Exception as e:
            error_msg = f'Error saving roster: {str(e)}'
            print(f"ERROR in save_roster (attempt {attempt + 1}/{max_retries}): {error_msg}")
            
            # Check for permission-related errors
            if 'Permission denied' in str(e) or '[Errno 13]' in str(e):
                detailed_msg = (
                    f'Cannot save roster: Permission denied.\n\n'
                    f'The file "{os.path.basename(roster_file)}" may be:\n'
                    f'- Open in Microsoft Excel (most common cause)\n'
                    f'- Locked by another program\n'
                    f'- Protected by file permissions\n\n'
                    f'File location: {roster_file}\n\n'
                    f'Action: Please close the file in Excel or any other program and try again.'
                )
                if attempt < max_retries - 1:
                    print(f"Retrying in {retry_delay} seconds...")
                    time.sleep(retry_delay)
                    continue
                flash(detailed_msg, 'error')
                return False
            else:
                if attempt < max_retries - 1:
                    time.sleep(retry_delay)
                    continue
                import traceback
                print(traceback.format_exc())
                flash(error_msg, 'error')
                return False
    
    # If all retries failed
    error_msg = f'Failed to save roster after {max_retries} attempts. Please check file permissions and ensure the file is not open in another program.'
    flash(error_msg, 'error')
    return False

def format_date_for_roster(date_input):
    """Convert date to MM.DD format for roster columns"""
    if isinstance(date_input, str):
        # Try parsing various date formats
        try:
            # Try YYYY-MM-DD format first
            dt = datetime.strptime(date_input, "%Y-%m-%d")
        except ValueError:
            try:
                # Try MM/DD/YYYY format
                dt = datetime.strptime(date_input, "%m/%d/%Y")
            except ValueError:
                try:
                    # Try MM-DD-YYYY format
                    dt = datetime.strptime(date_input, "%m-%d-%Y")
                except ValueError:
                    # If already in MM.DD format, return as is
                    if re.match(r'^\d{1,2}\.\d{1,2}$', date_input):
                        return date_input
                    # Default to today
                    dt = datetime.now()
    elif isinstance(date_input, datetime):
        dt = date_input
    else:
        dt = datetime.now()
    
    # Format as MM.DD (e.g., "10.23", "11.4")
    return f"{dt.month}.{dt.day}"

def extract_clean_dsl_code(gemini_response_text):
    """
    Extract clean DSL code from Gemini response, removing explanations and markdown.
    
    Args:
        gemini_response_text: Raw text response from Gemini
        
    Returns:
        Clean DSL code string with only commands
    """
    if not gemini_response_text:
        return ""
    
    text = gemini_response_text.strip()
    
    # Remove markdown code blocks
    if text.startswith('```'):
        lines = text.split('\n')
        # Find the closing ```
        code_start = 0
        code_end = len(lines)
        for i, line in enumerate(lines):
            if line.strip().startswith('```') and i > 0:
                code_start = 1  # Skip first line with ```
                code_end = i
                break
        
        text = '\n'.join(lines[code_start:code_end])
    
    # Split into lines and filter
    lines = text.split('\n')
    dsl_lines = []
    in_code_section = False
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Skip explanation lines (common patterns)
        if any(skip in line.lower() for skip in [
            'here is', 'here\'s', 'here are',
            'the dsl code', 'dsl code is',
            'you can use', 'to do this',
            'generate', 'following', 'below',
            'note:', 'important:', 'remember:',
            'this will', 'this code',
            'i\'ll', 'i will', 'i can',
            'the code', 'code to'
        ]) and not line.startswith(('LOAD', 'PROCESS', 'SHOW', 'VIEW', 'ECHO', 'DELETE', 'SAVE', 'CALCULATE', '#')):
            continue
        
        # Skip lines that look like explanations (long sentences without DSL keywords)
        if len(line) > 100 and not any(keyword in line.upper() for keyword in [
            'LOAD', 'PROCESS', 'SHOW', 'VIEW', 'ECHO', 'DELETE', 'SAVE', 'CALCULATE'
        ]):
            continue
        
        # Keep DSL commands and comments
        if (line.startswith(('LOAD', 'PROCESS', 'SHOW', 'VIEW', 'ECHO', 'DELETE', 'SAVE', 'CALCULATE', '#')) or
            any(keyword in line.upper() for keyword in ['LOAD', 'PROCESS', 'SHOW', 'VIEW', 'ECHO', 'DELETE', 'SAVE', 'CALCULATE'])):
            dsl_lines.append(line)
            in_code_section = True
        elif in_code_section and (line.startswith('#') or len(line) < 80):
            # Keep short lines and comments in code section
            dsl_lines.append(line)
    
    # If we found DSL commands, return them; otherwise return empty
    if dsl_lines:
        return '\n'.join(dsl_lines)
    
    # Fallback: return original if no filtering worked
    return text

def find_matching_date_column(roster_df, date_input):
    """
    Find existing date column in roster that matches the given date.
    Looks for columns like "R,Oct.23", "T,Oct.23", "Oct.23", "10.23", etc.
    Returns the column name if found, None otherwise.
    """
    # Convert date_input to datetime if it's a string
    if isinstance(date_input, str):
        # Try parsing various formats
        try:
            if re.match(r'^\d{1,2}\.\d{1,2}$', date_input):
                # Format is MM.DD
                month, day = map(int, date_input.split('.'))
                dt = datetime(2024, month, day)  # Use 2024 as default year
            else:
                # Try other date formats
                try:
                    dt = datetime.strptime(date_input, "%Y-%m-%d")
                except ValueError:
                    try:
                        dt = datetime.strptime(date_input, "%m/%d/%Y")
                    except ValueError:
                        dt = datetime.strptime(date_input, "%m-%d-%Y")
        except:
            return None
    elif isinstance(date_input, datetime):
        dt = date_input
    else:
        return None
    
    # Get month name abbreviations and numeric formats
    month_names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                   'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    month_name = month_names[dt.month - 1]
    month_num = dt.month
    day = dt.day
    
    # Search for matching columns (case-insensitive)
    # Key date patterns to match: "Oct.23", "10.23" (with or without day prefix)
    primary_patterns = [
        f"{month_name}.{day}",      # "Oct.23"
        f"{month_num}.{day}",       # "10.23"
        f"{month_name}.{day:02d}",  # "Oct.23" (zero-padded)
        f"{month_num}.{day:02d}",   # "10.23" (zero-padded)
    ]
    
    for col in roster_df.columns:
        col_str = str(col).strip()
        col_lower = col_str.lower()
        
        # Check each primary pattern
        for pattern in primary_patterns:
            pattern_lower = pattern.lower()
            
            # Exact match
            if col_lower == pattern_lower:
                return col
            
            # Match with day prefix (e.g., "R,Oct.23", "T,Oct.23")
            # Pattern should appear at the end of the column name (after prefix and comma)
            if col_lower.endswith(pattern_lower):
                # Check if it's after a comma (like "R,Oct.23")
                if ',' in col_lower and col_lower.split(',')[-1].strip() == pattern_lower:
                    return col
                # Or if the pattern is the entire column (minus potential prefix)
                elif len(col_lower) == len(pattern_lower) + 3 and col_lower[-len(pattern_lower):] == pattern_lower:
                    # Likely a prefix like "R," or "T,"
                    return col
    
    # No matching column found
    return None

def init_session():
    """Initialize session defaults"""
    defaults = {
        'roster_loaded': False,
        'use_gemini': False,
        'gemini_api_key': os.getenv('AIzaSyAcR924DTqb4X30QpoM98eqJ3q5IQCXtEQ', ''),
        'late_threshold_minutes': 15,
        'class_start_time': '09:00',
        'early_bird_start_time': '11:00',  # Start time for early bird (0.6 points)
        'regular_start_time': '11:36',      # Start time for regular (0.2 points) - after this time
        'onedrive_connected': False,
        'onedrive_client_id': os.getenv('ONEDRIVE_CLIENT_ID', ''),
        'onedrive_access_token': None,
        'onedrive_file_id': None,
        'onedrive_file_path': 'roster_attendance.xlsx'
    }
    for key, value in defaults.items():
        if key not in session:
            session[key] = value

@app.route('/')
def index():
    init_session()
    roster_df = load_roster()
    return render_template('index.html', roster=roster_df, roster_loaded=roster_df is not None)

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
        # Reset file pointer to beginning
        file.seek(0)
        
        if file.filename.lower().endswith('.csv'):
            # Read CSV file - try UTF-8 first, fallback to latin-1
            try:
                roster_df = pd.read_csv(file, encoding='utf-8')
            except (UnicodeDecodeError, UnicodeError):
                # Fallback to latin-1 encoding
                file.seek(0)
                roster_df = pd.read_csv(file, encoding='latin-1')
        else:
            # Read Excel file (.xlsx, .xls)
            roster_df = pd.read_excel(file, engine='openpyxl')
        
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
    return render_template('checkin.html', 
                         roster=roster_df,
                         early_bird_start_time=session.get('early_bird_start_time', '11:00'),
                         regular_start_time=session.get('regular_start_time', '11:36'))

@app.route('/update_checkin_settings', methods=['POST'])
def update_checkin_settings():
    """Update check-in time settings"""
    init_session()
    try:
        early_bird_start = request.form.get('early_bird_start_time', '11:00')
        regular_start = request.form.get('regular_start_time', '11:36')
        
        # Validate time format
        try:
            datetime.strptime(early_bird_start, '%H:%M')
            datetime.strptime(regular_start, '%H:%M')
        except ValueError:
            flash('Invalid time format. Please use HH:MM format (e.g., 11:00)', 'error')
            return redirect(url_for('checkin'))
        
        session['early_bird_start_time'] = early_bird_start
        session['regular_start_time'] = regular_start
        flash(f'Check-in time settings updated: Early Bird starts at {early_bird_start}, Regular starts at {regular_start}', 'success')
    except Exception as e:
        flash(f'Error updating settings: {str(e)}', 'error')
    
    return redirect(url_for('checkin'))

@app.route('/process_checkin', methods=['POST'])
def process_checkin():
    roster_df = load_roster()
    if roster_df is None:
        flash('Please upload a roster file first', 'warning')
        return redirect(url_for('checkin'))
    
    if 'checkin_file' in request.files:
        file = request.files['checkin_file']
        if file.filename:
            try:
                # Reset file pointer to beginning
                file.seek(0)
                
                if file.filename.lower().endswith('.csv'):
                    # Read CSV file - try UTF-8 first, fallback to latin-1
                    try:
                        checkins_df = pd.read_csv(file, encoding='utf-8')
                    except (UnicodeDecodeError, UnicodeError):
                        # Fallback to latin-1 encoding
                        file.seek(0)
                        checkins_df = pd.read_csv(file, encoding='latin-1')
                else:
                    # Read Excel file (.xlsx, .xls)
                    checkins_df = pd.read_excel(file, engine='openpyxl')
                
                # Find Start Date column first (Qualtrics format) - do this before name detection
                start_date_col = None
                for col in checkins_df.columns:
                    col_lower = col.lower().strip()
                    if 'start date' in col_lower or 'startdate' in col_lower.replace(' ', ''):
                        start_date_col = col
                        break
                
                if start_date_col is None:
                    flash('Warning: "Start Date" column not found. Using current date/time for all check-ins.', 'warning')
                
                # Find name column (case-insensitive search for "Name")
                # Exclude date columns and check if column actually contains name-like data
                name_col = None
                excluded_cols = [start_date_col] if start_date_col else []
                
                # Helper function to check if a column contains name-like data (not dates)
                def is_name_column(col_name, df_col):
                    """Check if a column likely contains names, not dates"""
                    if col_name in excluded_cols:
                        return False
                    sample_values = df_col.dropna().head(10)
                    if len(sample_values) == 0:
                        return False
                    
                    date_count = 0
                    name_count = 0
                    for val in sample_values:
                        val_str = str(val).strip()
                        if not val_str or val_str.lower() in ['nan', 'none', '']:
                            continue
                        # Check if it looks like a date/time
                        if (re.match(r'^\d{4}-\d{2}-\d{2}', val_str) or 
                            re.match(r'^\d{2}/\d{2}/\d{4}', val_str) or
                            re.match(r'^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}', val_str)):
                            date_count += 1
                        # Check if it looks like a name (has letters, not just numbers)
                        elif any(c.isalpha() for c in val_str) and not val_str.replace('.', '').replace('-', '').isdigit():
                            name_count += 1
                    
                    # If more than 50% are dates, it's not a name column
                    total = date_count + name_count
                    if total == 0:
                        return False
                    return (name_count / total) > 0.5
                
                # First, try exact match for "Name"
                for col in checkins_df.columns:
                    if col.lower().strip() == 'name':
                        if is_name_column(col, checkins_df[col]):
                            name_col = col
                            break
                
                # If not found, try columns containing 'name' but not date-related
                if name_col is None:
                    for col in checkins_df.columns:
                        col_lower = col.lower().strip()
                        if 'name' in col_lower and 'date' not in col_lower and 'time' not in col_lower:
                            if is_name_column(col, checkins_df[col]):
                                name_col = col
                                break
                
                # If still not found, try to find a column that looks like it contains names
                # (has letters, not dates, not the start date column)
                if name_col is None:
                    best_name_col = None
                    best_name_ratio = 0
                    for col in checkins_df.columns:
                        if col in excluded_cols:
                            continue
                        sample_values = checkins_df[col].dropna().head(10)
                        if len(sample_values) == 0:
                            continue
                        
                        name_like_count = 0
                        total_valid = 0
                        for val in sample_values:
                            val_str = str(val).strip()
                            if not val_str or val_str.lower() in ['nan', 'none', '']:
                                continue
                            total_valid += 1
                            # Check if it looks like a name (has letters, contains spaces or commas, not a date)
                            if (any(c.isalpha() for c in val_str) and 
                                (' ' in val_str or ',' in val_str) and
                                not re.match(r'^\d{4}-\d{2}-\d{2}', val_str) and
                                not re.match(r'^\d{2}/\d{2}/\d{4}', val_str)):
                                name_like_count += 1
                        
                        if total_valid > 0:
                            name_ratio = name_like_count / total_valid
                            if name_ratio > best_name_ratio and name_ratio > 0.5:
                                best_name_ratio = name_ratio
                                best_name_col = col
                    
                    if best_name_col:
                        name_col = best_name_col
                
                if name_col is None:
                    flash(f'Name column not found. Available columns: {list(checkins_df.columns)}', 'error')
                    return redirect(url_for('checkin'))
                
                # Debug: Show detected columns and sample values
                sample_names = checkins_df[name_col].dropna().head(3).tolist()
                flash(f'Detected columns - Name: {name_col}, Start Date: {start_date_col if start_date_col else "Not found"}', 'info')
                flash(f'Sample names from Name column: {sample_names}', 'info')
                
                # Get configurable time thresholds
                early_bird_start_str = session.get('early_bird_start_time', '11:00')
                regular_start_str = session.get('regular_start_time', '11:36')
                
                try:
                    early_bird_start = datetime.strptime(early_bird_start_str, '%H:%M').time()
                    regular_start = datetime.strptime(regular_start_str, '%H:%M').time()
                except:
                    flash('Error parsing time settings. Using defaults (11:00 and 11:36).', 'warning')
                    early_bird_start = datetime.strptime('11:00', '%H:%M').time()
                    regular_start = datetime.strptime('11:36', '%H:%M').time()
                
                processed_count = 0
                errors = []
                unmatched_students = []  # Track unmatched students with suggestions
                date_points_map = {}  # Track which date gets which points
                
                for idx, row in checkins_df.iterrows():
                    # Get student name
                    student_name = str(row[name_col]).strip()
                    if not student_name or student_name.lower() in ['nan', 'none', '']:
                        continue
                    
                    # Skip if the name looks like a date (wrong column selected)
                    if re.match(r'^\d{4}-\d{2}-\d{2}', student_name) or re.match(r'^\d{2}/\d{2}/\d{4}', student_name):
                        # This is likely a date, not a name - skip this row
                        continue
                    
                    # Get check-in date and time from Start Date column
                    check_in_datetime = None
                    meeting_date = None
                    
                    if start_date_col and pd.notna(row.get(start_date_col)):
                        try:
                            # Try to parse the Start Date column (could be datetime string)
                            check_in_datetime = pd.to_datetime(row[start_date_col])
                            meeting_date = check_in_datetime.date()
                        except:
                            try:
                                # Try parsing as string
                                check_in_datetime = datetime.strptime(str(row[start_date_col]), '%Y-%m-%d %H:%M:%S')
                                meeting_date = check_in_datetime.date()
                            except:
                                try:
                                    check_in_datetime = datetime.strptime(str(row[start_date_col]), '%m/%d/%Y %H:%M:%S')
                                    meeting_date = check_in_datetime.date()
                                except:
                                    pass
                    
                    # If no date found, use today's date and current time (or form date)
                    if meeting_date is None:
                        if 'meeting_date' in request.form and request.form['meeting_date']:
                            try:
                                meeting_date = datetime.strptime(request.form['meeting_date'], '%Y-%m-%d').date()
                                check_in_datetime = datetime.combine(meeting_date, datetime.now().time())
                            except:
                                meeting_date = datetime.now().date()
                                check_in_datetime = datetime.now()
                        else:
                            meeting_date = datetime.now().date()
                            check_in_datetime = datetime.now()
                    
                    # Get check-in time
                    if check_in_datetime:
                        check_in_time = check_in_datetime.time()
                    else:
                        check_in_time = datetime.now().time()
                    
                    # Calculate points based on check-in time
                    # Early bird: from early_bird_start to before regular_start -> 0.6 points
                    # Regular: from regular_start onwards -> 0.2 points
                    # Before early_bird_start: also 0.6 points (students checking in early)
                    
                    check_in_dt = datetime.combine(meeting_date, check_in_time)
                    early_bird_dt = datetime.combine(meeting_date, early_bird_start)
                    regular_dt = datetime.combine(meeting_date, regular_start)
                    
                    if check_in_dt >= regular_dt:
                        points = 0.2  # Regular check-in (11:36 or later)
                    else:
                        # Early bird check-in (before 11:36, including before 11:00)
                        points = 0.6
                    
                    # Format date and try to find matching existing column
                    meeting_datetime = datetime.combine(meeting_date, datetime.min.time())
                    date_str = format_date_for_roster(meeting_datetime)
                    matching_date_col = find_matching_date_column(roster_df, meeting_datetime)
                    if matching_date_col:
                        date_str = matching_date_col
                    
                    # Track date for summary
                    if meeting_date not in date_points_map:
                        date_points_map[meeting_date] = {'early_bird': 0, 'regular': 0, 'total': 0}
                    date_points_map[meeting_date]['total'] += 1
                    if points == 0.6:
                        date_points_map[meeting_date]['early_bird'] += 1
                    elif points == 0.2:
                        date_points_map[meeting_date]['regular'] += 1
                    
                    # Update roster with attendance
                    use_gemini_flag = session.get('use_gemini', False)
                    roster_df, found, confidence, matched_name = update_roster_with_attendance(
                        roster_df, student_name, points, date_str, use_gemini_flag
                    )
                    
                    if found:
                        processed_count += 1
                    else:
                        # Get suggestions by checking all roster entries with component matching
                        # Get name column from roster
                        name_col_roster = None
                        for col in roster_df.columns:
                            col_str = str(col).lower().strip()
                            if ('name' in col_str and 'unnamed' not in col_str and 
                                col_str not in ['id', 'email', 'major', 'level']):
                                name_col_roster = col
                                break
                        if name_col_roster is None:
                            if len(roster_df.columns) > 2:
                                name_col_roster = roster_df.columns[2]
                            else:
                                name_col_roster = roster_df.columns[0] if len(roster_df.columns) > 0 else None
                        
                        # Collect suggestions from all roster entries
                        all_suggestions = []
                        if name_col_roster:
                            for roster_idx, roster_row in roster_df.iterrows():
                                roster_name = str(roster_row[name_col_roster]).strip()
                                if not roster_name or roster_name.lower() in ['nan', 'none', '']:
                                    continue
                                
                                # Try component matching for suggestions
                                comp_conf = match_name_with_components(student_name, roster_name)
                                
                                # Also try with swapped first/last (in case names are in wrong order)
                                # Extract components and try matching with swapped order
                                checkin_components = extract_name_components(student_name)
                                roster_components = extract_name_components(roster_name)
                                
                                # If check-in has First Last format and roster has Last, First format,
                                # try matching check-in's first with roster's first, and check-in's last with roster's last
                                # But also try the reverse (check-in first with roster last, check-in last with roster first)
                                swapped_conf = 0.0
                                if (checkin_components.get('first') and checkin_components.get('last') and
                                    roster_components.get('first') and roster_components.get('last')):
                                    # Try matching: checkin first = roster last, checkin last = roster first
                                    if (checkin_components['first'].lower() == roster_components['last'].lower() and
                                        checkin_components['last'].lower() == roster_components['first'].lower()):
                                        swapped_conf = 0.85  # High confidence for swapped match
                                
                                # Use the higher confidence
                                best_conf = max(comp_conf, swapped_conf)
                                if best_conf > 0.30:  # Low threshold for suggestions (30% confidence)
                                    all_suggestions.append({
                                        'name': roster_name,
                                        'confidence': best_conf,
                                        'index': roster_idx
                                    })
                                
                                # Also try fuzzy matching for additional suggestions
                                if comp_conf < 0.90:  # Only if component matching didn't find a high match
                                    name_variations = get_all_name_variations(student_name)
                                    for name_var in name_variations[:3]:  # Try first 3 variations
                                        from difflib import SequenceMatcher
                                        similarity = SequenceMatcher(
                                            None, 
                                            name_var.lower().strip().replace(', ', ',').replace(' ,', ','),
                                            roster_name.lower().strip().replace(', ', ',').replace(' ,', ',')
                                        ).ratio()
                                        if similarity > 0.50 and similarity not in [s['confidence'] for s in all_suggestions if s['name'] == roster_name]:
                                            all_suggestions.append({
                                                'name': roster_name,
                                                'confidence': similarity,
                                                'index': roster_idx
                                            })
                                            break  # Only add once per roster entry
                        
                        # Remove duplicates (keep highest confidence for each name)
                        seen_names = {}
                        for sugg in all_suggestions:
                            if sugg['name'] not in seen_names or sugg['confidence'] > seen_names[sugg['name']]['confidence']:
                                seen_names[sugg['name']] = sugg
                        
                        # Sort by confidence and get top 3
                        suggestions = sorted(seen_names.values(), key=lambda x: x['confidence'], reverse=True)[:3]
                        
                        unmatched_students.append({
                            'checkin_name': student_name,
                            'checkin_time': check_in_time.strftime('%H:%M'),
                            'meeting_date': meeting_date.isoformat(),
                            'date_str': date_str,
                            'points': points,
                            'suggestions': suggestions,
                            'best_confidence': confidence,
                            'best_match': matched_name if matched_name else None
                        })
                        errors.append(f"{student_name} (checked in at {check_in_time.strftime('%H:%M')})")
                
                # Save roster with processed check-ins
                if save_roster(roster_df):
                    success_msg = f'Processed {processed_count} check-ins successfully'
                    if date_points_map:
                        dates_info = []
                        for date, info in date_points_map.items():
                            dates_info.append(f"{date.strftime('%m/%d/%Y')}: {info['early_bird']} early bird (0.6pts), {info['regular']} regular (0.2pts)")
                        if dates_info:
                            success_msg += f" - {'; '.join(dates_info)}"
                    flash(success_msg, 'success')
                    
                    # If there are unmatched students, store them for review
                    if unmatched_students:
                        session['unmatched_students'] = unmatched_students
                        flash(f'Found {len(unmatched_students)} unmatched students. Please review and confirm matches.', 'warning')
                        return redirect(url_for('review_checkin_matches'))
                    elif errors:
                        if len(errors) <= 5:
                            flash(f'Could not match: {"; ".join(errors)}', 'warning')
                        else:
                            flash(f'Could not match: {"; ".join(errors[:5])}... and {len(errors) - 5} more', 'warning')
                else:
                    flash('Failed to save roster', 'error')
            except Exception as e:
                flash(f'Error processing file: {str(e)}', 'error')
                import traceback
                print(f"Error in process_checkin: {traceback.format_exc()}")
    
    return redirect(url_for('checkin'))

@app.route('/review_checkin_matches')
def review_checkin_matches():
    """Review page for unmatched students with suggestions"""
    init_session()
    unmatched_students = session.get('unmatched_students', [])
    
    if not unmatched_students:
        flash('No unmatched students to review', 'info')
        return redirect(url_for('checkin'))
    
    roster_df = load_roster()
    if roster_df is None:
        flash('Please upload a roster file first', 'warning')
        return redirect(url_for('index'))
    
    return render_template('review_checkin_matches.html', 
                         unmatched_students=unmatched_students,
                         roster=roster_df)

@app.route('/confirm_checkin_matches', methods=['POST'])
def confirm_checkin_matches():
    """Process confirmed matches for unmatched students"""
    init_session()
    roster_df = load_roster()
    if roster_df is None:
        flash('Please upload a roster file first', 'warning')
        return redirect(url_for('index'))
    
    unmatched_students = session.get('unmatched_students', [])
    if not unmatched_students:
        flash('No unmatched students to process', 'info')
        return redirect(url_for('checkin'))
    
    confirmed_count = 0
    skipped_count = 0
    errors_list = []
    
    # Debug: Print form data
    print(f"Form data received: {list(request.form.keys())}")
    print(f"Number of unmatched students: {len(unmatched_students)}")
    
    # Process confirmed matches
    for i, unmatched in enumerate(unmatched_students):
        print(f"Processing unmatched student {i}: {unmatched.get('checkin_name', 'Unknown')}")
        # Check if this student was confirmed
        match_key = f'match_{i}'
        confirmed_roster_name = request.form.get(match_key)
        
        # Check for manual entry
        if confirmed_roster_name == '__manual__':
            manual_key = f'manual_match_{i}'
            confirmed_roster_name = request.form.get(manual_key)
        
        if confirmed_roster_name and confirmed_roster_name != 'skip' and confirmed_roster_name != '__manual__':
            # Find the student in roster by name
            name_col = None
            for col in roster_df.columns:
                col_str = str(col).lower().strip()
                if ('name' in col_str and 'unnamed' not in col_str and 
                    col_str not in ['id', 'email', 'major', 'level']):
                    name_col = col
                    break
            if name_col is None:
                if len(roster_df.columns) > 2:
                    name_col = roster_df.columns[2]
            
            if name_col:
                # Find the matching row (case-insensitive, handle whitespace)
                confirmed_name_clean = confirmed_roster_name.strip()
                matches = roster_df[roster_df[name_col].astype(str).str.strip().str.lower() == confirmed_name_clean.lower()]
                if not matches.empty:
                    idx = matches.index[0]
                    date_str = unmatched['date_str']
                    points = unmatched['points']
                    
                    # Try to find matching date column (in case it exists in different format)
                    meeting_date = datetime.fromisoformat(unmatched['meeting_date'])
                    matching_date_col = find_matching_date_column(roster_df, meeting_date)
                    if matching_date_col:
                        date_str = matching_date_col
                    
                    # Update roster - ensure date column exists
                    if date_str not in roster_df.columns:
                        roster_df[date_str] = 0.0
                    
                    # Get current points (handle NaN and convert to float)
                    current_points = roster_df.loc[idx, date_str]
                    if pd.isna(current_points):
                        current_points = 0.0
                    else:
                        try:
                            current_points = float(current_points)
                        except (ValueError, TypeError):
                            current_points = 0.0
                    
                    # Update points (use max to handle multiple check-ins)
                    new_points = points if current_points == 0 else max(current_points, points)
                    roster_df.loc[idx, date_str] = new_points
                    
                    # Verify the update was applied
                    verify_value = roster_df.loc[idx, date_str]
                    print(f"Updated student {confirmed_roster_name} at index {idx} with {new_points} points for date {date_str}")
                    print(f"Verification - Value in roster after update: {verify_value} (type: {type(verify_value)})")
                    print(f"Student name from roster: {roster_df.loc[idx, name_col]}")
                    
                    if pd.isna(verify_value) or (isinstance(verify_value, (int, float)) and verify_value != new_points):
                        error_msg = f"WARNING: Update may have failed for {confirmed_roster_name}. Expected {new_points}, got {verify_value}"
                        print(error_msg)
                        errors_list.append(error_msg)
                    else:
                        confirmed_count += 1
                else:
                    # Student name not found in roster - this shouldn't happen but handle it
                    error_msg = f'Could not find "{confirmed_roster_name}" in roster. Available names sample: {list(roster_df[name_col].head(5))}'
                    print(error_msg)
                    errors_list.append(error_msg)
                    flash(f'Warning: {error_msg}', 'warning')
        elif confirmed_roster_name == 'skip':
            skipped_count += 1
            print(f"Skipped student {i}: {unmatched.get('checkin_name', 'Unknown')}")
        else:
            print(f"No match selected for student {i}: {unmatched.get('checkin_name', 'Unknown')} (match_key={match_key}, value={confirmed_roster_name})")
    
    # Save updated roster
    print(f"Saving roster with {confirmed_count} confirmed matches and {skipped_count} skipped")
    print(f"Roster shape before save: {roster_df.shape}")
    date_cols = [col for col in roster_df.columns if any(x in str(col).lower() for x in ['nov', 'oct', 'dec', 'jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', '.'])]
    print(f"Date columns in roster: {date_cols}")
    
    if confirmed_count > 0 or skipped_count > 0:
        if save_roster(roster_df):
            success_msg = f'Successfully confirmed {confirmed_count} match{"es" if confirmed_count != 1 else ""}. {skipped_count} student{"s" if skipped_count != 1 else ""} skipped.'
            if errors_list:
                success_msg += f' Warnings: {"; ".join(errors_list[:3])}'
            flash(success_msg, 'success')
            # Clear unmatched students from session
            session.pop('unmatched_students', None)
            # Redirect to view roster so user can see the updates
            return redirect(url_for('view_roster'))
        else:
            flash('Failed to save roster. Please check the file permissions.', 'error')
            print("ERROR: Failed to save roster")
            import traceback
            print(traceback.format_exc())
    else:
        flash('No matches were confirmed. Please select matches or skip students.', 'warning')
    
    return redirect(url_for('checkin'))

@app.route('/zoom')
def zoom():
    init_session()
    roster_df = load_roster()
    if roster_df is None:
        flash('Please upload a roster file first', 'warning')
        return redirect(url_for('index'))
    return render_template('zoom.html', roster=roster_df)

@app.route('/process_zoom', methods=['POST'])
def process_zoom():
    roster_df = load_roster()
    if roster_df is None:
        flash('Please upload a roster file first', 'warning')
        return redirect(url_for('zoom'))
    
    if 'zoom_file' not in request.files:
        flash('No file selected', 'error')
        return redirect(url_for('zoom'))
    
    file = request.files['zoom_file']
    if not file.filename:
        flash('No file selected', 'error')
        return redirect(url_for('zoom'))
    
    try:
        # Reset file pointer to beginning
        file.seek(0)
        
        # Read Zoom file with headers (new format has column names)
        if file.filename.lower().endswith('.csv'):
            # Read CSV file - try UTF-8 first, fallback to latin-1
            try:
                zoom_df = pd.read_csv(file, encoding='utf-8')
            except (UnicodeDecodeError, UnicodeError):
                # Fallback to latin-1 encoding
                file.seek(0)
                zoom_df = pd.read_csv(file, encoding='latin-1')
        else:
            # Read Excel file (.xlsx, .xls) with headers
            zoom_df = pd.read_excel(file, engine='openpyxl')
        
        flash(f"Read {len(zoom_df)} rows from Zoom file", 'info')
        flash(f"Columns found: {list(zoom_df.columns)}", 'info')
        
        # Find columns by header name (like terminal GUI)
        name_col = None
        duration_col = None
        guest_col = None
        
        # Find Name column (could be "Name", "Name (original name)", etc.)
        for col in zoom_df.columns:
            col_str = str(col).lower().strip()
            if 'name' in col_str and 'guest' not in col_str:
                name_col = col
                break
        
        # Find Total duration column (could be "Total duration", "Total duration (minutes)", etc.)
        for col in zoom_df.columns:
            col_str = str(col).lower().strip()
            if 'duration' in col_str or ('time' in col_str and 'guest' not in col_str):
                duration_col = col
                break
        
        # Find Guest column (to ignore it)
        for col in zoom_df.columns:
            col_str = str(col).lower().strip()
            if 'guest' in col_str:
                guest_col = col
                break
        
        if name_col is None:
            flash("Error: Could not find 'Name' column in Zoom file", 'error')
            flash(f"Available columns: {list(zoom_df.columns)}", 'error')
            return redirect(url_for('zoom'))
        
        if duration_col is None:
            flash("Error: Could not find 'Total duration' column in Zoom file", 'error')
            flash(f"Available columns: {list(zoom_df.columns)}", 'error')
            return redirect(url_for('zoom'))
        
        flash(f"Using columns - Name: {name_col}, Duration: {duration_col}", 'info')
        
        # Get cut time from form (default to 30 minutes if not provided)
        try:
            cut_time_minutes = int(request.form.get('cut_time', session.get('cut_time', 30)))
            if cut_time_minutes <= 0:
                flash('Cut time must be greater than 0. Using default: 30 minutes', 'warning')
                cut_time_minutes = 30
            # Save to session for next time
            session['cut_time'] = cut_time_minutes
        except (ValueError, TypeError):
            flash('Invalid cut time. Using default: 30 minutes', 'warning')
            cut_time_minutes = session.get('cut_time', 30)
        
        # Extract meeting date
        meeting_date = None
        
        # First, try to get date from form input
        form_date = request.form.get('meeting_date', '')
        if form_date:
            try:
                meeting_date = datetime.strptime(form_date, "%Y-%m-%d")
            except ValueError:
                pass
        
        # If no form date, use today's date
        if meeting_date is None:
            meeting_date = datetime.now()
        
        # Format date as MM.DD for roster column
        date_str = format_date_for_roster(meeting_date)
        
        # Try to find existing date column that matches this date
        # Look for columns like "R,Oct.23", "T,Oct.23", "Oct.23", "10.23", etc.
        matching_date_col = None
        if roster_df is not None and len(roster_df) > 0:
            matching_date_col = find_matching_date_column(roster_df, meeting_date)
            if matching_date_col:
                date_str = matching_date_col
                flash(f"Found existing date column: {date_str}", 'info')
            else:
                # No matching column found, will create new one in MM.DD format
                flash(f"No existing date column found for {date_str}, will create new column", 'info')
        
        # Remove header row if it's in the data (check if first row looks like a header)
        if len(zoom_df) > 0:
            first_row_name = str(zoom_df.iloc[0].get(name_col, '')).lower().strip()
            if first_row_name in ['name', 'name (original name)', 'participant']:
                zoom_df = zoom_df.iloc[1:].reset_index(drop=True)
                flash("Removed header row from data", 'info')
        
        flash(f"Processing {len(zoom_df)} student records", 'info')
        
        # Debug info: show detected columns and meeting date
        flash(f"Meeting date: {meeting_date.strftime('%Y-%m-%d') if meeting_date else 'Not found'}, formatted as: {date_str}", 'info')
        flash(f"Cut time: {cut_time_minutes} minutes (≥{cut_time_minutes} min = 0.6 pts, <{cut_time_minutes} min = 0.2 pts)", 'info')
        
        # Additional debug: show sample values from detected columns to verify they're correct
        if duration_col and len(zoom_df) > 0:
            sample_durations = zoom_df[duration_col].dropna().head(5).tolist()
            flash(f"Sample duration values: {sample_durations}", 'info')
        if name_col and len(zoom_df) > 0:
            sample_names = zoom_df[name_col].dropna().head(5).tolist()
            flash(f"Sample student names: {sample_names}", 'info')
        
        if date_str not in roster_df.columns:
            roster_df[date_str] = 0.0
        
        processed_count = 0
        errors = []
        skipped_count = 0
        skip_reasons = {}  # Track why rows are skipped
        
        for idx, row in zoom_df.iterrows():
            # Skip rows where name is missing or invalid
            name_val = row.get(name_col)
            if pd.isna(name_val) or str(name_val).strip().lower() in ['nan', '', 'none', 'name', 'participant']:
                skipped_count += 1
                skip_reasons['missing_name'] = skip_reasons.get('missing_name', 0) + 1
                continue
            
            student_name = str(name_val).strip()
            
            # Skip header rows, summary rows, or meeting titles (common patterns)
            # Be more specific - only skip if the entire name matches a header pattern
            name_lower = student_name.lower().strip()
            # Only skip if name is exactly a header word, not if it contains a header word
            header_exact_matches = ['name', 'participant', 'total', 'summary', 'meeting', 'zoom', 'class', 
                                   'attendance', 'report', 'participants']
            # Skip if name is exactly a header or starts with common header patterns
            if name_lower in header_exact_matches or name_lower.startswith('name (') or 'original name' in name_lower:
                skipped_count += 1
                skip_reasons['header_exact'] = skip_reasons.get('header_exact', 0) + 1
                continue
            
            # Skip if the name looks like a meeting title (contains common class identifiers)
            if re.search(r'(mhr|class|zoom|meeting|session)\s*\d+', name_lower, re.IGNORECASE):
                skipped_count += 1
                skip_reasons['meeting_title'] = skip_reasons.get('meeting_title', 0) + 1
                continue
            
            # Skip rows where the name looks like a number (wrong column selected or summary row)
            try:
                float(student_name)
                skipped_count += 1
                skip_reasons['numeric_name'] = skip_reasons.get('numeric_name', 0) + 1
                continue
            except ValueError:
                pass  # Good, it's not a number
            
            # Skip if name is too short (likely not a real name) or too long (likely a title)
            if len(student_name) < 3 or len(student_name) > 100:
                skipped_count += 1
                skip_reasons['name_length'] = skip_reasons.get('name_length', 0) + 1
                continue
            
            # Get duration (should already be in minutes from "Total duration (minutes)" column)
            duration_val = row.get(duration_col) if duration_col else None
            
            # Try to parse as numeric first (since it's "Total duration (minutes)")
            try:
                duration_minutes = float(duration_val)
            except (ValueError, TypeError):
                # If not numeric, try parse_duration function
                duration_minutes = parse_duration(duration_val)
            
            if duration_minutes is None or pd.isna(duration_minutes):
                errors.append(f"Could not parse duration for {student_name}: {duration_val}")
                continue
            
            # Calculate points based on cut time
            # If duration >= cut_time_minutes: 0.6 points
            # If duration < cut_time_minutes but > 0: 0.2 points
            # If duration == 0: 0.0 points
            if duration_minutes >= cut_time_minutes:
                points = 0.6  # Full attendance
            elif duration_minutes > 0:
                points = 0.2  # Partial attendance
            else:
                points = 0.0  # No attendance
            
            # Check if student's name is in the full roster
            use_gemini_flag = session.get('use_gemini', False)
            
            # Normalize the Zoom name to roster format for matching
            # "Andrea Morales" (Zoom) -> "Morales,Andrea" (roster format)
            normalized_student_name = normalize_name_for_zoom(student_name)
            
            roster_df, found, confidence, matched_name = update_roster_with_attendance(
                roster_df, normalized_student_name if normalized_student_name != student_name else student_name, 
                points, date_str, use_gemini_flag
            )
            
            if found:
                processed_count += 1
            else:
                # Student name not found in roster - add to errors for review
                # Include more detail about why it failed
                error_detail = f"{student_name} (confidence: {confidence:.2f}"
                if matched_name:
                    error_detail += f", best match: {matched_name}"
                error_detail += ")"
                errors.append(f"Student not found in roster: {error_detail}")
        
        if save_roster(roster_df):
            success_msg = f'Processed {processed_count} Zoom attendees successfully for date {date_str}'
            if skipped_count > 0:
                skip_details = ', '.join([f'{k}: {v}' for k, v in skip_reasons.items()])
                success_msg += f' (skipped {skipped_count} rows'
                if skip_details:
                    success_msg += f': {skip_details}'
                success_msg += ')'
            flash(success_msg, 'success')
            if errors:
                error_msg = f'Warning: {len(errors)} issues found'
                if len(errors) <= 3:
                    flash(f'{error_msg}: {"; ".join(errors)}', 'warning')
                else:
                    flash(f'{error_msg}: {"; ".join(errors[:3])}... and {len(errors) - 3} more', 'warning')
        else:
            flash('Failed to save roster', 'error')
    except Exception as e:
        flash(f'Error processing Zoom file: {str(e)}', 'error')
    
    return redirect(url_for('zoom'))

@app.route('/view_roster')
def view_roster():
    init_session()
    roster_df = load_roster()
    if roster_df is None:
        flash('Please upload a roster file first', 'warning')
        return redirect(url_for('index'))
    
    # Match date columns - detect various formats:
    # MM.DD format (e.g., 10.23, 11.4)
    # Month.Day format (e.g., Oct.23, Nov.4, R,Oct.23, T,Oct.21)
    date_columns = []
    non_date_columns = ['Unnamed: 0', 'No.', 'ID', 'Name', 'Major', 'Level', 'Total Points']
    
    for col in roster_df.columns:
        col_str = str(col).strip()
        col_lower = col_str.lower()
        
        # Skip known non-date columns
        if col_str in non_date_columns or col_lower in [c.lower() for c in non_date_columns]:
            continue
        
        # Match MM.DD format (e.g., 10.23, 11.4, 1.5)
        if re.match(r'^\d{1,2}\.\d{1,2}$', col_str):
            date_columns.append(col)
        # Match Month.Day format (e.g., Oct.23, Nov.4, R,Oct.23, T,Oct.21)
        elif re.search(r'(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\.\d{1,2}', col_lower):
            date_columns.append(col)
        # Match date-like patterns with prefixes (R,Oct.23, T,Oct.21, etc.)
        elif re.match(r'^[A-Z],(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\.\d{1,2}', col_lower):
            date_columns.append(col)
    
    roster_with_total = roster_df.copy()
    
    # Rename "Unnamed: 0" column to "No." if it exists
    if 'Unnamed: 0' in roster_with_total.columns:
        roster_with_total = roster_with_total.rename(columns={'Unnamed: 0': 'No.'})
    
    # Convert "No." column to integers (if it exists)
    if 'No.' in roster_with_total.columns:
        # Convert to numeric, then to int, handling NaN values
        roster_with_total['No.'] = pd.to_numeric(roster_with_total['No.'], errors='coerce')
        roster_with_total['No.'] = roster_with_total['No.'].apply(
            lambda x: int(x) if pd.notna(x) else ''
        )
    
    # Convert "ID" column to integers (if it exists)
    if 'ID' in roster_with_total.columns:
        # Convert to numeric, then to int, handling NaN values
        roster_with_total['ID'] = pd.to_numeric(roster_with_total['ID'], errors='coerce')
        roster_with_total['ID'] = roster_with_total['ID'].apply(
            lambda x: int(x) if pd.notna(x) else ''
        )
    
    if date_columns:
        # Only sum numeric date columns (excluding non-date columns)
        numeric_date_cols = [col for col in date_columns if roster_with_total[col].dtype in ['int64', 'float64']]
        if numeric_date_cols:
            # Calculate total points, replacing NaN with 0 for calculation
            roster_with_total['Total Points'] = roster_with_total[numeric_date_cols].fillna(0).sum(axis=1)
    
    # Format numeric columns (date columns and Total Points) to show 1 decimal place
    # Exclude "No." and "ID" columns as they should be integers
    # Do this before fillna so we can check dtypes
    numeric_cols = [col for col in roster_with_total.columns 
                   if roster_with_total[col].dtype in ['int64', 'float64'] 
                   and col not in ['No.', 'ID']]
    for col in numeric_cols:
        # Format numeric values, but keep NaN as NaN for now
        roster_with_total[col] = roster_with_total[col].apply(
            lambda x: f"{x:.1f}" if pd.notna(x) else x
        )
    
    # Replace NaN values with empty strings for cleaner display
    # This will handle all NaN values including in date columns and Total Points
    roster_with_total = roster_with_total.fillna('')
    
    # Convert DataFrame to HTML for display with gridlines
    roster_html = roster_with_total.to_html(classes='table table-striped table-hover table-bordered', escape=False, index=False, na_rep='')
    
    return render_template('view_roster.html', 
                         roster=roster_df,
                         roster_html=roster_html,
                         roster_with_total=roster_with_total,
                         date_columns=date_columns)

@app.route('/delete_date_column', methods=['POST'])
def delete_date_column():
    """Delete a date column and its data from the roster"""
    try:
        init_session()
        roster_df = load_roster()
        if roster_df is None:
            flash('Please upload a roster file first', 'warning')
            return redirect(url_for('view_roster'))
        
        date_column = request.form.get('date_column')
        if not date_column:
            flash('No date column specified', 'error')
            return redirect(url_for('view_roster'))
        
        # Debug: Log the received column name
        print(f"Delete request received for column: {date_column}")
        print(f"Available columns: {list(roster_df.columns)}")
    except Exception as e:
        flash(f'Error processing delete request: {str(e)}', 'error')
        import traceback
        print(f"Error in delete_date_column: {traceback.format_exc()}")
        return redirect(url_for('view_roster'))
    
    # Verify the column exists
    if date_column not in roster_df.columns:
        flash(f'Column "{date_column}" does not exist', 'error')
        return redirect(url_for('view_roster'))
    
    # Check if it's actually a date column (safety check)
    col_str = str(date_column).strip()
    col_lower = col_str.lower()
    is_date_column = (
        re.match(r'^\d{1,2}\.\d{1,2}$', col_str) or
        re.search(r'(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\.\d{1,2}', col_lower) or
        re.match(r'^[A-Z],(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\.\d{1,2}', col_lower)
    )
    
    # Prevent deletion of non-date columns (safety check)
    non_date_columns = ['Unnamed: 0', 'No.', 'ID', 'Name', 'Major', 'Level', 'Total Points']
    if col_str in non_date_columns or col_lower in [c.lower() for c in non_date_columns]:
        flash(f'Cannot delete protected column: {date_column}', 'error')
        return redirect(url_for('view_roster'))
    
    if not is_date_column:
        flash(f'Column "{date_column}" does not appear to be a date column. Deletion cancelled for safety.', 'error')
        return redirect(url_for('view_roster'))
    
    try:
        # Delete the column
        roster_df = roster_df.drop(columns=[date_column])
        
        # Save the updated roster
        if save_roster(roster_df):
            flash(f'Successfully deleted date column: {date_column}', 'success')
        else:
            flash('Failed to save roster after deletion', 'error')
    except Exception as e:
        flash(f'Error deleting column: {str(e)}', 'error')
    
    return redirect(url_for('view_roster'))

@app.route('/download_roster')
def download_roster():
    roster_df = load_roster()
    if roster_df is None:
        flash('No roster available', 'error')
        return redirect(url_for('index'))
    
    # Calculate Total Points before downloading
    date_columns = []
    non_date_columns = ['Unnamed: 0', 'No.', 'ID', 'Name', 'Major', 'Level', 'Total Points']
    
    for col in roster_df.columns:
        col_str = str(col).strip()
        col_lower = col_str.lower()
        
        # Skip known non-date columns
        if col_str in non_date_columns or col_lower in [c.lower() for c in non_date_columns]:
            continue
        
        # Match MM.DD format (e.g., 10.23, 11.4, 1.5)
        if re.match(r'^\d{1,2}\.\d{1,2}$', col_str):
            date_columns.append(col)
        # Match Month.Day format (e.g., Oct.23, Nov.4, R,Oct.23, T,Oct.21)
        elif re.search(r'(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\.\d{1,2}', col_lower):
            date_columns.append(col)
        # Match date-like patterns with prefixes (R,Oct.23, T,Oct.21, etc.)
        elif re.match(r'^[A-Z],(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\.\d{1,2}', col_lower):
            date_columns.append(col)
    
    # Calculate Total Points from date columns
    if date_columns:
        numeric_date_cols = [col for col in date_columns if roster_df[col].dtype in ['int64', 'float64']]
        if numeric_date_cols:
            roster_df['Total Points'] = roster_df[numeric_date_cols].fillna(0).sum(axis=1)
    
    output = io.BytesIO()
    roster_df.to_excel(output, index=False, engine='openpyxl')
    output.seek(0)
    
    filename = f"attendance_roster_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(output, 
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    as_attachment=True,
                    download_name=filename)

@app.route('/generate_qr', methods=['POST'])
def generate_qr():
    qr_data = request.form.get('qr_url', '')
    if qr_data:
        qr = qrcode.QRCode(version=1, box_size=10, border=5)
        qr.add_data(qr_data)
        qr.make(fit=True)
        qr_img = qr.make_image(fill_color="black", back_color="white")
        
        img_io = io.BytesIO()
        qr_img.save(img_io, 'PNG')
        img_io.seek(0)
        img_base64 = base64.b64encode(img_io.getvalue()).decode()
        return jsonify({'qr_image': f'data:image/png;base64,{img_base64}'})
    return jsonify({'error': 'No URL provided'}), 400

# DSL Executor Routes
@app.route('/dsl')
def dsl_interface():
    """DSL script execution interface"""
    init_session()
    return render_template('dsl.html')

@app.route('/execute_dsl', methods=['POST'])
def execute_dsl():
    """Execute a DSL script"""
    init_session()
    
    try:
        from dsl.dsl_integrated import IntegratedDSLExecutor
        
        script_content = request.form.get('script_content', '')
        if not script_content:
            flash('No script content provided', 'error')
            return redirect(url_for('dsl_interface'))
        
        # Create executor with app functions
        app_functions = {
            'load_roster': load_roster,
            'save_roster': save_roster,
            'format_date_for_roster': format_date_for_roster,
            'find_matching_date_column': find_matching_date_column,
            'find_student_in_roster': find_student_in_roster,
        }
        
        executor = IntegratedDSLExecutor(app_functions, session_obj=session)
        result = executor.execute_script(script_content)
        
        if result['success']:
            flash(result['message'], 'success')
            session['dsl_last_result'] = result
        else:
            flash(f"Error: {result['error']}", 'error')
            if result.get('line_num'):
                flash(f"Error at line {result['line_num']}: {result.get('line', '')}", 'error')
            session['dsl_last_result'] = result
        
        return redirect(url_for('dsl_interface'))
    
    except Exception as e:
        flash(f'Error executing DSL script: {str(e)}', 'error')
        import traceback
        print(f"DSL execution error: {traceback.format_exc()}")
        return redirect(url_for('dsl_interface'))

# API Endpoints for Next.js Frontend
@app.route('/api/roster/load', methods=['POST'])
def api_load_roster():
    """API endpoint to load roster file"""
    init_session()
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': 'No file provided'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'success': False, 'error': 'No file selected'}), 400
        
        file.seek(0)
        if file.filename.lower().endswith('.csv'):
            try:
                roster_df = pd.read_csv(file, encoding='utf-8')
            except (UnicodeDecodeError, UnicodeError):
                file.seek(0)
                roster_df = pd.read_csv(file, encoding='latin-1')
        else:
            roster_df = pd.read_excel(file, engine='openpyxl')
        
        if save_roster(roster_df):
            session['roster_loaded'] = True
            return jsonify({
                'success': True,
                'message': f'Roster loaded successfully: {len(roster_df)} students',
                'student_count': len(roster_df),
                'columns': list(roster_df.columns)
            })
        else:
            return jsonify({'success': False, 'error': 'Failed to save roster'}), 500
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/roster/info', methods=['GET'])
def api_roster_info():
    """API endpoint to get roster information"""
    init_session()
    roster_df = load_roster()
    if roster_df is None:
        return jsonify({'success': False, 'loaded': False})
    
    date_columns = [str(col) for col in roster_df.columns 
                   if any(x in str(col).lower() for x in ['nov', 'oct', 'dec', 'jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', '.'])]
    
    return jsonify({
        'success': True,
        'loaded': True,
        'student_count': len(roster_df),
        'columns': list(roster_df.columns),
        'date_columns': date_columns,
        'roster_file': app.config['ROSTER_FILE']
    })

@app.route('/api/attendance/process', methods=['POST'])
def api_process_attendance():
    """API endpoint to process attendance with Gemini"""
    init_session()
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': 'No file provided'}), 400
        
        attendance_file = request.files['file']
        date = request.form.get('date', None)
        
        # Save uploaded file temporarily
        filename = secure_filename(attendance_file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        attendance_file.save(filepath)
        
        roster_df = load_roster()
        if roster_df is None:
            return jsonify({'success': False, 'error': 'Please load a roster file first'}), 400
        
        # Read attendance file
        if filename.lower().endswith('.csv'):
            try:
                attendance_df = pd.read_csv(filepath, encoding='utf-8')
            except:
                attendance_df = pd.read_csv(filepath, encoding='latin-1')
        else:
            attendance_df = pd.read_excel(filepath, engine='openpyxl')
        
        # Prepare context for Gemini
        roster_sample = roster_df.head(10).to_string()
        attendance_sample = attendance_df.head(10).to_string()
        
        # Get Gemini model
        api_key = os.getenv('GEMINI_API_KEY')
        if not api_key:
            return jsonify({'success': False, 'error': 'Gemini API key not configured'}), 500
        
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-2.0-flash')
        
        # Create prompt
        try:
            from dsl.gemini_prompts import create_attendance_processing_prompt
            prompt = create_attendance_processing_prompt(
                roster_sample=roster_sample,
                attendance_sample=attendance_sample,
                attendance_file=filepath,
                date=date,
                roster_file=app.config['ROSTER_FILE']
            )
        except ImportError:
            prompt = f"Generate DSL code to process attendance file: {filepath}"
        
        # Generate DSL code
        response = model.generate_content(prompt)
        dsl_code = extract_clean_dsl_code(response.text)
        
        return jsonify({
            'success': True,
            'dsl_code': dsl_code,
            'attendance_file': filename
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/dsl/execute', methods=['POST'])
def api_execute_dsl():
    """API endpoint to execute DSL code"""
    init_session()
    try:
        data = request.get_json()
        dsl_code = data.get('dsl_code', '')
        
        if not dsl_code:
            return jsonify({'success': False, 'error': 'No DSL code provided'}), 400
        
        from dsl.dsl_integrated import IntegratedDSLExecutor
        
        app_functions = {
            'load_roster': load_roster,
            'save_roster': save_roster,
            'format_date_for_roster': format_date_for_roster,
            'find_matching_date_column': find_matching_date_column,
        }
        
        executor = IntegratedDSLExecutor(app_functions, session_obj=session)
        result = executor.execute_script(dsl_code)
        
        return jsonify(result)
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/query')
def query():
    """Query/View Information page"""
    init_session()
    roster_df = load_roster()
    query_result = session.get('query_result')
    # Clear query result from session after retrieving it (so it doesn't persist)
    if query_result:
        session.pop('query_result', None)
    return render_template('query.html', 
                         roster=roster_df, 
                         roster_loaded=roster_df is not None,
                         query_result=query_result)

@app.route('/query', methods=['POST'])
def process_query():
    """Process natural language query"""
    init_session()
    roster_df = load_roster()
    if roster_df is None:
        flash('Please upload a roster file first', 'warning')
        return redirect(url_for('index'))
    
    user_query = request.form.get('user_query', '').strip()
    if not user_query:
        flash('Please enter a query', 'error')
        return redirect(url_for('query'))
    
    try:
        # Get Gemini API key
        api_key = session.get('gemini_api_key') or os.getenv('GEMINI_API_KEY') or os.getenv('AIzaSyAcR924DTqb4X30QpoM98eqJ3q5IQCXtEQ')
        if not api_key:
            flash('Gemini API key not configured. Please set it in Settings.', 'error')
            return redirect(url_for('query'))
        
        genai.configure(api_key=api_key)
        
        # Try newer models first, fallback to older ones
        try:
            model = genai.GenerativeModel('gemini-2.0-flash')
        except:
            try:
                model = genai.GenerativeModel('gemini-2.5-pro')
            except:
                model = genai.GenerativeModel('gemini-pro')
        
        # Prepare context
        roster_info = f"Roster has {len(roster_df)} students"
        date_columns = [str(col) for col in roster_df.columns 
                       if any(x in str(col).lower() for x in ['nov', 'oct', 'dec', 'jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', '.'])]
        
        # Create prompt
        try:
            from dsl.gemini_prompts import create_query_prompt
            prompt = create_query_prompt(
                user_query=user_query,
                roster_info=roster_info,
                date_columns=date_columns,
                roster_file=app.config['ROSTER_FILE']
            )
        except ImportError:
            prompt = f"User request: {user_query}\nGenerate DSL code to fulfill this request."
        
        # Generate DSL code
        response = model.generate_content(prompt)
        dsl_code = extract_clean_dsl_code(response.text)
        
        if not dsl_code:
            flash('Could not generate DSL code from query. Please try rephrasing your question.', 'error')
            return redirect(url_for('query'))
        
        # Execute the DSL code
        try:
            from dsl.dsl_integrated import IntegratedDSLExecutor
            
            app_functions = {
                'load_roster': load_roster,
                'save_roster': save_roster,
                'format_date_for_roster': format_date_for_roster,
                'find_matching_date_column': find_matching_date_column,
                'find_student_in_roster': find_student_in_roster,
            }
            
            executor = IntegratedDSLExecutor(app_functions, session_obj=session)
            result = executor.execute_script(dsl_code)
            
            # Store result in session
            result['dsl_code'] = dsl_code
            session['query_result'] = result
            
            if result['success']:
                flash(f'Query processed successfully: {result.get("message", "Done")}', 'success')
            else:
                flash(f'Error processing query: {result.get("error", "Unknown error")}', 'error')
        except Exception as e:
            flash(f'Error executing query: {str(e)}', 'error')
            import traceback
            print(f"Query execution error: {traceback.format_exc()}")
            session['query_result'] = {
                'success': False,
                'error': str(e),
                'dsl_code': dsl_code
            }
    except Exception as e:
        flash(f'Error processing query: {str(e)}', 'error')
        import traceback
        print(f"Query processing error: {traceback.format_exc()}")
        session['query_result'] = {
            'success': False,
            'error': str(e)
        }
    
    return redirect(url_for('query'))

@app.route('/api/query', methods=['POST'])
def api_query():
    """API endpoint for natural language queries"""
    init_session()
    try:
        data = request.get_json()
        user_query = data.get('query', '')
        
        if not user_query:
            return jsonify({'success': False, 'error': 'No query provided'}), 400
        
        roster_df = load_roster()
        if roster_df is None:
            return jsonify({'success': False, 'error': 'Please load a roster file first'}), 400
        
        # Get Gemini model
        api_key = os.getenv('GEMINI_API_KEY')
        if not api_key:
            return jsonify({'success': False, 'error': 'Gemini API key not configured'}), 500
        
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-2.0-flash')
        
        # Prepare context
        roster_info = f"Roster has {len(roster_df)} students"
        date_columns = [str(col) for col in roster_df.columns 
                       if any(x in str(col).lower() for x in ['nov', 'oct', 'dec', 'jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', '.'])]
        
        # Create prompt
        try:
            from dsl.gemini_prompts import create_query_prompt
            prompt = create_query_prompt(
                user_query=user_query,
                roster_info=roster_info,
                date_columns=date_columns,
                roster_file=app.config['ROSTER_FILE']
            )
        except ImportError:
            prompt = f"User request: {user_query}\nGenerate DSL code to fulfill this request."
        
        # Generate DSL code
        response = model.generate_content(prompt)
        dsl_code = extract_clean_dsl_code(response.text)
        
        return jsonify({
            'success': True,
            'dsl_code': dsl_code
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/student/find', methods=['POST'])
def api_find_student():
    """API endpoint to find student's total points"""
    init_session()
    try:
        data = request.get_json()
        student_name = data.get('student_name', '')
        
        if not student_name:
            return jsonify({'success': False, 'error': 'No student name provided'}), 400
        
        roster_df = load_roster()
        if roster_df is None:
            return jsonify({'success': False, 'error': 'Please load a roster file first'}), 400
        
        # Find name column
        name_col = None
        for col in roster_df.columns:
            col_str = str(col).lower().strip()
            if ('name' in col_str and 'unnamed' not in col_str and 
                col_str not in ['id', 'email', 'major', 'level']):
                name_col = col
                break
        if name_col is None:
            if len(roster_df.columns) > 2:
                name_col = roster_df.columns[2]
        
        # Get date columns
        date_columns = []
        for col in roster_df.columns:
            col_str = str(col).lower()
            if any(x in col_str for x in ['nov', 'oct', 'dec', 'jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', '.']) and col_str not in ['total', 'points']:
                date_columns.append(str(col))
        
        # Find matching students
        matching_students = []
        if name_col:
            normalized_input = student_name.strip().replace(', ', ',').replace(' ,', ',')
            
            for idx, row in roster_df.iterrows():
                roster_name = str(row[name_col]).strip()
                if not roster_name or roster_name.lower() in ['nan', 'none', '']:
                    continue
                
                normalized_roster = roster_name.replace(', ', ',').replace(' ,', ',')
                input_lower = normalized_input.lower()
                roster_lower = normalized_roster.lower()
                
                if input_lower == roster_lower:
                    # Calculate total points
                    total = 0.0
                    attendance = {}
                    for col in date_columns:
                        val = roster_df.loc[idx, col]
                        if pd.notna(val):
                            try:
                                points = float(val)
                                total += points
                                if points > 0:
                                    attendance[col] = points
                            except (ValueError, TypeError):
                                pass
                    
                    matching_students.insert(0, {
                        'name': roster_name,
                        'total_points': roster_df.loc[idx, 'Total Points'] if 'Total Points' in roster_df.columns else total,
                        'calculated_total': total,
                        'attendance': attendance
                    })
                elif input_lower in roster_lower or roster_lower in input_lower:
                    total = 0.0
                    attendance = {}
                    for col in date_columns:
                        val = roster_df.loc[idx, col]
                        if pd.notna(val):
                            try:
                                points = float(val)
                                total += points
                                if points > 0:
                                    attendance[col] = points
                            except (ValueError, TypeError):
                                pass
                    
                    matching_students.append({
                        'name': roster_name,
                        'total_points': roster_df.loc[idx, 'Total Points'] if 'Total Points' in roster_df.columns else total,
                        'calculated_total': total,
                        'attendance': attendance
                    })
        
        # No need to call Gemini - we already have the student information
        # Return empty dsl_code since student lookup is done directly
        return jsonify({
            'success': True,
            'students': matching_students,
            'dsl_code': ''
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# Enable CORS for Next.js frontend
@app.after_request
def after_request(response):
    response.headers.add('Access-Control-Allow-Origin', 'http://localhost:3000')
    response.headers.add('Access-Control-Allow-Headers', 'Content-Type,Authorization')
    response.headers.add('Access-Control-Allow-Methods', 'GET,PUT,POST,DELETE,OPTIONS')
    return response

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5001)
