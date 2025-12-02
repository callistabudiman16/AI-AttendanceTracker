import streamlit as st
import pandas as pd
import qrcode
from PIL import Image
import io
import cv2
from pyzbar import pyzbar
from datetime import datetime
import os
import re
from difflib import SequenceMatcher
import requests
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

# Page configuration
st.set_page_config(page_title="Attendance Tracker", page_icon="📋", layout="wide")

# Initialize session state
if 'roster' not in st.session_state:
    st.session_state.roster = None
if 'attendance_records' not in st.session_state:
    st.session_state.attendance_records = []
if 'roster_file_path' not in st.session_state:
    st.session_state.roster_file_path = 'roster_attendance.xlsx'  # Default persistent roster file

# Initialize Gemini API key from environment variable if available
if 'gemini_api_key' not in st.session_state:
    # Try to load from environment variable first (more secure)
    st.session_state.gemini_api_key = os.getenv('GEMINI_API_KEY', '')

# Initialize OneDrive configuration
if 'onedrive_connected' not in st.session_state:
    st.session_state.onedrive_connected = False
if 'onedrive_client_id' not in st.session_state:
    st.session_state.onedrive_client_id = os.getenv('ONEDRIVE_CLIENT_ID', '')
if 'onedrive_client_secret' not in st.session_state:
    st.session_state.onedrive_client_secret = os.getenv('ONEDRIVE_CLIENT_SECRET', '')
if 'onedrive_access_token' not in st.session_state:
    st.session_state.onedrive_access_token = None
if 'onedrive_file_id' not in st.session_state:
    st.session_state.onedrive_file_id = None
if 'onedrive_file_path' not in st.session_state:
    st.session_state.onedrive_file_path = 'roster_attendance.xlsx'  # Default OneDrive file name

def normalize_name_for_roster(name):
    """Convert 'Last, First' or 'First, Last' to 'First Last' for matching with roster"""
    if ',' in name:
        parts = [p.strip() for p in name.split(',')]
        if len(parts) == 2:
            # Try both orderings: "Last, First" and "First, Last"
            # Check which order makes more sense by trying both
            return f"{parts[1]} {parts[0]}"  # Assume "Last, First" format
    return name

def normalize_name_for_zoom(name):
    """Convert 'First Last' to 'Last, First' if needed"""
    if ',' not in name:
        parts = name.split()
        if len(parts) >= 2:
            return f"{parts[-1]}, {' '.join(parts[:-1])}"
    return name

def extract_name_components(name):
    """Extract first, middle, last name components from various formats"""
    name = name.strip()
    components = {'first': '', 'middle': '', 'last': '', 'middle_initial': ''}
    
    if ',' in name:
        # Format: "Last, First" or "Last, First Middle"
        parts = [p.strip() for p in name.split(',')]
        if len(parts) == 2:
            last_part = parts[0].strip()
            first_middle = parts[1].strip().split()
            
            components['last'] = last_part
            
            if len(first_middle) >= 1:
                components['first'] = first_middle[0]
            if len(first_middle) >= 2:
                # Could be middle name or middle initial
                middle = first_middle[1]
                if len(middle) == 1 or (len(middle) == 2 and middle.endswith('.')):
                    components['middle_initial'] = middle.replace('.', '')
                else:
                    components['middle'] = middle
                # Also store full middle part
                if len(first_middle) > 2:
                    components['middle'] = ' '.join(first_middle[1:])
    else:
        # Format: "First Last" or "First Middle Last" or "First M Last"
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

def get_all_name_variations(name):
    """Get all possible name format variations for flexible matching, including partial matches"""
    variations = [name]  # Original
    components = extract_name_components(name)
    
    # Build variations based on components
    first = components['first']
    middle = components['middle']
    middle_initial = components['middle_initial']
    last = components['last']
    
    if not first or not last:
        # If we couldn't parse, fall back to simple variations
        if ',' in name:
            parts = [p.strip() for p in name.split(',')]
            if len(parts) == 2:
                variations.append(f"{parts[1]} {parts[0]}")
                variations.append(f"{parts[0]} {parts[1]}")
        else:
            parts = name.split()
            if len(parts) >= 2:
                variations.append(f"{parts[-1]}, {' '.join(parts[:-1])}")
                variations.append(f"{parts[0]}, {' '.join(parts[1:])}")
    else:
        # Full name variations
        # "Last, First" format
        if middle:
            variations.append(f"{last}, {first} {middle}")
        if middle_initial:
            variations.append(f"{last}, {first} {middle_initial}")
            variations.append(f"{last}, {first} {middle_initial}.")
        variations.append(f"{last}, {first}")
        
        # "First Last" format
        if middle:
            variations.append(f"{first} {middle} {last}")
        if middle_initial:
            variations.append(f"{first} {middle_initial} {last}")
            variations.append(f"{first} {middle_initial}. {last}")
        variations.append(f"{first} {last}")
        
        # Partial matches (without middle)
        variations.append(f"{first} {last}")
        variations.append(f"{last}, {first}")
        
        # With middle initial variations
        if middle:
            # Extract first letter of middle name
            if middle and len(middle) > 0:
                mi = middle[0].upper()
                variations.append(f"{first} {mi} {last}")
                variations.append(f"{first} {mi}. {last}")
                variations.append(f"{last}, {first} {mi}")
                variations.append(f"{last}, {first} {mi}.")
        if middle_initial:
            variations.append(f"{first} {middle_initial} {last}")
            variations.append(f"{last}, {first} {middle_initial}")
    
    # Remove duplicates while preserving order
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
    
    # Exact match on components
    if (check_components['first'].lower() == roster_components['first'].lower() and
        check_components['last'].lower() == roster_components['last'].lower()):
        
        # Check middle name/initial match
        check_middle = check_components.get('middle', '') or check_components.get('middle_initial', '')
        roster_middle = roster_components.get('middle', '') or roster_components.get('middle_initial', '')
        
        # If both have middle info, they should match (or at least initial should match)
        if not check_middle or not roster_middle:
            # One doesn't have middle - still a match (partial match)
            return 0.95  # High confidence for partial match
        elif check_middle.lower() == roster_middle.lower():
            return 1.0  # Perfect match
        elif (check_components.get('middle_initial', '').lower() == 
              (roster_components.get('middle_initial', '') or 
               (roster_components.get('middle', '')[0] if roster_components.get('middle') else '')).lower()):
            return 0.92  # Middle initial matches
        elif (roster_components.get('middle_initial', '').lower() == 
              (check_components.get('middle_initial', '') or 
               (check_components.get('middle', '')[0] if check_components.get('middle') else '')).lower()):
            return 0.92  # Middle initial matches (reverse)
    
    return 0.0

def calculate_similarity(str1, str2):
    """Calculate similarity ratio between two strings (0.0 to 1.0)"""
    return SequenceMatcher(None, str1.lower().strip(), str2.lower().strip()).ratio()

def find_student_in_roster(student_name, roster_df, use_gemini=False, min_confidence=0.75):
    """
    Find student in roster by matching names (handles different formats)
    Returns: (index, confidence_score, matched_name) or (None, 0, None)
    """
    # Get all possible name format variations
    # This handles cases like "First, Last" vs "Last, First" and typos
    name_variations = get_all_name_variations(student_name)
    # Also add the original normalization methods for compatibility
    name_variations.extend([
        normalize_name_for_roster(student_name),
        normalize_name_for_zoom(student_name)
    ])
    # Remove duplicates while preserving order
    seen = set()
    unique_variations = []
    for v in name_variations:
        v_lower = v.lower().strip()
        if v_lower not in seen:
            seen.add(v_lower)
            unique_variations.append(v)
    name_variations = unique_variations
    
    # Check if roster has a name column
    name_col = None
    for col in roster_df.columns:
        if 'name' in col.lower() or 'student' in col.lower():
            name_col = col
            break
    
    if name_col is None:
        # Try to find first text column
        for col in roster_df.columns:
            if roster_df[col].dtype == 'object':
                name_col = col
                break
    
    if name_col is None:
        return None, 0.0, None
    
    best_match = None
    best_confidence = 0.0
    best_matched_name = None
    
    # First, try component-based matching for all roster names (handles partial names and middle initials)
    for idx, row in roster_df.iterrows():
        roster_name = str(row[name_col]).strip()
        
        # Try component-based matching (handles partial names, middle initials)
        component_confidence = match_name_with_components(student_name, roster_name)
        if component_confidence > best_confidence:
            best_confidence = component_confidence
            best_match = idx
            best_matched_name = roster_name
    
    # Try exact and fuzzy matches with variations
    for name_var in name_variations:
        # Exact match (case-insensitive)
        exact_matches = roster_df[roster_df[name_col].str.lower().str.strip() == name_var.lower().strip()]
        if not exact_matches.empty:
            return exact_matches.index[0], 1.0, exact_matches.iloc[0][name_col]
        
        # Fuzzy matching for all roster names
        for idx, row in roster_df.iterrows():
            roster_name = str(row[name_col]).strip()
            similarity = calculate_similarity(name_var, roster_name)
            
            if similarity > best_confidence:
                best_confidence = similarity
                best_match = idx
                best_matched_name = roster_name
        
        # Split name and try partial matches (for better matching)
        if ',' in name_var:
            last, first = [p.strip() for p in name_var.split(',')]
            # Try matching first and last separately
            matches = roster_df[
                (roster_df[name_col].str.contains(first, case=False, na=False)) &
                (roster_df[name_col].str.contains(last, case=False, na=False))
            ]
            if not matches.empty:
                # Check similarity for each match
                for idx in matches.index:
                    roster_name = str(matches.loc[idx, name_col]).strip()
                    similarity = calculate_similarity(name_var, roster_name)
                    if similarity > best_confidence:
                        best_confidence = similarity
                        best_match = idx
                        best_matched_name = roster_name
        else:
            parts = name_var.split()
            if len(parts) >= 2:
                matches = roster_df[
                    (roster_df[name_col].str.contains(parts[0], case=False, na=False)) &
                    (roster_df[name_col].str.contains(parts[-1], case=False, na=False))
                ]
                if not matches.empty:
                    # Check similarity for each match
                    for idx in matches.index:
                        roster_name = str(matches.loc[idx, name_col]).strip()
                        similarity = calculate_similarity(name_var, roster_name)
                        if similarity > best_confidence:
                            best_confidence = similarity
                            best_match = idx
                            best_matched_name = roster_name
    
    # Try Gemini API if enabled (especially useful for complex name matching with middle names/initials)
    # Use Gemini even if we have a match, as it's better at handling partial names and variations
    if use_gemini and GEMINI_AVAILABLE:
        gemini_match = find_student_with_gemini(student_name, roster_df, name_col)
        if gemini_match:
            gemini_idx, gemini_conf, gemini_matched = gemini_match
            # Prefer Gemini result if it has higher confidence or if current match is below threshold
            if gemini_conf > best_confidence or best_confidence < min_confidence:
                return gemini_match
    elif use_gemini and not GEMINI_AVAILABLE:
        # Gemini was requested but not available - continue with regular matching
        pass
    
    # Return best match if confidence is high enough
    if best_confidence >= min_confidence:
        return best_match, best_confidence, best_matched_name
    
    return None, best_confidence, None

def find_student_with_gemini(student_name, roster_df, name_col):
    """Use Gemini API to find student in roster - especially good for partial names and middle initials"""
    if 'gemini_api_key' not in st.session_state or not st.session_state.gemini_api_key:
        return None
    
    try:
        genai.configure(api_key=st.session_state.gemini_api_key)
        model = genai.GenerativeModel('gemini-pro')
        
        # Get all roster names
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
        matched_name = response.text.strip()
        
        # Clean up response (remove quotes, extra whitespace, etc.)
        matched_name = matched_name.strip('"\'`').strip()
        
        if matched_name.upper() == "NO_MATCH" or not matched_name:
            return None
        
        # Find the matched name in roster (try exact match first, then case-insensitive)
        matches = roster_df[roster_df[name_col].astype(str).str.strip() == matched_name]
        if matches.empty:
            # Try case-insensitive match
            matches = roster_df[roster_df[name_col].astype(str).str.strip().str.lower() == matched_name.lower()]
        
        if not matches.empty:
            return matches.index[0], 0.95, matches.iloc[0][name_col]  # High confidence for Gemini matches
        
    except Exception as e:
        # Don't show warning on every call - only log silently or on first error
        if 'gemini_error_logged' not in st.session_state:
            st.session_state.gemini_error_logged = True
            st.warning(f"Gemini API error: {str(e)}")
    
    return None

def update_roster_with_attendance(roster_df, student_name, points, date_str, use_gemini=False, min_confidence=0.75):
    """Update roster with attendance points"""
    idx, confidence, matched_name = find_student_in_roster(student_name, roster_df, use_gemini, min_confidence)
    if idx is None:
        return roster_df, False, confidence, matched_name
    
    # Create date column if it doesn't exist
    if date_str not in roster_df.columns:
        roster_df[date_str] = 0.0
    
    # Update points (accumulate if already exists)
    current_points = roster_df.loc[idx, date_str]
    if pd.isna(current_points) or current_points == 0:
        roster_df.loc[idx, date_str] = points
    else:
        # If already has points, keep the higher value
        roster_df.loc[idx, date_str] = max(current_points, points)
    
    return roster_df, True, confidence, matched_name

def save_roster_to_file(roster_df, file_path, show_errors=True):
    """Save roster DataFrame to Excel file and sync to OneDrive if connected"""
    try:
        # Save locally
        roster_df.to_excel(file_path, index=False, engine='openpyxl')
        
        # Sync to OneDrive if connected (failures are non-fatal)
        if st.session_state.onedrive_connected:
            try:
                sync_roster_to_onedrive(roster_df)
            except Exception as onedrive_error:
                if show_errors:
                    st.warning(f"⚠️ Saved locally, but OneDrive sync failed: {str(onedrive_error)}")
        
        return True
    except Exception as e:
        if show_errors:
            st.error(f"Error saving roster file: {str(e)}")
        return False

def load_roster_from_file(file_path):
    """Load roster DataFrame from Excel file, checking OneDrive first if connected"""
    try:
        # Try OneDrive first if connected
        if st.session_state.onedrive_connected:
            onedrive_roster = load_roster_from_onedrive()
            if onedrive_roster is not None:
                # Also save locally for backup
                try:
                    onedrive_roster.to_excel(file_path, index=False, engine='openpyxl')
                except:
                    pass  # If local save fails, still return OneDrive version
                return onedrive_roster
        
        # Fall back to local file
        if os.path.exists(file_path):
            roster_df = pd.read_excel(file_path, engine='openpyxl')
            return roster_df
        return None
    except Exception as e:
        st.error(f"Error loading roster file: {str(e)}")
        return None

def process_zoom_attendance(roster_df, zoom_df, date_str):
    """Process Zoom attendance from Excel file"""
    # Find name and duration columns in Zoom file
    name_col = None
    duration_col = None
    
    for col in zoom_df.columns:
        col_lower = col.lower()
        if 'name' in col_lower or 'participant' in col_lower:
            name_col = col
        if 'duration' in col_lower or 'time' in col_lower or 'length' in col_lower:
            duration_col = col
    
    if name_col is None:
        st.error("Could not find name column in Zoom file. Please ensure the file has a column with participant names.")
        return roster_df
    
    if duration_col is None:
        st.error("Could not find duration column in Zoom file. Please ensure the file has a column with meeting duration.")
        return roster_df
    
    updated_count = 0
    errors = []
    
    # Create date column if it doesn't exist
    if date_str not in roster_df.columns:
        roster_df[date_str] = 0.0
    
    for idx, row in zoom_df.iterrows():
        student_name = str(row[name_col]).strip()
        duration_str = str(row[duration_col]).strip()
        
        # Parse duration (handle formats like "1:30:45" or "90:30" or "90 minutes")
        duration_minutes = parse_duration(duration_str)
        
        if duration_minutes is None:
            errors.append(f"Could not parse duration for {student_name}: {duration_str}")
            continue
        
        # Calculate points
        if duration_minutes >= 30:
            points = 0.6
        elif duration_minutes > 0:
            points = 0.2
        else:
            points = 0.0
        
        # Update roster (use Gemini if enabled)
        use_gemini_flag = st.session_state.get('use_gemini', False)
        min_conf = st.session_state.get('min_confidence', 0.75)
        roster_df, found, confidence, matched_name = update_roster_with_attendance(
            roster_df, student_name, points, date_str, use_gemini_flag, min_conf
        )
        if found:
            updated_count += 1
            if confidence < 1.0:
                # Log low confidence matches
                st.info(f"⚠️ Low confidence match: '{student_name}' → '{matched_name}' (confidence: {confidence:.2f})")
        else:
            errors.append(f"Student not found in roster: {student_name} (best match confidence: {confidence:.2f})")
    
    if errors:
        st.warning(f"Some issues encountered:\n" + "\n".join(errors[:10]))
    
    st.success(f"Updated attendance for {updated_count} students from Zoom file.")
    return roster_df

def parse_duration(duration_str):
    """Parse duration string to minutes"""
    # Remove any non-digit characters except : and .
    duration_str = re.sub(r'[^\d:.]', '', duration_str)
    
    # Try different formats
    # Format: HH:MM:SS or MM:SS
    if ':' in duration_str:
        parts = duration_str.split(':')
        if len(parts) == 3:  # HH:MM:SS
            hours, minutes, seconds = map(int, parts)
            return hours * 60 + minutes + seconds / 60
        elif len(parts) == 2:  # MM:SS
            minutes, seconds = map(int, parts)
            return minutes + seconds / 60
        elif len(parts) == 1:  # Just minutes
            return int(parts[0])
    else:
        # Try to parse as number (assume minutes)
        try:
            return float(duration_str)
        except:
            pass
    
    return None

# OneDrive API Integration Functions
def get_onedrive_access_token(client_id, client_secret, redirect_uri="http://localhost:8501"):
    """Get Microsoft Graph API access token using device code flow"""
    if not MSAL_AVAILABLE:
        st.error("⚠️ MSAL library not installed. Run: `pip install msal`")
        return None
    
    try:
        # Microsoft Graph API endpoint
        authority = "https://login.microsoftonline.com/common"
        scope = ["Files.ReadWrite.All"]
        
        # Create MSAL app
        app = msal.PublicClientApplication(
            client_id=client_id,
            authority=authority
        )
        
        # Try to get token from cache first
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(scope, account=accounts[0])
            if result and "access_token" in result:
                return result["access_token"]
        
        # If no cached token, use device code flow (better for Streamlit)
        flow = app.initiate_device_flow(scopes=scope)
        if "user_code" not in flow:
            st.error("⚠️ Failed to initiate device code flow. Please check your Client ID.")
            return None
        
        # Display device code to user
        verification_uri = flow.get('verification_uri', 'https://microsoft.com/devicelogin')
        user_code = flow['user_code']
        message_placeholder = st.empty()
        message_placeholder.info(
            f"📱 **Device Code:** `{user_code}`\n\n"
            f"1. Open this URL: **{verification_uri}**\n"
            f"2. Enter the code: **{user_code}**\n"
            f"3. Grant permissions to access OneDrive\n\n"
            f"⏳ Waiting for authentication (this may take up to 2 minutes)..."
        )
        
        # Poll for token (with timeout)
        import time
        max_wait_time = 120  # 2 minutes
        start_time = time.time()
        
        while time.time() - start_time < max_wait_time:
            try:
                result = app.acquire_token_by_device_flow(flow)
                
                if "access_token" in result:
                    message_placeholder.empty()
                    return result["access_token"]
                elif "error" in result:
                    if result["error"] == "authorization_pending":
                        # Still waiting - continue polling
                        time.sleep(5)  # Wait 5 seconds before checking again
                        continue
                    else:
                        message_placeholder.empty()
                        error_msg = result.get('error_description', result.get('error', 'Unknown error'))
                        st.error(f"❌ Authentication failed: {error_msg}")
                        return None
                else:
                    # Still waiting
                    time.sleep(5)
                    continue
            except Exception as e:
                message_placeholder.empty()
                st.error(f"❌ Error during authentication: {str(e)}")
                return None
        
        # Timeout
        message_placeholder.empty()
        st.error("⏰ Authentication timeout. Please try again.")
        return None
            
    except Exception as e:
        st.error(f"❌ OneDrive authentication error: {str(e)}")
        return None

def find_onedrive_file(access_token, file_name):
    """Find a file in OneDrive root folder"""
    try:
        url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{file_name}"
        headers = {"Authorization": f"Bearer {access_token}"}
        
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            file_info = response.json()
            return file_info.get('id'), file_info
        elif response.status_code == 404:
            return None, None
        else:
            st.error(f"Error finding OneDrive file: {response.status_code} - {response.text}")
            return None, None
    except Exception as e:
        st.error(f"Error searching OneDrive: {str(e)}")
        return None, None

def upload_to_onedrive(roster_df, access_token, file_id, file_name):
    """Upload roster file to OneDrive"""
    try:
        # Convert DataFrame to Excel bytes
        output = io.BytesIO()
        roster_df.to_excel(output, index=False, engine='openpyxl')
        file_content = output.getvalue()
        
        if file_id:
            # Update existing file
            url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/content"
        else:
            # Create new file
            url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{file_name}:/content"
        
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        }
        
        response = requests.put(url, headers=headers, data=file_content)
        
        if response.status_code in [200, 201]:
            file_info = response.json()
            return True, file_info.get('id')
        else:
            st.error(f"Error uploading to OneDrive: {response.status_code} - {response.text}")
            return False, None
            
    except Exception as e:
        st.error(f"Error uploading to OneDrive: {str(e)}")
        return False, None

def download_from_onedrive(access_token, file_id):
    """Download roster file from OneDrive"""
    try:
        url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/content"
        headers = {"Authorization": f"Bearer {access_token}"}
        
        response = requests.get(url, headers=headers)
        
        if response.status_code == 200:
            file_content = io.BytesIO(response.content)
            roster_df = pd.read_excel(file_content, engine='openpyxl')
            return roster_df
        else:
            st.error(f"Error downloading from OneDrive: {response.status_code} - {response.text}")
            return None
            
    except Exception as e:
        st.error(f"Error downloading from OneDrive: {str(e)}")
        return None

def sync_roster_to_onedrive(roster_df):
    """Sync roster to OneDrive if connected"""
    if not st.session_state.onedrive_connected or not st.session_state.onedrive_access_token:
        return False
    
    access_token = st.session_state.onedrive_access_token
    file_name = st.session_state.onedrive_file_path
    
    # Find or create file
    file_id, file_info = find_onedrive_file(access_token, file_name)
    
    # Upload file
    success, new_file_id = upload_to_onedrive(roster_df, access_token, file_id, file_name)
    
    if success:
        st.session_state.onedrive_file_id = new_file_id or file_id
        return True
    return False

def load_roster_from_onedrive():
    """Load roster from OneDrive if connected"""
    if not st.session_state.onedrive_connected or not st.session_state.onedrive_access_token:
        return None
    
    if not st.session_state.onedrive_file_id:
        # Try to find file
        access_token = st.session_state.onedrive_access_token
        file_name = st.session_state.onedrive_file_path
        file_id, file_info = find_onedrive_file(access_token, file_name)
        
        if not file_id:
            return None
        
        st.session_state.onedrive_file_id = file_id
    
    # Download file
    access_token = st.session_state.onedrive_access_token
    file_id = st.session_state.onedrive_file_id
    roster_df = download_from_onedrive(access_token, file_id)
    
    return roster_df

def scan_qr_code():
    """Scan QR code from webcam"""
    cap = cv2.VideoCapture(0)
    
    if not cap.isOpened():
        st.error("Could not open webcam")
        return None
    
    st.info("Position QR code in front of camera. Press 'q' to quit.")
    frame_placeholder = st.empty()
    
    while True:
        ret, frame = cap.read()
        if not ret:
            break
        
        # Decode QR codes
        qr_codes = pyzbar.decode(frame)
        
        # Draw rectangles and decode
        for qr in qr_codes:
            # Extract QR code data
            data = qr.data.decode('utf-8')
            
            # Draw rectangle around QR code
            points = qr.polygon
            if len(points) == 4:
                pts = [(point.x, point.y) for point in points]
                for i in range(4):
                    cv2.line(frame, pts[i], pts[(i+1)%4], (0, 255, 0), 2)
            
            cap.release()
            cv2.destroyAllWindows()
            return data
        
        # Display frame
        frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        frame_placeholder.image(frame_rgb, channels="RGB", use_container_width=True)
    
    cap.release()
    cv2.destroyAllWindows()
    return None

# Main App
st.title("📋 Attendance Tracker App")
st.caption("For Professors and Teaching Assistants Only - Manage student attendance and update roster files")

# Auto-load roster on startup
if st.session_state.roster is None:
    existing_roster = load_roster_from_file(st.session_state.roster_file_path)
    if existing_roster is not None:
        st.session_state.roster = existing_roster
        st.session_state.class_start_time = datetime.now().replace(hour=9, minute=0, second=0, microsecond=0).time()
        st.session_state.late_threshold_minutes = 15

# Sidebar for settings
with st.sidebar:
    st.header("Settings")
    
    # Late threshold
    late_threshold_minutes = st.number_input(
        "Late Threshold (minutes after class start)", 
        min_value=0, 
        value=15,
        help="Students checking in after this time will receive 0.2 points instead of 0.6"
    )
    
    # Class start time
    class_start_time = st.time_input(
        "Class Start Time",
        value=datetime.now().replace(hour=9, minute=0, second=0, microsecond=0).time()
    )
    
    st.divider()
    
    # Name Matching Settings
    st.header("Name Matching")
    
    if 'min_confidence' not in st.session_state:
        st.session_state.min_confidence = 0.75
    
    min_confidence = st.slider(
        "Minimum Confidence Threshold",
        min_value=0.5,
        max_value=1.0,
        value=st.session_state.min_confidence,
        step=0.05,
        help="Higher values require more exact name matches. Lower values allow more fuzzy matching."
    )
    st.session_state.min_confidence = min_confidence
    
    # Gemini API configuration (recommended for complex name matching)
    st.subheader("AI-Assisted Matching (Recommended)")
    st.info("💡 **Recommended for:** Matching partial names, middle initials, and name variations.\n"
            "Example: Roster has 'Budiman, Natasha Callista' but attendance has 'Natasha Budiman' or 'Budiman, Natasha C'")
    
    if 'use_gemini' not in st.session_state:
        st.session_state.use_gemini = False
    
    use_gemini = st.checkbox(
        "✅ Enable Gemini API for smart name matching",
        value=st.session_state.use_gemini,
        help="Uses Google's Gemini AI to intelligently match names, especially useful for:\n"
             "- Partial names (missing middle names)\n"
             "- Middle initial variations (C vs Callista)\n"
             "- Different name formats\n"
             "Free API tier available at https://makersuite.google.com/app/apikey"
    )
    st.session_state.use_gemini = use_gemini
    
    if use_gemini:
        # Check if API key is loaded from environment variable
        env_key = os.getenv('GEMINI_API_KEY', '')
        if env_key and not st.session_state.gemini_api_key:
            st.session_state.gemini_api_key = env_key
        
        gemini_api_key = st.text_input(
            "Gemini API Key",
            value=st.session_state.gemini_api_key,
            type="password",
            help="Get your free API key from: https://makersuite.google.com/app/apikey\n\n"
                 "💡 Tip: You can also set the GEMINI_API_KEY environment variable instead of entering it here."
        )
        st.session_state.gemini_api_key = gemini_api_key
        
        # Show status
        if env_key and gemini_api_key == env_key:
            st.info("🔐 API key loaded from environment variable (GEMINI_API_KEY)")
        
        if not GEMINI_AVAILABLE:
            st.warning("⚠️ Install google-generativeai package: `pip install google-generativeai`")
        elif not gemini_api_key:
            st.info("💡 Enter your Gemini API key above, or set the GEMINI_API_KEY environment variable. Free tier available!")
        else:
            st.success("✅ Gemini API configured - Ready for smart name matching!")
    
    st.divider()
    
    # Roster Management
    st.header("Roster Management")
    
    # Try to load existing roster file first
    if st.session_state.roster is None:
        existing_roster = load_roster_from_file(st.session_state.roster_file_path)
        if existing_roster is not None:
            st.session_state.roster = existing_roster
            st.info(f"📂 Loaded existing roster from {st.session_state.roster_file_path} ({len(existing_roster)} students)")
    
    # Option to upload new/initial roster file
    roster_file = st.file_uploader(
        "Upload/Replace Roster File (Excel or CSV)",
        type=['xlsx', 'xls', 'csv'],
        help="Upload your initial student roster file. The app will automatically save and update this file with all attendance records."
    )
    
    if roster_file is not None:
        try:
            if roster_file.name.endswith('.csv'):
                roster_df = pd.read_csv(roster_file)
            else:
                roster_df = pd.read_excel(roster_file)
            
            st.session_state.roster = roster_df
            st.session_state.class_start_time = class_start_time
            st.session_state.late_threshold_minutes = late_threshold_minutes
            
            # Save to persistent file
            if save_roster_to_file(roster_df, st.session_state.roster_file_path):
                st.success(f"✅ Roster loaded and saved: {len(roster_df)} students")
            else:
                st.success(f"Roster loaded: {len(roster_df)} students")
            
            st.dataframe(roster_df.head())
        except Exception as e:
            st.error(f"Error loading roster: {str(e)}")
    
    # Display current roster info
    if st.session_state.roster is not None:
        storage_location = "OneDrive (auto-sync)" if st.session_state.onedrive_connected else st.session_state.roster_file_path
        st.info(f"**Current Roster:** {len(st.session_state.roster)} students\n\n**Saved to:** {storage_location}\n\nAll attendance updates are automatically saved to this file.")
    
    st.divider()
    
    # OneDrive Integration
    st.header("OneDrive Sync (Real-Time Collaboration)")
    st.info("🔗 **Enable OneDrive sync for:**\n"
            "- Real-time collaboration (multiple professors/TAs)\n"
            "- Automatic multi-device sync\n"
            "- No manual download/upload needed")
    
    if not MSAL_AVAILABLE:
        st.warning("⚠️ Install msal package: `pip install msal`")
    else:
        # OneDrive connection status
        if st.session_state.onedrive_connected:
            st.success("✅ OneDrive Connected - Real-time sync enabled!")
            
            if st.button("🔄 Sync Now", help="Manually sync roster to OneDrive"):
                if st.session_state.roster is not None:
                    if sync_roster_to_onedrive(st.session_state.roster):
                        st.success("✅ Roster synced to OneDrive!")
                    else:
                        st.error("❌ Failed to sync to OneDrive")
            
            if st.button("❌ Disconnect OneDrive"):
                st.session_state.onedrive_connected = False
                st.session_state.onedrive_access_token = None
                st.session_state.onedrive_file_id = None
                st.rerun()
            
            # OneDrive file path
            onedrive_file_path = st.text_input(
                "OneDrive File Name",
                value=st.session_state.onedrive_file_path,
                help="Name of the file in OneDrive (will be created if it doesn't exist)"
            )
            st.session_state.onedrive_file_path = onedrive_file_path
            
        else:
            st.info("💡 Connect to OneDrive for real-time collaboration and multi-device sync")
            
            # Client ID and Secret input (or from environment)
            client_id = st.text_input(
                "Azure App Client ID",
                value=st.session_state.onedrive_client_id,
                help="Get this from Azure Portal after registering your app (or set ONEDRIVE_CLIENT_ID env var)",
                type="default"
            )
            st.session_state.onedrive_client_id = client_id
            
            if client_id:
                if st.button("🔗 Connect to OneDrive", type="primary"):
                    if not client_id:
                        st.error("Please enter Azure App Client ID")
                    else:
                        with st.spinner("Authenticating with OneDrive..."):
                            access_token = get_onedrive_access_token(client_id, "")
                            
                            if access_token:
                                st.session_state.onedrive_access_token = access_token
                                st.session_state.onedrive_connected = True
                                
                                # Try to find existing file
                                file_name = st.session_state.onedrive_file_path
                                file_id, file_info = find_onedrive_file(access_token, file_name)
                                if file_id:
                                    st.session_state.onedrive_file_id = file_id
                                    st.success(f"✅ Connected! Found existing file: {file_name}")
                                    # Try to load roster from OneDrive
                                    onedrive_roster = load_roster_from_onedrive()
                                    if onedrive_roster is not None:
                                        st.session_state.roster = onedrive_roster
                                        st.info(f"📂 Loaded roster from OneDrive: {len(onedrive_roster)} students")
                                else:
                                    st.success(f"✅ Connected! Will create new file: {file_name}")
                                
                                st.rerun()
                            else:
                                st.error("Failed to authenticate. Please check your Client ID and try again.")
            
            st.markdown("""
            **📋 Setup Instructions:**
            1. Register app in [Azure Portal](https://portal.azure.com)
            2. Add "Files.ReadWrite.All" API permission
            3. Enable "Public client flows" 
            4. Enter your Client ID above and click "Connect"
            5. Follow device code authentication steps
            """)

# Main content area
tab1, tab2, tab3 = st.tabs(["📸 In-Person Check-In", "💻 Zoom Attendance", "📊 View Roster"])

with tab1:
    st.header("In-Person Attendance Check-In")
    
    if st.session_state.roster is None:
        st.warning("⚠️ Please upload a roster file in the sidebar first.")
    else:
        # QR Code Display
        st.subheader("1. Generate QR Code for Students")
        
        # Generate QR code with check-in URL
        # This QR code links to the Qualtrics check-in form
        qr_data = st.text_input(
            "QR Code URL (Qualtrics check-in form)",
            value="https://youruniversity.qualtrics.com/your-form",
            help="Enter your Qualtrics check-in form URL. Students will scan this QR code to check in."
        )
        
        if qr_data:
            qr = qrcode.QRCode(version=1, box_size=10, border=5)
            qr.add_data(qr_data)
            qr.make(fit=True)
            qr_img = qr.make_image(fill_color="black", back_color="white")
        
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.image(qr_img, caption="Display this QR code for students to scan", use_container_width=True)
        
        st.info("💡 Display this QR code in class. Students scan it to check in via Qualtrics. After class, export responses from Qualtrics and import the data below.")
        
        st.divider()
        
        # Import check-in data from file
        st.subheader("2. Import Check-In Data")
        st.write("After students check in via Qualtrics, export the responses and import the data here:")
        st.info("💡 **Smart Name Matching:** The app handles complex name matching, including:\n"
                "- Partial names (e.g., 'Natasha Budiman' matches 'Budiman, Natasha Callista')\n"
                "- Middle initial variations (e.g., 'Natasha C Budiman' matches 'Budiman, Natasha Callista')\n"
                "- Different formats ('Last, First' vs 'First Last')\n"
                "- Minor typos and spelling variations\n\n"
                "**💡 Tip:** Enable Gemini API in the sidebar for even better matching of partial names and middle initials!")
        
        checkin_import_file = st.file_uploader(
            "Upload Qualtrics Export File (CSV or Excel)",
            type=['csv', 'xlsx', 'xls'],
            help="Upload your Qualtrics response export file. File should have a column with student names (format 'Last Name, First Name') and optionally a timestamp column."
        )
        
        date_str = datetime.now().strftime("%Y-%m-%d")
        checkin_file = f"checkins_{date_str}.csv"
        
        checkins_df = None
        
        # Load from uploaded file or existing file
        if checkin_import_file is not None:
            try:
                if checkin_import_file.name.endswith('.csv'):
                    checkins_df = pd.read_csv(checkin_import_file)
                else:
                    checkins_df = pd.read_excel(checkin_import_file)
                st.success(f"Loaded {len(checkins_df)} check-ins from file")
                
                # Auto-detect name column
                name_column = None
                for col in checkins_df.columns:
                    col_lower = col.lower()
                    if 'name' in col_lower or 'student' in col_lower:
                        name_column = col
                        break
                
                if name_column is None:
                    # Use first text column as fallback
                    for col in checkins_df.columns:
                        if checkins_df[col].dtype == 'object':
                            name_column = col
                            break
                
                if name_column:
                    st.info(f"📝 Detected name column: **{name_column}**")
                else:
                    st.error("Could not detect name column. Please ensure your file has a column with student names.")
                
                st.dataframe(checkins_df.head(10))
            except Exception as e:
                st.error(f"Error loading file: {str(e)}")
        elif os.path.exists(checkin_file):
            checkins_df = pd.read_csv(checkin_file)
            name_column = 'Name'  # Default for existing files
            st.info(f"Found existing check-in file for today with {len(checkins_df)} entries")
        
        if checkins_df is not None and not checkins_df.empty:
            st.subheader("3. Process Check-Ins and Update Roster")
            
            # Show preview
            st.write("Preview of check-ins to process:")
            st.dataframe(checkins_df, use_container_width=True)
            
            # Process each check-in
            if st.button("Process Check-Ins and Update Roster", type="primary"):
                if name_column is None:
                    st.error("❌ Could not detect name column. Please check your file format.")
                else:
                    processed_count = 0
                    errors = []
                    low_confidence_matches = []
                    
                    for idx, row in checkins_df.iterrows():
                        student_name = str(row[name_column]).strip()
                        
                        # Calculate points based on check-in time
                        points = 0.6  # Default to on-time
                        status = "On Time"
                        
                        # Check timestamp if available - get settings from sidebar session state or use defaults
                        # Note: These variables need to be accessible, so we'll use session state
                        current_class_start = st.session_state.get('class_start_time', datetime.now().replace(hour=9, minute=0, second=0, microsecond=0).time())
                        current_late_threshold = st.session_state.get('late_threshold_minutes', 15)
                        
                        if 'Timestamp' in row and pd.notna(row['Timestamp']):
                            try:
                                checkin_datetime = pd.to_datetime(row['Timestamp'])
                                class_start_datetime = datetime.combine(datetime.today(), current_class_start)
                                minutes_late = (checkin_datetime - class_start_datetime).total_seconds() / 60
                                
                                if minutes_late > current_late_threshold:
                                    points = 0.2
                                    status = "Late"
                            except:
                                pass
                        elif 'Time' in row and pd.notna(row['Time']):
                            # Try parsing time string
                            try:
                                checkin_time = datetime.strptime(str(row['Time']), "%H:%M:%S").time()
                                checkin_datetime = datetime.combine(datetime.today(), checkin_time)
                                class_start_datetime = datetime.combine(datetime.today(), current_class_start)
                                minutes_late = (checkin_datetime - class_start_datetime).total_seconds() / 60
                                
                                if minutes_late > current_late_threshold:
                                    points = 0.2
                                    status = "Late"
                            except:
                                pass
                        
                        # Check if already processed (by checking if student already has points for today)
                        roster_idx, confidence, matched_name = find_student_in_roster(
                            student_name, 
                            st.session_state.roster,
                            st.session_state.get('use_gemini', False),
                            st.session_state.get('min_confidence', 0.75)
                        )
                        if roster_idx is not None:
                            if date_str in st.session_state.roster.columns:
                                existing_points = st.session_state.roster.loc[roster_idx, date_str]
                                if pd.notna(existing_points) and existing_points > 0:
                                    # Keep higher points if already processed
                                    if existing_points >= points:
                                        continue
                        
                        # Update roster
                        use_gemini_flag = st.session_state.get('use_gemini', False)
                        min_conf = st.session_state.get('min_confidence', 0.75)
                        updated_roster, found, confidence, matched_name = update_roster_with_attendance(
                            st.session_state.roster,
                            student_name,
                            points,
                            date_str,
                            use_gemini_flag,
                            min_conf
                        )
                        
                        if found:
                            st.session_state.roster = updated_roster
                            processed_count += 1
                            
                            # Track low confidence matches for review
                            if confidence < 1.0:
                                low_confidence_matches.append({
                                    'Check-in Name': student_name,
                                    'Matched Roster Name': matched_name,
                                    'Confidence': f"{confidence:.2%}",
                                    'Points': points,
                                    'Status': status
                                })
                            
                            st.session_state.attendance_records.append({
                                'Name': student_name,
                                'Date': date_str,
                                'Points': points,
                                'Status': status,
                                'Confidence': f"{confidence:.2%}"
                            })
                        else:
                            errors.append({
                                'Name': student_name,
                                'Best Match Confidence': f"{confidence:.2%}",
                                'Suggested Match': matched_name if matched_name else "None"
                            })
                
                    if processed_count > 0:
                        # Save roster to persistent file
                        if save_roster_to_file(st.session_state.roster, st.session_state.roster_file_path):
                            st.success(f"✅ Processed {processed_count} check-ins and updated roster! File automatically saved.")
                        else:
                            st.success(f"✅ Processed {processed_count} check-ins and updated roster!")
                        
                        # Show low confidence matches for review
                        if low_confidence_matches:
                            st.subheader("⚠️ Low Confidence Matches - Please Review")
                            st.dataframe(pd.DataFrame(low_confidence_matches), use_container_width=True)
                            st.info("💡 These matches had confidence below 100%. Please verify they are correct. You can adjust the confidence threshold in the sidebar if needed.")
                        
                        # Save processed check-ins to file for record keeping
                        if checkin_import_file is not None:
                            # Save a copy of the imported file
                            processed_file = f"checkins_{date_str}.csv"
                            checkins_df.to_csv(processed_file, index=False)
                            st.info(f"💾 Check-in data saved to {processed_file}")
                    
                    if errors:
                        st.subheader("⚠️ Students Not Found in Roster")
                        errors_df = pd.DataFrame(errors)
                        st.dataframe(errors_df, use_container_width=True)
                        st.info("💡 Check the suggested matches. You can lower the confidence threshold or use Gemini API to improve matching.")
            
            if checkins_df is not None and not checkins_df.empty:
                st.dataframe(checkins_df, use_container_width=True)
        
        st.divider()
        
        # Manual entry option
        st.subheader("4. Manual Check-In Entry (Alternative)")
        st.write("You can also manually enter individual student check-ins:")
        st.info("💡 **Flexible Format:** Enter name in any format - 'Last, First', 'First, Last', or 'First Last'. "
                "The app will automatically match to your roster, even with minor typos!")
        
        student_name = st.text_input(
            "Enter Student Name",
            placeholder="Any format: 'Last, First', 'First, Last', or 'First Last'",
            help="Accepts any name format. The app will automatically match to your roster."
        )
        
        if st.button("Add Check-In", type="primary"):
                if student_name.strip():
                    # Calculate points based on time
                    current_time = datetime.now().time()
                    class_start = class_start_time
                    
                    # Check if late
                    current_datetime = datetime.combine(datetime.today(), current_time)
                    start_datetime = datetime.combine(datetime.today(), class_start)
                    
                    minutes_late = (current_datetime - start_datetime).total_seconds() / 60
                    
                    if minutes_late <= late_threshold_minutes:
                        points = 0.6
                        status = "On Time"
                    else:
                        points = 0.2
                        status = "Late"
                    
                    # Update session state with current settings
                    st.session_state.class_start_time = class_start_time
                    st.session_state.late_threshold_minutes = late_threshold_minutes
                    
                    # Update roster (use improved matching with confidence scoring)
                    date_str = datetime.now().strftime("%Y-%m-%d")
                    use_gemini_flag = st.session_state.get('use_gemini', False)
                    min_conf = st.session_state.get('min_confidence', 0.75)
                    updated_roster, found, confidence, matched_name = update_roster_with_attendance(
                        st.session_state.roster, 
                        student_name, 
                        points, 
                        date_str,
                        use_gemini_flag,
                        min_conf
                    )
                    
                    if found:
                        st.session_state.roster = updated_roster
                        
                        # Save roster to persistent file
                        if save_roster_to_file(st.session_state.roster, st.session_state.roster_file_path):
                            if confidence < 1.0:
                                st.success(f"✅ {student_name} checked in ({status}) - {points} points. Matched to '{matched_name}' (confidence: {confidence:.1%}). Roster automatically saved.")
                            else:
                                st.success(f"✅ {student_name} checked in ({status}) - {points} points. Roster automatically saved.")
                        else:
                            if confidence < 1.0:
                                st.success(f"✅ {student_name} checked in ({status}) - {points} points. Matched to '{matched_name}' (confidence: {confidence:.1%}).")
                            else:
                                st.success(f"✅ {student_name} checked in ({status}) - {points} points")
                        
                        st.session_state.attendance_records.append({
                            'Name': student_name,
                            'Matched Name': matched_name if confidence < 1.0 else student_name,
                            'Date': date_str,
                            'Points': points,
                            'Status': status,
                            'Time': current_time.strftime("%H:%M:%S"),
                            'Confidence': f"{confidence:.1%}"
                        })
                    else:
                        # Show helpful suggestions
                        if matched_name:
                            st.error(f"❌ Student '{student_name}' not found in roster. Best match: '{matched_name}' (confidence: {confidence:.1%})")
                            st.info(f"💡 Tip: The student might have entered their name differently. Try: '{matched_name}' or adjust the confidence threshold in the sidebar.")
                        else:
                            st.error(f"❌ Student '{student_name}' not found in roster. Please check the name format or spelling.")
                            st.info(f"💡 Tip: Check for typos or try different name formats. You can also lower the confidence threshold or enable Gemini API in the sidebar for better matching.")
                else:
                    st.warning("Please enter a student name.")

with tab2:
    st.header("Zoom Meeting Attendance Processing")
    
    if st.session_state.roster is None:
        st.warning("⚠️ Please upload a roster file in the sidebar first.")
    else:
        st.subheader("Upload Zoom Attendance Report")
        
        zoom_file = st.file_uploader(
            "Upload Zoom Attendance Excel File",
            type=['xlsx', 'xls', 'csv'],
            help="Upload the Zoom meeting report with participant names and durations"
        )
        
        if zoom_file is not None:
            try:
                if zoom_file.name.endswith('.csv'):
                    zoom_df = pd.read_csv(zoom_file)
                else:
                    zoom_df = pd.read_excel(zoom_file)
                
                st.success(f"Zoom file loaded: {len(zoom_df)} participants")
                st.dataframe(zoom_df.head(10))
                
                # Date selection
                meeting_date = st.date_input(
                    "Select Meeting Date",
                    value=datetime.now().date()
                )
                
                if st.button("Process Zoom Attendance", type="primary"):
                    date_str = meeting_date.strftime("%Y-%m-%d")
                    st.session_state.roster = process_zoom_attendance(
                        st.session_state.roster,
                        zoom_df,
                        date_str
                    )
                    
                    # Save roster to persistent file
                    if save_roster_to_file(st.session_state.roster, st.session_state.roster_file_path):
                        st.success("✅ Zoom attendance processed successfully! Roster automatically saved.")
                    else:
                        st.success("Zoom attendance processed successfully!")
                    
            except Exception as e:
                st.error(f"Error processing Zoom file: {str(e)}")
                st.exception(e)

with tab3:
    st.header("View and Export Roster")
    
    if st.session_state.roster is None:
        st.warning("⚠️ Please upload a roster file in the sidebar first.")
    else:
        st.subheader("Current Roster with Attendance")
        st.dataframe(st.session_state.roster, use_container_width=True)
        
        # Show date columns (attendance dates)
        date_columns = [col for col in st.session_state.roster.columns 
                       if re.match(r'\d{4}-\d{2}-\d{2}', str(col))]
        
        if date_columns:
            st.info(f"📅 **Attendance Dates in Roster:** {len(date_columns)} meetings recorded\n\n**Dates:** {', '.join(sorted(date_columns))}")
            
            roster_with_total = st.session_state.roster.copy()
            roster_with_total['Total Points'] = roster_with_total[date_columns].sum(axis=1)
            st.subheader("Roster with Total Points")
            st.dataframe(roster_with_total, use_container_width=True)
            
            # Download button (backup copy)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                roster_with_total.to_excel(writer, index=False, sheet_name='Attendance')
            
            st.download_button(
                label="📥 Download Backup Copy (Excel)",
                data=output.getvalue(),
                file_name=f"attendance_roster_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Download a backup copy. The main roster is automatically saved to: " + st.session_state.roster_file_path
            )
            
            st.info(f"💡 **Note:** The roster is automatically saved to `{st.session_state.roster_file_path}` after each update. No need to download unless you want a backup copy.")
        
        # Show recent check-ins
        if st.session_state.attendance_records:
            st.subheader("Recent Check-Ins")
            recent_df = pd.DataFrame(st.session_state.attendance_records)
            st.dataframe(recent_df.tail(20), use_container_width=True)
