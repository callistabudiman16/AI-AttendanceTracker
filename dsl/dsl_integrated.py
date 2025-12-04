"""
Integrated DSL Executor for Attendance Tracker Flask App

This module provides DSL execution that integrates directly with Flask app functions.
"""

import os
import re
import time
import pandas as pd
from datetime import datetime
from typing import Dict, List, Tuple, Optional, Any
import qrcode
from PIL import Image
# Flask session import - optional for non-Flask usage
try:
    from flask import session as flask_session
    FLASK_AVAILABLE = True
except ImportError:
    FLASK_AVAILABLE = False
    flask_session = None


class IntegratedDSLExecutor:
    """
    DSL Executor that integrates with Flask app functions.
    This version can access the app's session and functions directly.
    """
    
    def __init__(self, app_functions: Dict[str, Any], session_obj=None):
        """
        Initialize the integrated DSL executor.
        
        Args:
            app_functions: Dictionary mapping function names to callable functions
            session_obj: Flask session object for state management
        """
        self.app_funcs = app_functions
        self.session = session_obj
        self.context = {
            'roster': None,
            'settings': {
                'early_bird_start_time': '11:00',
                'regular_start_time': '11:36',
                'use_gemini': False,
                'gemini_api_key': ''
            }
        }
        self.output = []
        
        # Map DSL commands to execution methods
        self.commands = {
            'LOAD ROSTER': self._cmd_load_roster,
            'SAVE ROSTER': self._cmd_save_roster,
            'DOWNLOAD ROSTER': self._cmd_download_roster,
            'PROCESS CHECKIN': self._cmd_process_checkin,
            'SET CHECKIN TIMES': self._cmd_set_checkin_times,
            'PROCESS ZOOM': self._cmd_process_zoom,
            'VIEW ROSTER': self._cmd_view_roster,
            'DELETE DATE': self._cmd_delete_date,
            'ENABLE GEMINI': self._cmd_enable_gemini,
            'DISABLE GEMINI': self._cmd_disable_gemini,
            'SET GEMINI KEY': self._cmd_set_gemini_key,
            'GENERATE QR': self._cmd_generate_qr,
            'ECHO': self._cmd_echo,
            'WAIT': self._cmd_wait,
            'SHOW LATE STUDENTS': self._cmd_show_late_students,
            'SHOW EARLY STUDENTS': self._cmd_show_early_students,
            'SHOW STUDENT TOTAL': self._cmd_show_student_total,
            'FIND STUDENT': self._cmd_show_student_total,  # Alias
        }
    
    def parse_line(self, line: str) -> Tuple[Optional[str], List[str]]:
        """Parse a single line of DSL code."""
        # Remove comments
        if '#' in line:
            line = line[:line.index('#')]
        
        line = line.strip()
        if not line:
            return None, []
        
        # Parse quoted strings and tokens
        tokens = []
        current_token = ""
        in_quote = False
        quote_char = None
        i = 0
        
        while i < len(line):
            char = line[i]
            
            if not in_quote:
                if char in ['"', "'"]:
                    in_quote = True
                    quote_char = char
                elif char.isspace():
                    if current_token:
                        tokens.append(current_token)
                        current_token = ""
                else:
                    current_token += char
            else:
                if char == quote_char:
                    if i + 1 < len(line) and line[i + 1] == quote_char:
                        current_token += quote_char
                        i += 1
                    else:
                        in_quote = False
                        quote_char = None
                        tokens.append(current_token)
                        current_token = ""
                else:
                    current_token += char
            i += 1
        
        if current_token:
            tokens.append(current_token)
        
        if not tokens:
            return None, []
        
        # Match command (case-insensitive)
        command = None
        args_start = 0
        
        for cmd_name in sorted(self.commands.keys(), key=lambda x: len(x.split()), reverse=True):
            cmd_words = cmd_name.split()
            if len(tokens) >= len(cmd_words):
                matched = ' '.join(tokens[:len(cmd_words)]).upper()
                if matched == cmd_name:
                    command = cmd_name
                    args_start = len(cmd_words)
                    break
        
        args = tokens[args_start:] if command else tokens
        return command, args
    
    def execute_script(self, script_content: str) -> Dict[str, Any]:
        """
        Execute DSL script from content string.
        
        Args:
            script_content: Content of the DSL script
            
        Returns:
            Dictionary with execution results
        """
        self.output = []
        # Initialize structured data attributes for commands that use them
        self._last_student_list = None
        self._last_header = None
        
        try:
            lines = script_content.split('\n')
            
            for line_num, line in enumerate(lines, 1):
                command, args = self.parse_line(line)
                
                if command is None:
                    continue
                
                try:
                    result = self.execute_command(command, args)
                    if result:
                        # Check if command stored structured student list data
                        output_item = {
                            'line': line_num,
                            'command': command,
                            'result': result,
                            'success': True
                        }
                        # Add structured data if available (from SHOW LATE/EARLY STUDENTS commands)
                        if hasattr(self, '_last_student_list') and self._last_student_list is not None:
                            output_item['student_list'] = self._last_student_list
                            output_item['header'] = getattr(self, '_last_header', None)
                            # Clear the stored data after using it
                            self._last_student_list = None
                            self._last_header = None
                        self.output.append(output_item)
                except Exception as e:
                    return {
                        'success': False,
                        'error': f"Error at line {line_num}: {str(e)}",
                        'output': self.output,
                        'line': line.strip(),
                        'line_num': line_num
                    }
            
            return {
                'success': True,
                'output': self.output,
                'message': f"Script executed successfully: {len(self.output)} commands processed"
            }
        
        except Exception as e:
            return {
                'success': False,
                'error': f"Error executing script: {str(e)}",
                'output': self.output
            }
    
    def execute_command(self, command: str, args: List[str]) -> Optional[str]:
        """Execute a single DSL command."""
        if command not in self.commands:
            raise ValueError(f"Unknown command: {command}")
        
        return self.commands[command](args)
    
    def _cmd_load_roster(self, args: List[str]) -> str:
        """LOAD ROSTER command"""
        if not args:
            raise ValueError("LOAD ROSTER requires a file path")
        
        file_path = args[0].strip('"\'')
        
        # Use app's load_roster function if available
        if 'load_roster' in self.app_funcs:
            roster_df = self.app_funcs['load_roster']()
            if roster_df is not None:
                self.context['roster'] = roster_df
                return f"Roster loaded: {len(roster_df)} students"
        
        # Fallback to direct file reading
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Roster file not found: {file_path}")
        
        if file_path.lower().endswith('.csv'):
            roster_df = pd.read_csv(file_path, encoding='utf-8')
        else:
            roster_df = pd.read_excel(file_path, engine='openpyxl')
        
        self.context['roster'] = roster_df
        
        # Save to app's default location
        if 'save_roster' in self.app_funcs:
            self.app_funcs['save_roster'](roster_df)
        
        return f"Roster loaded: {len(roster_df)} students from {file_path}"
    
    def _cmd_save_roster(self, args: List[str]) -> str:
        """SAVE ROSTER command"""
        if self.context['roster'] is None and 'load_roster' in self.app_funcs:
            self.context['roster'] = self.app_funcs['load_roster']()
        
        if self.context['roster'] is None:
            raise ValueError("No roster loaded. Use LOAD ROSTER first.")
        
        # Calculate Total Points before saving
        df = self.context['roster'].copy()
        date_columns = []
        non_date_columns = ['Unnamed: 0', 'No.', 'ID', 'Name', 'Major', 'Level', 'Total Points']
        
        import re
        for col in df.columns:
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
            numeric_date_cols = [col for col in date_columns if df[col].dtype in ['int64', 'float64']]
            if numeric_date_cols:
                df['Total Points'] = df[numeric_date_cols].fillna(0).sum(axis=1)
                self.context['roster'] = df
        
        file_path = args[0].strip('"\'') if args else None
        
        if 'save_roster' in self.app_funcs:
            self.app_funcs['save_roster'](df)
            return f"Roster saved successfully"
        
        if file_path:
            df.to_excel(file_path, index=False, engine='openpyxl')
            return f"Roster saved to {file_path}"
        else:
            raise ValueError("No file path specified and save_roster function not available")
    
    def _cmd_download_roster(self, args: List[str]) -> str:
        """DOWNLOAD ROSTER command"""
        return self._cmd_save_roster(args)
    
    def _cmd_process_checkin(self, args: List[str]) -> str:
        """PROCESS CHECKIN command"""
        # Parse arguments
        file_path = None
        date = None
        early_bird = self.context['settings']['early_bird_start_time']
        regular = self.context['settings']['regular_start_time']
        
        i = 0
        while i < len(args):
            arg = args[i].upper()
            if arg == 'DATE' and i + 1 < len(args):
                date = args[i + 1].strip('"\'')
                i += 2
            elif arg == 'EARLY_BIRD' and i + 1 < len(args):
                early_bird = args[i + 1].strip('"\'')
                i += 2
            elif arg == 'REGULAR' and i + 1 < len(args):
                regular = args[i + 1].strip('"\'')
                i += 2
            else:
                if file_path is None:
                    file_path = args[i].strip('"\'')
                i += 1
        
        if not file_path:
            raise ValueError("PROCESS CHECKIN requires a file path")
        
        # Update session settings
        if self.session:
            self.session['early_bird_start_time'] = early_bird
            self.session['regular_start_time'] = regular
        
        return f"Check-in processing initiated for {file_path} (date: {date or 'auto-detect'}, early_bird: {early_bird}, regular: {regular})"
    
    def _cmd_set_checkin_times(self, args: List[str]) -> str:
        """SET CHECKIN TIMES command"""
        early_bird = None
        regular = None
        
        i = 0
        while i < len(args):
            arg = args[i].upper()
            if arg == 'EARLY_BIRD' and i + 1 < len(args):
                early_bird = args[i + 1].strip('"\'')
                i += 2
            elif arg == 'REGULAR' and i + 1 < len(args):
                regular = args[i + 1].strip('"\'')
                i += 2
            else:
                i += 1
        
        if early_bird:
            self.context['settings']['early_bird_start_time'] = early_bird
            if self.session:
                self.session['early_bird_start_time'] = early_bird
        if regular:
            self.context['settings']['regular_start_time'] = regular
            if self.session:
                self.session['regular_start_time'] = regular
        
        return f"Check-in times set: Early Bird={early_bird or self.context['settings']['early_bird_start_time']}, Regular={regular or self.context['settings']['regular_start_time']}"
    
    def _cmd_process_zoom(self, args: List[str]) -> str:
        """PROCESS ZOOM command"""
        file_path = None
        date = None
        
        i = 0
        while i < len(args):
            arg = args[i].upper()
            if arg == 'DATE' and i + 1 < len(args):
                date = args[i + 1].strip('"\'')
                i += 2
            else:
                if file_path is None:
                    file_path = args[i].strip('"\'')
                i += 1
        
        if not file_path:
            raise ValueError("PROCESS ZOOM requires a file path")
        
        return f"Zoom processing initiated for {file_path} (date: {date or 'auto-detect'})"
    
    def _cmd_view_roster(self, args: List[str]) -> str:
        """VIEW ROSTER command"""
        if self.context['roster'] is None and 'load_roster' in self.app_funcs:
            self.context['roster'] = self.app_funcs['load_roster']()
        
        if self.context['roster'] is None:
            raise ValueError("No roster loaded. Use LOAD ROSTER first.")
        
        df = self.context['roster']
        return f"Roster: {len(df)} students, {len(df.columns)} columns"
    
    def _cmd_delete_date(self, args: List[str]) -> str:
        """DELETE DATE command"""
        if not args:
            raise ValueError("DELETE DATE requires a date column name")
        
        if self.context['roster'] is None and 'load_roster' in self.app_funcs:
            self.context['roster'] = self.app_funcs['load_roster']()
        
        if self.context['roster'] is None:
            raise ValueError("No roster loaded. Use LOAD ROSTER first.")
        
        date_col = args[0].strip('"\'')
        df = self.context['roster']
        
        if date_col not in df.columns:
            raise ValueError(f"Date column '{date_col}' not found in roster")
        
        df.drop(columns=[date_col], inplace=True)
        
        # Save updated roster
        if 'save_roster' in self.app_funcs:
            self.app_funcs['save_roster'](df)
        
        return f"Deleted date column: {date_col}"
    
    def _cmd_enable_gemini(self, args: List[str]) -> str:
        """ENABLE GEMINI command"""
        self.context['settings']['use_gemini'] = True
        if self.session:
            self.session['use_gemini'] = True
        return "Gemini AI matching enabled"
    
    def _cmd_disable_gemini(self, args: List[str]) -> str:
        """DISABLE GEMINI command"""
        self.context['settings']['use_gemini'] = False
        if self.session:
            self.session['use_gemini'] = False
        return "Gemini AI matching disabled"
    
    def _cmd_set_gemini_key(self, args: List[str]) -> str:
        """SET GEMINI KEY command"""
        if not args:
            raise ValueError("SET GEMINI KEY requires an API key")
        
        api_key = args[0].strip('"\'')
        self.context['settings']['gemini_api_key'] = api_key
        if self.session:
            self.session['gemini_api_key'] = api_key
        return "Gemini API key set"
    
    def _cmd_generate_qr(self, args: List[str]) -> str:
        """GENERATE QR command"""
        if not args:
            raise ValueError("GENERATE QR requires a URL")
        
        url = args[0].strip('"\'')
        output_path = None
        
        i = 1
        while i < len(args):
            if args[i].upper() == 'OUTPUT' and i + 1 < len(args):
                output_path = args[i + 1].strip('"\'')
                i += 2
            else:
                i += 1
        
        qr = qrcode.QRCode(version=1, box_size=10, border=5)
        qr.add_data(url)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")
        
        if output_path:
            img.save(output_path)
            return f"QR code generated and saved to {output_path}"
        else:
            return f"QR code generated for URL: {url}"
    
    def _cmd_echo(self, args: List[str]) -> str:
        """ECHO command"""
        message = ' '.join(args).strip('"\'')
        return f"ECHO: {message}"
    
    def _cmd_show_late_students(self, args: List[str]) -> str:
        """SHOW LATE STUDENTS command - shows students who got 0.2 points (late) for a given date"""
        if self.context['roster'] is None and 'load_roster' in self.app_funcs:
            self.context['roster'] = self.app_funcs['load_roster']()
        
        if self.context['roster'] is None:
            raise ValueError("No roster loaded. Use LOAD ROSTER first.")
        
        # Parse arguments to find DATE
        date_str = None
        i = 0
        while i < len(args):
            arg = args[i].upper()
            if arg == 'DATE' and i + 1 < len(args):
                date_str = args[i + 1].strip('"\'')
                i += 2
            else:
                i += 1
        
        if not date_str:
            raise ValueError("SHOW LATE STUDENTS requires DATE parameter")
        
        df = self.context['roster']
        
        # Try to find the matching date column
        date_col = None
        
        # First, try exact column name match (handles formats like "T,Nov.4", "R,Oct.23", etc.)
        if date_str in df.columns:
            date_col = date_str
        else:
            # Try case-insensitive exact match
            for col in df.columns:
                if str(col).strip().lower() == date_str.strip().lower():
                    date_col = col
                    break
        
        # If not found, try using find_matching_date_column function
        if not date_col and 'find_matching_date_column' in self.app_funcs:
            try:
                # Try to parse date string in various formats
                meeting_date = None
                # Try ISO format first (YYYY-MM-DD)
                if '-' in date_str and len(date_str.split('-')) == 3:
                    try:
                        meeting_date = datetime.strptime(date_str, '%Y-%m-%d')
                    except:
                        pass
                # Try MM/DD/YYYY format
                if not meeting_date and '/' in date_str:
                    try:
                        meeting_date = datetime.strptime(date_str, '%m/%d/%Y')
                    except:
                        pass
                # Try MM.DD format (e.g., "11.4")
                if not meeting_date and '.' in date_str and len(date_str.split('.')) == 2:
                    try:
                        parts = date_str.split('.')
                        month, day = int(parts[0]), int(parts[1])
                        meeting_date = datetime(2024, month, day)  # Use 2024 as default year
                    except:
                        pass
                
                if meeting_date:
                    date_col = self.app_funcs['find_matching_date_column'](df, meeting_date)
            except Exception as e:
                pass
        
        # If still not found, try pattern matching for date-like column names
        # This handles formats like "T,Nov.4", "R,Oct.23", "Nov.4", etc.
        if not date_col:
            date_str_lower = date_str.lower().strip()
            import re
            
            # Extract month and day from date string if possible
            # Patterns: "T,Nov.4", "R,Oct.23", "Nov.4", "11.4", etc.
            month_day_pattern = re.search(r'(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\.(\d{1,2})', date_str_lower)
            if month_day_pattern:
                # Found month.day pattern (e.g., "nov.4")
                month_day_str = month_day_pattern.group(0)  # e.g., "nov.4"
                # Also check if there's a prefix like "T," or "R,"
                prefix_pattern = re.match(r'^([A-Z]),', date_str)
                prefix = prefix_pattern.group(1) if prefix_pattern else None
                
                # Match columns that contain this exact month.day pattern
                for col in df.columns:
                    col_str = str(col).lower().strip()
                    # Check if column contains the same month.day pattern
                    if month_day_str in col_str:
                        # If there's a prefix in the date string, prefer columns with the same prefix
                        if prefix:
                            col_prefix_match = re.match(r'^([a-z]),', col_str)
                            if col_prefix_match and col_prefix_match.group(1).upper() == prefix:
                                date_col = col
                                break
                        # If no prefix or prefix doesn't match, still accept if month.day matches
                        if not date_col:
                            date_col = col
            else:
                # Try numeric format like "11.4"
                numeric_pattern = re.match(r'^(\d{1,2})\.(\d{1,2})$', date_str)
                if numeric_pattern:
                    month, day = numeric_pattern.groups()
                    # Check for prefix
                    prefix_pattern = re.match(r'^([A-Z]),', date_str)
                    prefix = prefix_pattern.group(1) if prefix_pattern else None
                    
                    for col in df.columns:
                        col_str = str(col).lower().strip()
                        # Match columns like "11.4" or "T,11.4"
                        if f"{month}.{day}" in col_str:
                            # If there's a prefix, prefer columns with the same prefix
                            if prefix:
                                col_prefix_match = re.match(r'^([a-z]),', col_str)
                                if col_prefix_match and col_prefix_match.group(1).upper() == prefix:
                                    date_col = col
                                    break
                            # If no prefix or prefix doesn't match, still accept if numeric pattern matches
                            if not date_col:
                                date_col = col
        
        if not date_col or date_col not in df.columns:
            raise ValueError(f"Date column not found for date: {date_str}")
        
        # Find name column
        name_col = None
        for col in df.columns:
            col_str = str(col).lower().strip()
            if ('name' in col_str and 'unnamed' not in col_str and 
                col_str not in ['id', 'email', 'major', 'level']):
                name_col = col
                break
        if name_col is None:
            if len(df.columns) > 2:
                name_col = df.columns[2]
            else:
                name_col = df.columns[0] if len(df.columns) > 0 else None
        
        if name_col is None:
            raise ValueError("Could not find name column in roster")
        
        # Filter students with exactly 0.2 points
        late_students = []
        for idx, row in df.iterrows():
            points_value = row[date_col]
            # Check if points value is exactly 0.2
            if pd.notna(points_value):
                try:
                    points = float(points_value)
                    if abs(points - 0.2) < 0.01:  # Using small epsilon for float comparison
                        student_name = str(row[name_col]).strip()
                        if student_name and student_name.lower() not in ['nan', 'none', '']:
                            late_students.append(student_name)
                except (ValueError, TypeError):
                    pass
        
        if not late_students:
            return f"No late students found for date {date_str} (date column: {date_col})"
        
        # Format the result with structured data for better display
        result_header = f"Late students for {date_str} ({len(late_students)} students):"
        result_text = result_header + '\n' + '\n'.join([f"  {i}. {name}" for i, name in enumerate(late_students, 1)])
        
        # Store structured data for later use in execute_script
        # The execute_script will add this to output, so we store it as an attribute
        self._last_student_list = late_students
        self._last_header = result_header
        
        return result_text
    
    def _cmd_show_early_students(self, args: List[str]) -> str:
        """SHOW EARLY STUDENTS command - shows students who got 0.6 points (on-time/early) for a given date"""
        # Ensure roster is loaded
        if self.context['roster'] is None and 'load_roster' in self.app_funcs:
            self.context['roster'] = self.app_funcs['load_roster']()
        
        if self.context['roster'] is None:
            raise ValueError("No roster loaded. Use LOAD ROSTER first.")
        
        # Parse arguments to find DATE
        date_str = None
        i = 0
        while i < len(args):
            arg = args[i].upper()
            if arg == 'DATE' and i + 1 < len(args):
                date_str = args[i + 1].strip('"\'')
                i += 2
            else:
                i += 1
        
        if not date_str:
            raise ValueError("SHOW EARLY STUDENTS requires DATE parameter")
        
        df = self.context['roster']
        
        # Try to find the matching date column
        date_col = None
        
        # First, try exact column name match (handles formats like "T,Nov.4", "R,Oct.23", etc.)
        if date_str in df.columns:
            date_col = date_str
        else:
            # Try case-insensitive exact match
            for col in df.columns:
                if str(col).strip().lower() == date_str.strip().lower():
                    date_col = col
                    break
        
        # If not found, try using find_matching_date_column function
        if not date_col and 'find_matching_date_column' in self.app_funcs:
            try:
                # Try to parse date string in various formats
                meeting_date = None
                # Try ISO format first (YYYY-MM-DD)
                if '-' in date_str and len(date_str.split('-')) == 3:
                    try:
                        meeting_date = datetime.strptime(date_str, '%Y-%m-%d')
                    except:
                        pass
                # Try MM/DD/YYYY format
                if not meeting_date and '/' in date_str:
                    try:
                        meeting_date = datetime.strptime(date_str, '%m/%d/%Y')
                    except:
                        pass
                # Try MM.DD format (e.g., "11.4")
                if not meeting_date and '.' in date_str and len(date_str.split('.')) == 2:
                    try:
                        parts = date_str.split('.')
                        month, day = int(parts[0]), int(parts[1])
                        meeting_date = datetime(2024, month, day)  # Use 2024 as default year
                    except:
                        pass
                
                if meeting_date:
                    date_col = self.app_funcs['find_matching_date_column'](df, meeting_date)
            except Exception as e:
                pass
        
        # If still not found, try pattern matching for date-like column names
        # This handles formats like "T,Nov.4", "R,Oct.23", "Nov.4", etc.
        if not date_col:
            date_str_lower = date_str.lower().strip()
            import re
            
            # Extract month and day from date string if possible
            # Patterns: "T,Nov.4", "R,Oct.23", "Nov.4", "11.4", etc.
            month_day_pattern = re.search(r'(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\.(\d{1,2})', date_str_lower)
            if month_day_pattern:
                # Found month.day pattern (e.g., "nov.4")
                month_day_str = month_day_pattern.group(0)  # e.g., "nov.4"
                # Also check if there's a prefix like "T," or "R,"
                prefix_pattern = re.match(r'^([A-Z]),', date_str)
                prefix = prefix_pattern.group(1) if prefix_pattern else None
                
                # Match columns that contain this exact month.day pattern
                for col in df.columns:
                    col_str = str(col).lower().strip()
                    # Check if column contains the same month.day pattern
                    if month_day_str in col_str:
                        # If there's a prefix in the date string, prefer columns with the same prefix
                        if prefix:
                            col_prefix_match = re.match(r'^([a-z]),', col_str)
                            if col_prefix_match and col_prefix_match.group(1).upper() == prefix:
                                date_col = col
                                break
                        # If no prefix or prefix doesn't match, still accept if month.day matches
                        if not date_col:
                            date_col = col
            else:
                # Try numeric format like "11.4"
                numeric_pattern = re.match(r'^(\d{1,2})\.(\d{1,2})$', date_str)
                if numeric_pattern:
                    month, day = numeric_pattern.groups()
                    # Check for prefix
                    prefix_pattern = re.match(r'^([A-Z]),', date_str)
                    prefix = prefix_pattern.group(1) if prefix_pattern else None
                    
                    for col in df.columns:
                        col_str = str(col).lower().strip()
                        # Match columns like "11.4" or "T,11.4"
                        if f"{month}.{day}" in col_str:
                            # If there's a prefix, prefer columns with the same prefix
                            if prefix:
                                col_prefix_match = re.match(r'^([a-z]),', col_str)
                                if col_prefix_match and col_prefix_match.group(1).upper() == prefix:
                                    date_col = col
                                    break
                            # If no prefix or prefix doesn't match, still accept if numeric pattern matches
                            if not date_col:
                                date_col = col
        
        if not date_col or date_col not in df.columns:
            raise ValueError(f"Date column not found for date: {date_str}")
        
        # Find name column
        name_col = None
        for col in df.columns:
            col_str = str(col).lower().strip()
            if ('name' in col_str and 'unnamed' not in col_str and 
                col_str not in ['id', 'email', 'major', 'level']):
                name_col = col
                break
        if name_col is None:
            if len(df.columns) > 2:
                name_col = df.columns[2]
            else:
                name_col = df.columns[0] if len(df.columns) > 0 else None
        
        if name_col is None:
            raise ValueError("Could not find name column in roster")
        
        # Filter students with exactly 0.6 points (on-time/early)
        early_students = []
        for idx, row in df.iterrows():
            points_value = row[date_col]
            if pd.notna(points_value):
                try:
                    points = float(points_value)
                    if abs(points - 0.6) < 0.01:  # epsilon for float comparison
                        student_name = str(row[name_col]).strip()
                        if student_name and student_name.lower() not in ['nan', 'none', '']:
                            early_students.append(student_name)
                except (ValueError, TypeError):
                    pass
        
        if not early_students:
            return f"No early/on-time students found for date {date_str} (date column: {date_col})"
        
        # Format the result with structured data for better display
        result_header = f"Early/on-time students for {date_str} ({len(early_students)} students):"
        result_text = result_header + '\n' + '\n'.join([f"  {i}. {name}" for i, name in enumerate(early_students, 1)])
        
        # Store structured data for later use in execute_script
        # The execute_script will add this to output, so we store it as an attribute
        self._last_student_list = early_students
        self._last_header = result_header
        
        return result_text
    
    def _cmd_show_student_total(self, args: List[str]) -> str:
        """SHOW STUDENT TOTAL command - shows total points for a specific student"""
        if self.context['roster'] is None and 'load_roster' in self.app_funcs:
            self.context['roster'] = self.app_funcs['load_roster']()
        
        if self.context['roster'] is None:
            raise ValueError("No roster loaded. Use LOAD ROSTER first.")
        
        if not args:
            raise ValueError("SHOW STUDENT TOTAL requires a student name")
        
        # Get student name (may be quoted)
        student_name = ' '.join(args).strip('"\'')
        
        df = self.context['roster']
        
        # Find name column
        name_col = None
        for col in df.columns:
            col_str = str(col).lower().strip()
            if ('name' in col_str and 'unnamed' not in col_str and 
                col_str not in ['id', 'email', 'major', 'level']):
                name_col = col
                break
        if name_col is None:
            if len(df.columns) > 2:
                name_col = df.columns[2]
            else:
                name_col = df.columns[0] if len(df.columns) > 0 else None
        
        if name_col is None:
            raise ValueError("Could not find name column in roster")
        
        # Try to find student using app function if available
        student_idx = None
        matched_name = None
        
        if 'find_student_in_roster' in self.app_funcs:
            try:
                # Import find_student_in_roster from app if available
                from app import find_student_in_roster
                use_gemini = self.context.get('settings', {}).get('use_gemini', False)
                idx, confidence, matched = find_student_in_roster(student_name, df, use_gemini=use_gemini)
                if idx is not None:
                    student_idx = idx
                    matched_name = matched
            except ImportError:
                pass
        
        # If not found via app function, try simple matching
        if student_idx is None:
            # Try exact match first (case-insensitive)
            student_name_lower = student_name.lower().strip()
            for idx, row in df.iterrows():
                roster_name = str(row[name_col]).strip()
                if roster_name.lower() == student_name_lower:
                    student_idx = idx
                    matched_name = roster_name
                    break
            
            # Try partial match (contains)
            if student_idx is None:
                for idx, row in df.iterrows():
                    roster_name = str(row[name_col]).strip()
                    roster_name_lower = roster_name.lower()
                    # Check if student name is contained in roster name or vice versa
                    if (student_name_lower in roster_name_lower or 
                        roster_name_lower in student_name_lower):
                        student_idx = idx
                        matched_name = roster_name
                        break
        
        if student_idx is None:
            return f"Student '{student_name}' not found in roster"
        
        # Get total points
        total_points = 0.0
        if 'Total Points' in df.columns:
            total_val = df.loc[student_idx, 'Total Points']
            if pd.notna(total_val):
                try:
                    total_points = float(total_val)
                except (ValueError, TypeError):
                    total_points = 0.0
        
        # If Total Points column doesn't exist or is empty, calculate from date columns
        if total_points == 0.0 or 'Total Points' not in df.columns:
            # Find date columns and sum them
            date_columns = []
            non_date_columns = ['Unnamed: 0', 'No.', 'ID', 'Name', 'Major', 'Level', 'Total Points']
            
            for col in df.columns:
                col_str = str(col).strip()
                col_lower = col_str.lower()
                
                if col_str in non_date_columns or col_lower in [c.lower() for c in non_date_columns]:
                    continue
                
                # Match MM.DD format or Month.Day format
                import re
                if (re.match(r'^\d{1,2}\.\d{1,2}$', col_str) or
                    re.search(r'(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\.\d{1,2}', col_lower) or
                    re.match(r'^[A-Z],(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\.\d{1,2}', col_lower)):
                    date_columns.append(col)
            
            # Sum date columns
            for col in date_columns:
                val = df.loc[student_idx, col]
                if pd.notna(val):
                    try:
                        total_points += float(val)
                    except (ValueError, TypeError):
                        pass
        
        # Format result
        result_text = f"Student: {matched_name}\nTotal Points: {total_points:.1f}"
        
        return result_text
    
    def _cmd_wait(self, args: List[str]) -> str:
        """WAIT command"""
        if not args:
            raise ValueError("WAIT requires a number of seconds")
        
        try:
            seconds = float(args[0])
            time.sleep(seconds)
            return f"Waited {seconds} seconds"
        except ValueError:
            raise ValueError(f"Invalid wait time: {args[0]}")

