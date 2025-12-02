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
        
        try:
            lines = script_content.split('\n')
            
            for line_num, line in enumerate(lines, 1):
                command, args = self.parse_line(line)
                
                if command is None:
                    continue
                
                try:
                    result = self.execute_command(command, args)
                    if result:
                        self.output.append({
                            'line': line_num,
                            'command': command,
                            'result': result,
                            'success': True
                        })
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
        
        file_path = args[0].strip('"\'') if args else None
        
        if 'save_roster' in self.app_funcs:
            self.app_funcs['save_roster'](self.context['roster'])
            return f"Roster saved successfully"
        
        if file_path:
            self.context['roster'].to_excel(file_path, index=False, engine='openpyxl')
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
        if 'find_matching_date_column' in self.app_funcs:
            try:
                # Parse date string
                meeting_date = datetime.fromisoformat(date_str) if '-' in date_str else datetime.strptime(date_str, '%m/%d/%Y')
                date_col = self.app_funcs['find_matching_date_column'](df, meeting_date)
            except:
                pass
        
        # If no matching column found, try direct column name match
        if not date_col:
            # Try various date formats
            possible_names = [
                date_str,
                date_str.replace('-', '.').replace('/', '.'),
            ]
            for col in df.columns:
                col_str = str(col)
                for possible in possible_names:
                    if possible.lower() in col_str.lower() or col_str.lower() in possible.lower():
                        date_col = col
                        break
                if date_col:
                    break
        
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
        
        # Format the result
        result_lines = [f"Late students for {date_str} ({len(late_students)} students):"]
        for i, name in enumerate(late_students, 1):
            result_lines.append(f"  {i}. {name}")
        
        # Also add to output for display
        result_text = '\n'.join(result_lines)
        self.output.append({
            'line': 0,
            'command': 'SHOW LATE STUDENTS',
            'result': result_text,
            'success': True
        })
        
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
        
        # Try to find the matching date column using app helper if available
        date_col = None
        if 'find_matching_date_column' in self.app_funcs:
            try:
                # Parse date string if it's in a date format
                meeting_date = datetime.fromisoformat(date_str) if '-' in date_str else datetime.strptime(date_str, '%m/%d/%Y')
                date_col = self.app_funcs['find_matching_date_column'](df, meeting_date)
            except:
                pass
        
        # If no matching column found, try direct column name match
        if not date_col:
            possible_names = [
                date_str,
                date_str.replace('-', '.').replace('/', '.'),
            ]
            for col in df.columns:
                col_str = str(col)
                for possible in possible_names:
                    if possible.lower() in col_str.lower() or col_str.lower() in possible.lower():
                        date_col = col
                        break
                if date_col:
                    break
        
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
        
        # Format the result
        result_lines = [f"Early/on-time students for {date_str} ({len(early_students)} students):"]
        for i, name in enumerate(early_students, 1):
            result_lines.append(f"  {i}. {name}")
        
        result_text = '\n'.join(result_lines)
        
        # Also add to output for display
        self.output.append({
            'line': 0,
            'command': 'SHOW EARLY STUDENTS',
            'result': result_text,
            'success': True
        })
        
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

