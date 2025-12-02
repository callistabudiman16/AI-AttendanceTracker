"""
Terminal-Based Attendance Tracker with Gemini API Integration

This program provides a terminal interface for managing attendance records
with AI-assisted DSL code generation using Google's Gemini API.
"""

import os
import sys
import pandas as pd
from datetime import datetime
from typing import Optional, Dict, Any, List
import json

# Import Gemini API
try:
    import google.generativeai as genai
    GEMINI_AVAILABLE = True
except ImportError:
    GEMINI_AVAILABLE = False
    print("Warning: google-generativeai not installed. Install with: pip install google-generativeai")

# Import DSL executor
try:
    from dsl.dsl_integrated import IntegratedDSLExecutor
    DSL_AVAILABLE = True
except ImportError:
    DSL_AVAILABLE = False
    print("Warning: DSL executor not available. Check dsl module.")

# Configuration
ROSTER_FILE = 'roster_attendance.xlsx'
ATTENDANCE_FOLDER = os.path.join(os.getcwd(), 'attendance record')
GEMINI_API_KEY_ENV = 'GEMINI_API_KEY'


class AttendanceTerminal:
    """Terminal-based attendance tracker with Gemini API integration"""
    
    def __init__(self):
        self.roster_df: Optional[pd.DataFrame] = None
        self.roster_file: Optional[str] = None
        self.gemini_model = None
        self.dsl_executor = None
        self.app_functions = {}
        
        # Initialize Gemini API if available
        if GEMINI_AVAILABLE:
            api_key = os.getenv(GEMINI_API_KEY_ENV)
            
            # Try to get from User environment variable if not in current session (Windows)
            if not api_key and os.name == 'nt':
                try:
                    import winreg
                    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment")
                    try:
                        api_key, _ = winreg.QueryValueEx(key, GEMINI_API_KEY_ENV)
                    except FileNotFoundError:
                        pass
                    finally:
                        winreg.CloseKey(key)
                except Exception:
                    pass
            
            if api_key:
                genai.configure(api_key=api_key)
                # Try available Gemini models in order of preference
                model_names = [
                    'gemini-2.0-flash',  # Latest stable flash model
                    'gemini-flash-latest',  # Latest flash (aliased)
                    'gemini-2.5-flash',  # Newer flash model
                    'gemini-2.0-flash-001',  # Specific version
                    'gemini-pro-latest',  # Latest pro (aliased)
                    'gemini-2.5-pro',  # Newer pro model
                ]
                
                self.gemini_model = None
                for model_name in model_names:
                    try:
                        self.gemini_model = genai.GenerativeModel(model_name)
                        print(f"✓ Using Gemini model: {model_name}")
                        break
                    except Exception:
                        continue
                
                if self.gemini_model is None:
                    print("Warning: Could not initialize any Gemini model.")
                    print("   Trying to list available models...")
                    try:
                        models = genai.list_models()
                        available = [m.name for m in models if 'generateContent' in m.supported_generation_methods]
                        print(f"   Found {len(available)} available models.")
                        if available:
                            print("   Trying first available model...")
                            try:
                                # Use just the model name without 'models/' prefix
                                first_model = available[0].replace('models/', '')
                                self.gemini_model = genai.GenerativeModel(first_model)
                                print(f"✓ Using model: {first_model}")
                            except Exception as e:
                                print(f"   Failed to use {available[0]}: {str(e)}")
                    except Exception as e:
                        print(f"   Error listing models: {str(e)}")
            else:
                print(f"Warning: {GEMINI_API_KEY_ENV} environment variable not set.")
                print(f"   Set it with: $env:{GEMINI_API_KEY_ENV}='your-api-key'")
        
        # Initialize DSL executor
        if DSL_AVAILABLE:
            # Create mock app functions for DSL executor
            self.app_functions = {
                'load_roster': self._load_roster_internal,
                'save_roster': self._save_roster_internal,
                'format_date_for_roster': self._format_date,
                'find_matching_date_column': self._find_matching_date_column,
            }
            self.dsl_executor = IntegratedDSLExecutor(self.app_functions, session_obj=None)
    
    def _load_roster_internal(self):
        """Internal method to load roster for DSL executor"""
        return self.roster_df
    
    def _save_roster_internal(self, df):
        """Internal method to save roster for DSL executor"""
        if self.roster_file:
            df.to_excel(self.roster_file, index=False, engine='openpyxl')
            self.roster_df = df
            return True
        return False
    
    def _format_date(self, date_input):
        """Format date for roster"""
        from datetime import datetime
        if isinstance(date_input, str):
            try:
                dt = datetime.strptime(date_input, "%Y-%m-%d")
            except ValueError:
                try:
                    dt = datetime.strptime(date_input, "%m/%d/%Y")
                except ValueError:
                    return date_input
        elif isinstance(date_input, datetime):
            dt = date_input
        else:
            return date_input
        return f"{dt.month}.{dt.day}"
    
    def _find_matching_date_column(self, df, date_input):
        """Find matching date column"""
        import re
        from datetime import datetime
        
        if isinstance(date_input, str):
            try:
                if '-' in date_input:
                    dt = datetime.fromisoformat(date_input)
                else:
                    dt = datetime.strptime(date_input, '%m/%d/%Y')
            except:
                return None
        elif isinstance(date_input, datetime):
            dt = date_input
        else:
            return None
        
        month_names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                      'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        month_abbr = month_names[dt.month - 1]
        
        possible_formats = [
            f"{month_abbr}.{dt.day}",
            f"{dt.month}.{dt.day}",
            f"R,{month_abbr}.{dt.day}",
            f"T,{month_abbr}.{dt.day}",
        ]
        
        for col in df.columns:
            col_str = str(col)
            for fmt in possible_formats:
                if fmt in col_str or col_str in fmt:
                    return col
        
        return None
    
    def clear_screen(self):
        """Clear the terminal screen"""
        os.system('cls' if os.name == 'nt' else 'clear')
    
    def print_header(self, title: str):
        """Print a formatted header"""
        print("\n" + "="*60)
        print(f"  {title}")
        print("="*60 + "\n")
    
    def load_roster(self, file_path: Optional[str] = None) -> bool:
        """Load roster file"""
        if not file_path:
            file_path = input("Enter roster file path (CSV or Excel): ").strip().strip('"\'')
        
        if not os.path.exists(file_path):
            print(f"Error: File not found: {file_path}")
            return False
        
        try:
            if file_path.lower().endswith('.csv'):
                self.roster_df = pd.read_csv(file_path, encoding='utf-8')
            else:
                self.roster_df = pd.read_excel(file_path, engine='openpyxl')
            
            self.roster_file = file_path
            print(f"✓ Roster loaded successfully: {len(self.roster_df)} students")
            return True
        except Exception as e:
            print(f"Error loading roster: {str(e)}")
            return False
    
    def process_attendance_with_gemini(self, attendance_file: str, date: Optional[str] = None):
        """Process attendance record using Gemini API to generate DSL code"""
        if self.roster_df is None:
            print("Error: Please load a roster file first.")
            return False
        
        if not self.gemini_model:
            print("Error: Gemini API not configured. Please set GEMINI_API_KEY environment variable.")
            return False
        
        # Resolve attendance file path
        original_path = attendance_file
        attendance_file = attendance_file.strip().strip('"\'')
        if not os.path.exists(attendance_file):
            # Try inside the default 'attendance record' folder
            if os.path.isdir(ATTENDANCE_FOLDER):
                candidate = os.path.join(ATTENDANCE_FOLDER, original_path)
                if os.path.exists(candidate):
                    attendance_file = candidate
                else:
                    print(f"Error: File not found: {original_path}")
                    print(f"Tried: {original_path} and {candidate}")
                    print("Tip: Put your files in the 'attendance record' folder or type the full path.")
                    return False
            else:
                print(f"Error: File not found: {original_path}")
                print("Tip: Create an 'attendance record' folder in this project, or type the full path.")
                return False
        
        # Read attendance file
        try:
            if attendance_file.lower().endswith('.csv'):
                attendance_df = pd.read_csv(attendance_file, encoding='utf-8')
            else:
                attendance_df = pd.read_excel(attendance_file, engine='openpyxl')
        except Exception as e:
            print(f"Error reading attendance file: {str(e)}")
            return False
        
        # Prepare context for Gemini
        roster_sample = self.roster_df.head(10).to_string()
        attendance_sample = attendance_df.head(10).to_string()
        
        # Read DSL specification
        dsl_spec = ""
        try:
            with open('dsl/ATTENDANCE_DSL.md', 'r', encoding='utf-8') as f:
                dsl_spec = f.read()
        except:
            print("Warning: Could not read DSL specification file.")
        
        # Use improved prompt template
        try:
            from dsl.gemini_prompts import create_attendance_processing_prompt
            prompt = create_attendance_processing_prompt(
                roster_sample=roster_sample,
                attendance_sample=attendance_sample,
                attendance_file=attendance_file,
                date=date,
                roster_file=self.roster_file or 'roster_attendance.xlsx'
            )
        except ImportError:
            # Fallback to basic prompt
            prompt = f"""
Generate DSL code to process attendance file: {attendance_file}
Meeting date: {date or 'auto-detect'}
Use PROCESS CHECKIN command with the attendance file.
"""
        
        try:
            print("Calling Gemini API to generate DSL code...")
            response = self.gemini_model.generate_content(prompt)
            dsl_code = response.text.strip()
            
            # Clean up the response (remove markdown code blocks if present)
            if dsl_code.startswith('```'):
                lines = dsl_code.split('\n')
                dsl_code = '\n'.join(lines[1:-1]) if len(lines) > 2 else dsl_code
            
            print("\nGenerated DSL Code:")
            print("-" * 60)
            print(dsl_code)
            print("-" * 60)
            
            # Execute the generated DSL code
            confirm = input("\nExecute this DSL code? (y/n): ").strip().lower()
            if confirm == 'y':
                return self.execute_dsl_code(dsl_code)
            else:
                print("DSL code generation cancelled.")
                return False
                
        except Exception as e:
            print(f"Error calling Gemini API: {str(e)}")
            return False
    
    def execute_dsl_code(self, dsl_code: str) -> bool:
        """Execute DSL code using the DSL executor"""
        if not self.dsl_executor:
            print("Error: DSL executor not available.")
            return False
        
        # Update executor's roster context
        self.dsl_executor.context['roster'] = self.roster_df
        self.dsl_executor.context['roster_file'] = self.roster_file
        
        try:
            result = self.dsl_executor.execute_script(dsl_code)
            
            if result['success']:
                print("\n✓ DSL code executed successfully!")
                for output in result.get('output', []):
                    if output and output.get('result'):
                        print(f"  {output['result']}")
                
                # Reload roster if it was modified
                if self.roster_file:
                    self.load_roster(self.roster_file)
                return True
            else:
                print(f"\n✗ Error executing DSL code: {result.get('error', 'Unknown error')}")
                if result.get('line_num'):
                    print(f"  Line {result['line_num']}: {result.get('line', '')}")
                return False
        except Exception as e:
            print(f"Error executing DSL code: {str(e)}")
            return False
    
    def query_with_gemini(self, user_query: str):
        """Use Gemini API to understand user query and generate DSL code"""
        if not self.gemini_model:
            print("Error: Gemini API not configured.")
            return False
        
        # Get current roster info
        roster_info = f"Roster has {len(self.roster_df)} students" if self.roster_df is not None else "No roster loaded"
        date_columns = []
        if self.roster_df is not None:
            date_columns = [str(col) for col in self.roster_df.columns 
                          if any(x in str(col).lower() for x in ['nov', 'oct', 'dec', 'jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', '.'])]
        
        # Read DSL specification
        dsl_spec = ""
        try:
            with open('dsl/ATTENDANCE_DSL.md', 'r', encoding='utf-8') as f:
                dsl_spec = f.read()
        except:
            print("Warning: Could not read DSL specification file.")
        
        # Use improved prompt template
        try:
            from dsl.gemini_prompts import create_query_prompt
            prompt = create_query_prompt(
                user_query=user_query,
                roster_info=roster_info,
                date_columns=date_columns,
                roster_file=self.roster_file or 'roster_attendance.xlsx'
            )
        except ImportError:
            # Fallback to basic prompt
            prompt = f"""
User request: {user_query}
Generate DSL code to fulfill this request.
"""
        
        try:
            print("Calling Gemini API to understand your request...")
            response = self.gemini_model.generate_content(prompt)
            dsl_code = response.text.strip()
            
            # Clean up the response
            if dsl_code.startswith('```'):
                lines = dsl_code.split('\n')
                dsl_code = '\n'.join(lines[1:-1]) if len(lines) > 2 else dsl_code
            
            print("\nGenerated DSL Code:")
            print("-" * 60)
            print(dsl_code)
            print("-" * 60)
            
            # Execute the generated DSL code
            confirm = input("\nExecute this DSL code? (y/n): ").strip().lower()
            if confirm == 'y':
                return self.execute_dsl_code(dsl_code)
            else:
                print("Execution cancelled.")
                return False
                
        except Exception as e:
            print(f"Error calling Gemini API: {str(e)}")
            return False
    
    def find_student_points_with_gemini(self, student_name: str):
        """Find a student's total points using Gemini API to generate DSL code"""
        if self.roster_df is None:
            print("Error: Please load a roster file first.")
            return False
        
        if not self.gemini_model:
            print("Error: Gemini API not configured. Please set GEMINI_API_KEY environment variable.")
            return False
        
        # Prepare roster context
        roster_sample = self.roster_df.head(10).to_string()
        roster_info = f"Roster has {len(self.roster_df)} students"
        
        # Get date columns
        date_columns = []
        for col in self.roster_df.columns:
            col_str = str(col).lower()
            if any(x in col_str for x in ['nov', 'oct', 'dec', 'jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', '.']) and col_str not in ['total', 'points']:
                date_columns.append(str(col))
        
        # Find name column
        name_col = None
        for col in self.roster_df.columns:
            col_str = str(col).lower().strip()
            if ('name' in col_str and 'unnamed' not in col_str and 
                col_str not in ['id', 'email', 'major', 'level']):
                name_col = col
                break
        if name_col is None:
            if len(self.roster_df.columns) > 2:
                name_col = self.roster_df.columns[2]
        
        # Try to find the student first (improved matching)
        student_found = False
        matching_students = []
        if name_col:
            # Normalize input name (remove extra spaces, handle comma variations)
            normalized_input = student_name.strip().replace(', ', ',').replace(' ,', ',')
            
            for idx, row in self.roster_df.iterrows():
                roster_name = str(row[name_col]).strip()
                if not roster_name or roster_name.lower() in ['nan', 'none', '']:
                    continue
                
                # Normalize roster name
                normalized_roster = roster_name.replace(', ', ',').replace(' ,', ',')
                
                # Try various matching strategies
                input_lower = normalized_input.lower()
                roster_lower = normalized_roster.lower()
                
                # Exact match
                if input_lower == roster_lower:
                    matching_students.insert(0, {
                        'index': idx,
                        'name': roster_name,
                        'total_points': self.roster_df.loc[idx, 'Total Points'] if 'Total Points' in self.roster_df.columns else None,
                        'match_quality': 'exact'
                    })
                    student_found = True
                # Partial match (contains or is contained)
                elif input_lower in roster_lower or roster_lower in input_lower:
                    matching_students.append({
                        'index': idx,
                        'name': roster_name,
                        'total_points': self.roster_df.loc[idx, 'Total Points'] if 'Total Points' in self.roster_df.columns else None,
                        'match_quality': 'partial'
                    })
                    student_found = True
                # Word-based matching (e.g., "Smith,John" matches "Smith, John")
                else:
                    input_words = set(input_lower.replace(',', ' ').split())
                    roster_words = set(roster_lower.replace(',', ' ').split())
                    if len(input_words.intersection(roster_words)) >= 2:  # At least 2 words match
                        matching_students.append({
                            'index': idx,
                            'name': roster_name,
                            'total_points': self.roster_df.loc[idx, 'Total Points'] if 'Total Points' in self.roster_df.columns else None,
                            'match_quality': 'word_match'
                        })
                        student_found = True
        
        # Use Gemini to generate DSL code
        try:
            from dsl.gemini_prompts import create_find_student_prompt
            prompt = create_find_student_prompt(
                student_name=student_name,
                roster_info=roster_info,
                roster_sample=roster_sample,
                date_columns=date_columns
            )
        except ImportError:
            # Fallback prompt
            prompt = f"""
Generate DSL code to find student "{student_name}" in the roster and show their total points.
"""
        
        try:
            print(f"Searching for student: {student_name}")
            print("Calling Gemini API to generate DSL code...")
            
            response = self.gemini_model.generate_content(prompt)
            dsl_code = response.text.strip()
            
            # Clean up the response
            if dsl_code.startswith('```'):
                lines = dsl_code.split('\n')
                dsl_code = '\n'.join(lines[1:-1]) if len(lines) > 2 else dsl_code
            
            # Direct student lookup (faster than DSL for this)
            if student_found:
                print(f"\n✓ Found {len(matching_students)} matching student(s):")
                print("-" * 60)
                for student in matching_students:
                    print(f"\nStudent: {student['name']}")
                    if student['total_points'] is not None:
                        print(f"Total Points: {student['total_points']}")
                    else:
                        # Calculate total points from date columns
                        total = 0.0
                        for col in date_columns:
                            val = self.roster_df.loc[student['index'], col]
                            if pd.notna(val):
                                try:
                                    total += float(val)
                                except (ValueError, TypeError):
                                    pass
                        print(f"Total Points: {total:.1f}")
                    
                    # Show attendance breakdown
                    print("\nAttendance Breakdown:")
                    for col in date_columns[:10]:  # Show first 10 dates
                        val = self.roster_df.loc[student['index'], col]
                        if pd.notna(val) and float(val) > 0:
                            print(f"  {col}: {val}")
                    if len(date_columns) > 10:
                        print(f"  ... and {len(date_columns) - 10} more date columns")
                print("-" * 60)
                
                print("\nGenerated DSL Code (for reference):")
                print("-" * 60)
                print(dsl_code)
                print("-" * 60)
            else:
                print(f"\nNo exact match found for '{student_name}'")
                print("\nGenerated DSL Code to search:")
                print("-" * 60)
                print(dsl_code)
                print("-" * 60)
                
                confirm = input("\nExecute this DSL code? (y/n): ").strip().lower()
                if confirm == 'y':
                    return self.execute_dsl_code(dsl_code)
            
            return True
                
        except Exception as e:
            print(f"Error calling Gemini API: {str(e)}")
            import traceback
            print(traceback.format_exc())
            return False
    
    def calculate_total_points(self):
        """Calculate total points for each student"""
        if self.roster_df is None:
            print("Error: No roster loaded.")
            return False
        
        # Find date columns
        date_columns = []
        for col in self.roster_df.columns:
            col_str = str(col).lower()
            if any(x in col_str for x in ['nov', 'oct', 'dec', 'jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', '.']) and col_str not in ['total', 'points']:
                date_columns.append(col)
        
        if not date_columns:
            print("No date columns found in roster.")
            return False
        
        # Calculate totals
        if 'Total Points' not in self.roster_df.columns:
            self.roster_df['Total Points'] = 0.0
        
        self.roster_df['Total Points'] = self.roster_df[date_columns].sum(axis=1, numeric_only=True, skipna=True)
        
        # Save roster
        if self.roster_file:
            self.roster_df.to_excel(self.roster_file, index=False, engine='openpyxl')
            print(f"✓ Total points calculated and saved to {self.roster_file}")
            return True
        
        return False
    
    def show_menu(self):
        """Display main menu"""
        self.print_header("Attendance Tracker Terminal")
        print("1. Load Roster File")
        print("2. Process Attendance Record (with Gemini AI)")
        print("3. Query/View Information (with Gemini AI)")
        print("4. Find Student's Total Points (with Gemini AI)")
        print("5. Execute DSL Code Manually")
        print("0. Exit")
        print()
    
    def run(self):
        """Main program loop"""
        while True:
            self.show_menu()
            choice = input("Enter your choice: ").strip()
            
            if choice == '0':
                print("\nGoodbye!")
                break
            elif choice == '1':
                self.clear_screen()
                self.print_header("Load Roster File")
                self.load_roster()
                input("\nPress Enter to continue...")
            elif choice == '2':
                self.clear_screen()
                self.print_header("Process Attendance Record")
                if self.roster_df is None:
                    print("Error: Please load a roster file first.")
                    input("\nPress Enter to continue...")
                    continue
                
                attendance_file = input("Enter attendance record file path (CSV or Excel): ").strip().strip('"\'')
                date = input("Enter meeting date (YYYY-MM-DD) or press Enter for auto-detect: ").strip()
                if not date:
                    date = None
                
                self.process_attendance_with_gemini(attendance_file, date)
                input("\nPress Enter to continue...")
            elif choice == '3':
                self.clear_screen()
                self.print_header("Query/View Information")
                if self.roster_df is None:
                    print("Error: Please load a roster file first.")
                    input("\nPress Enter to continue...")
                    continue
                
                query = input("Enter your query (e.g., 'show late students for November 4'): ").strip()
                if query:
                    self.query_with_gemini(query)
                input("\nPress Enter to continue...")
            elif choice == '4':
                self.clear_screen()
                self.print_header("Find Student's Total Points")
                if self.roster_df is None:
                    print("Error: Please load a roster file first.")
                    input("\nPress Enter to continue...")
                    continue
                
                print("Enter student name in 'Last Name, First Name' format (e.g., 'Smith, John'):")
                student_name = input("Student name: ").strip()
                if student_name:
                    self.find_student_points_with_gemini(student_name)
                else:
                    print("No student name provided.")
                input("\nPress Enter to continue...")
            elif choice == '5':
                self.clear_screen()
                self.print_header("Execute DSL Code Manually")
                if self.roster_df is None:
                    print("Error: Please load a roster file first.")
                    input("\nPress Enter to continue...")
                    continue
                
                print("Enter DSL code (type 'END' on a new line to finish):")
                dsl_lines = []
                while True:
                    line = input()
                    if line.strip().upper() == 'END':
                        break
                    dsl_lines.append(line)
                
                dsl_code = '\n'.join(dsl_lines)
                self.execute_dsl_code(dsl_code)
                input("\nPress Enter to continue...")
            else:
                print("Invalid choice. Please try again.")
                input("\nPress Enter to continue...")


def main():
    """Main entry point"""
    print("Initializing Attendance Tracker Terminal...")
    terminal = AttendanceTerminal()
    
    if not GEMINI_AVAILABLE:
        print("\nWarning: Gemini API not available. Install with: pip install google-generativeai")
        print("Some features will not work.\n")
    
    if not DSL_AVAILABLE:
        print("\nWarning: DSL executor not available. Check dsl module.")
        print("Some features will not work.\n")
    
    terminal.run()


if __name__ == '__main__':
    main()

