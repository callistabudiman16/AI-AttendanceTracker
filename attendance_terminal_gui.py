"""
GUI-Based Attendance Tracker with Gemini API Integration

This program provides a graphical interface for managing attendance records
with AI-assisted DSL code generation using Google's Gemini API.
"""

import os
import sys
import pandas as pd
import re
from datetime import datetime
from typing import Optional, Dict, Any, List
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading

# Import Gemini API
try:
    import google.generativeai as genai
    GEMINI_AVAILABLE = True
except ImportError:
    GEMINI_AVAILABLE = False

# Import DSL executor
try:
    from dsl.dsl_integrated import IntegratedDSLExecutor
    DSL_AVAILABLE = True
except ImportError:
    DSL_AVAILABLE = False

# Configuration
ROSTER_FILE = 'roster_attendance.xlsx'
ATTENDANCE_FOLDER = os.path.join(os.getcwd(), 'attendance record')
GEMINI_API_KEY_ENV = 'GEMINI_API_KEY'


class AttendanceTrackerGUI:
    """GUI-based attendance tracker with Gemini API integration"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Attendance Tracker - GUI")
        self.root.geometry("1000x700")
        
        self.roster_df: Optional[pd.DataFrame] = None
        self.roster_file: Optional[str] = None
        self.gemini_model = None
        self.dsl_executor = None
        self.app_functions = {}
        
        # Initialize Gemini API
        self.init_gemini()
        
        # Initialize DSL executor
        self.init_dsl_executor()
        
        # Create GUI
        self.create_widgets()
        
        # Load roster if default file exists
        if os.path.exists(ROSTER_FILE):
            self.load_roster(ROSTER_FILE, show_message=False)
    
    def init_gemini(self):
        """Initialize Gemini API"""
        if not GEMINI_AVAILABLE:
            return
        
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
            # Try available Gemini models
            model_names = [
                'gemini-2.0-flash',
                'gemini-flash-latest',
                'gemini-2.5-flash',
                'gemini-2.0-flash-001',
                'gemini-pro-latest',
                'gemini-2.5-pro',
            ]
            
            for model_name in model_names:
                try:
                    self.gemini_model = genai.GenerativeModel(model_name)
                    break
                except Exception:
                    continue
            
            if self.gemini_model is None:
                try:
                    models = genai.list_models()
                    available = [m.name for m in models if 'generateContent' in m.supported_generation_methods]
                    if available:
                        first_model = available[0].replace('models/', '')
                        self.gemini_model = genai.GenerativeModel(first_model)
                except Exception:
                    pass
    
    def init_dsl_executor(self):
        """Initialize DSL executor"""
        if DSL_AVAILABLE:
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
        """Internal method to save roster for DSL executor with error handling"""
        if not self.roster_file:
            return False
        
        import time
        
        # Try saving with retry logic
        max_retries = 3
        retry_delay = 0.5
        
        for attempt in range(max_retries):
            try:
                # Check if file is locked (Windows)
                if os.name == 'nt' and os.path.exists(self.roster_file):
                    try:
                        # Try to open the file to check if it's locked
                        test_file = open(self.roster_file, 'r+b')
                        test_file.close()
                    except (PermissionError, IOError):
                        if attempt < max_retries - 1:
                            time.sleep(retry_delay)
                            continue
                        error_msg = (
                            f'Cannot save roster: The file "{os.path.basename(self.roster_file)}" is currently open in another program.\n\n'
                            f'Please close the file in Excel or any other program and try again.'
                        )
                        messagebox.showerror("File Locked", error_msg)
                        self.log(f"ERROR: {error_msg}", "error")
                        return False
                
                # Save to Excel
                df.to_excel(self.roster_file, index=False, engine='openpyxl')
                self.roster_df = df
                return True
                
            except PermissionError:
                if attempt < max_retries - 1:
                    time.sleep(retry_delay)
                    continue
                error_msg = (
                    f'Permission denied: Cannot save roster file "{os.path.basename(self.roster_file)}".\n\n'
                    f'The file may be open in Microsoft Excel or another program.\n'
                    f'Please close the file and try again.'
                )
                messagebox.showerror("Permission Denied", error_msg)
                self.log(f"ERROR: {error_msg}", "error")
                return False
                
            except Exception as e:
                if 'Permission denied' in str(e) or '[Errno 13]' in str(e):
                    if attempt < max_retries - 1:
                        time.sleep(retry_delay)
                        continue
                    error_msg = (
                        f'Cannot save roster: The file "{os.path.basename(self.roster_file)}" is locked.\n\n'
                        f'Possible causes:\n'
                        f'1. File is open in Microsoft Excel - Please close it\n'
                        f'2. Another process is using the file\n'
                        f'3. File permissions issue\n\n'
                        f'Please close the file in Excel and try again.'
                    )
                    messagebox.showerror("File Locked", error_msg)
                    self.log(f"ERROR: {error_msg}", "error")
                    return False
                else:
                    if attempt < max_retries - 1:
                        time.sleep(retry_delay)
                        continue
                    error_msg = f'Error saving roster: {str(e)}'
                    messagebox.showerror("Save Error", error_msg)
                    self.log(f"ERROR: {error_msg}", "error")
                    return False
        
        return False
    
    def _format_date(self, date_input):
        """Format date for roster"""
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
    
    def create_widgets(self):
        """Create GUI widgets"""
        # Configure styles for colorful theme
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure colorful styles
        style.configure('Title.TLabel', font=('Arial', 14, 'bold'), foreground='#1e40af', background='#eff6ff')
        style.configure('Status.TLabel', font=('Arial', 10, 'bold'), foreground='#059669', background='#ecfdf5')
        style.configure('Info.TLabel', font=('Arial', 9), foreground='#475569', background='#f8fafc')
        
        # Button styles with colors
        style.configure('Blue.TButton', background='#3b82f6', foreground='white', font=('Arial', 9, 'bold'))
        style.map('Blue.TButton', background=[('active', '#2563eb'), ('pressed', '#1d4ed8')])
        
        style.configure('Purple.TButton', background='#8b5cf6', foreground='white', font=('Arial', 9, 'bold'))
        style.map('Purple.TButton', background=[('active', '#7c3aed'), ('pressed', '#6d28d9')])
        
        style.configure('Teal.TButton', background='#14b8a6', foreground='white', font=('Arial', 9, 'bold'))
        style.map('Teal.TButton', background=[('active', '#0d9488'), ('pressed', '#0f766e')])
        
        style.configure('Orange.TButton', background='#f97316', foreground='white', font=('Arial', 9, 'bold'))
        style.map('Orange.TButton', background=[('active', '#ea580c'), ('pressed', '#c2410c')])
        
        style.configure('Green.TButton', background='#10b981', foreground='white', font=('Arial', 9, 'bold'))
        style.map('Green.TButton', background=[('active', '#059669'), ('pressed', '#047857')])
        
        style.configure('Red.TButton', background='#ef4444', foreground='white', font=('Arial', 9, 'bold'))
        style.map('Red.TButton', background=[('active', '#dc2626'), ('pressed', '#b91c1c')])
        
        # Frame styles
        style.configure('Header.TFrame', background='#3b82f6')
        style.configure('Blue.TLabelframe', background='#eff6ff', borderwidth=2, relief='solid')
        style.configure('Blue.TLabelframe.Label', background='#eff6ff', foreground='#1e40af', font=('Arial', 10, 'bold'))
        
        # Main container with gradient-like background
        main_frame = tk.Frame(self.root, bg='#f0f9ff')
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.root.configure(bg='#f0f9ff')
        
        # Top frame - Status and file info with blue header
        top_frame = tk.Frame(main_frame, bg='#3b82f6', relief='raised', bd=2)
        top_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        top_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = tk.Label(top_frame, text="📊 Attendance Tracker", 
                              font=('Arial', 16, 'bold'), bg='#3b82f6', fg='white')
        title_label.grid(row=0, column=0, padx=15, pady=10, sticky=tk.W)
        
        # Status and info in a sub-frame
        info_frame = tk.Frame(top_frame, bg='#3b82f6')
        info_frame.grid(row=0, column=1, padx=15, pady=10, sticky=tk.E)
        
        self.status_label = tk.Label(info_frame, text="Status: Ready", 
                                     font=('Arial', 10, 'bold'), bg='#3b82f6', fg='#fef3c7')
        self.status_label.grid(row=0, column=0, padx=5, sticky=tk.W)
        
        self.roster_info_label = tk.Label(info_frame, text="No roster loaded", 
                                          font=('Arial', 9), bg='#3b82f6', fg='#e0e7ff')
        self.roster_info_label.grid(row=0, column=1, padx=10, sticky=tk.W)
        
        # Left frame - Buttons with colorful styling
        left_frame = ttk.LabelFrame(main_frame, text="⚡ Actions", padding="10", style='Blue.TLabelframe')
        left_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(10, 5), pady=10)
        left_frame.columnconfigure(0, weight=1)
        
        # Buttons with different colors
        btn1 = tk.Button(left_frame, text="📁 1. Load Roster File", 
                        command=self.load_roster_dialog, width=32, height=2,
                        bg='#3b82f6', fg='white', font=('Arial', 9, 'bold'),
                        activebackground='#2563eb', activeforeground='white',
                        relief='raised', bd=2, cursor='hand2')
        btn1.grid(row=0, column=0, pady=6, sticky=(tk.W, tk.E))
        
        btn2 = tk.Button(left_frame, text="✅ 2. Process Attendance Record", 
                        command=self.process_attendance_dialog, width=32, height=2,
                        bg='#8b5cf6', fg='white', font=('Arial', 9, 'bold'),
                        activebackground='#7c3aed', activeforeground='white',
                        relief='raised', bd=2, cursor='hand2')
        btn2.grid(row=1, column=0, pady=6, sticky=(tk.W, tk.E))
        
        btn3 = tk.Button(left_frame, text="🔍 3. Query/View Information", 
                        command=self.query_dialog, width=32, height=2,
                        bg='#14b8a6', fg='white', font=('Arial', 9, 'bold'),
                        activebackground='#0d9488', activeforeground='white',
                        relief='raised', bd=2, cursor='hand2')
        btn3.grid(row=2, column=0, pady=6, sticky=(tk.W, tk.E))
        
        btn4 = tk.Button(left_frame, text="👤 4. Find Student's Total Points", 
                        command=self.find_student_dialog, width=32, height=2,
                        bg='#f97316', fg='white', font=('Arial', 9, 'bold'),
                        activebackground='#ea580c', activeforeground='white',
                        relief='raised', bd=2, cursor='hand2')
        btn4.grid(row=3, column=0, pady=6, sticky=(tk.W, tk.E))
        
        btn5 = tk.Button(left_frame, text="⚡ 5. Execute DSL Code", 
                        command=self.execute_dsl_dialog, width=32, height=2,
                        bg='#10b981', fg='white', font=('Arial', 9, 'bold'),
                        activebackground='#059669', activeforeground='white',
                        relief='raised', bd=2, cursor='hand2')
        btn5.grid(row=4, column=0, pady=6, sticky=(tk.W, tk.E))
        
        # Right frame - Output with colorful styling
        right_frame = ttk.LabelFrame(main_frame, text="💻 Output", padding="10", style='Blue.TLabelframe')
        right_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(5, 10), pady=10)
        right_frame.columnconfigure(0, weight=1)
        right_frame.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # Output text area with dark theme
        self.output_text = scrolledtext.ScrolledText(right_frame, wrap=tk.WORD, width=60, height=30, 
                                                     font=('Consolas', 10), bg='#1e293b', fg='#10b981',
                                                     insertbackground='#10b981', selectbackground='#334155')
        self.output_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Bottom frame - Clear button with colors
        bottom_frame = tk.Frame(main_frame, bg='#f0f9ff')
        bottom_frame.grid(row=2, column=0, columnspan=2, pady=(5, 10))
        
        clear_btn = tk.Button(bottom_frame, text="🗑️ Clear Output", 
                             command=self.clear_output, width=15, height=1,
                             bg='#64748b', fg='white', font=('Arial', 9, 'bold'),
                             activebackground='#475569', activeforeground='white',
                             relief='raised', bd=2, cursor='hand2')
        clear_btn.grid(row=0, column=0, padx=5)
        
        exit_btn = tk.Button(bottom_frame, text="❌ Exit", 
                            command=self.root.quit, width=15, height=1,
                            bg='#ef4444', fg='white', font=('Arial', 9, 'bold'),
                            activebackground='#dc2626', activeforeground='white',
                            relief='raised', bd=2, cursor='hand2')
        exit_btn.grid(row=0, column=1, padx=5)
    
    def log(self, message: str, level: str = "info"):
        """Log message to output text area with color coding"""
        # Color code based on level
        tag = f"tag_{level}"
        
        # Configure tags for different log levels
        if level == "error":
            self.output_text.tag_config(tag, foreground='#ef4444', font=('Consolas', 10, 'bold'))
        elif level == "success":
            self.output_text.tag_config(tag, foreground='#10b981', font=('Consolas', 10, 'bold'))
        elif level == "warning":
            self.output_text.tag_config(tag, foreground='#f59e0b', font=('Consolas', 10))
        elif level == "info":
            self.output_text.tag_config(tag, foreground='#60a5fa', font=('Consolas', 10))
        else:
            self.output_text.tag_config(tag, foreground='#10b981', font=('Consolas', 10))
        
        # Insert with tag
        self.output_text.insert(tk.END, f"[{level.upper()}] {message}\n", tag)
        self.output_text.see(tk.END)
        self.root.update_idletasks()
    
    def clear_output(self):
        """Clear output text area"""
        self.output_text.delete(1.0, tk.END)
    
    def update_status(self, message: str):
        """Update status label"""
        # Color code status messages
        if "Error" in message or "error" in message.lower():
            color = '#fca5a5'  # Light red
        elif "Processing" in message or "Loading" in message:
            color = '#fde047'  # Yellow
        elif "Ready" in message:
            color = '#86efac'  # Light green
        else:
            color = '#fef3c7'  # Default yellow
        
        self.status_label.config(text=f"Status: {message}", fg=color)
        self.root.update_idletasks()
    
    def update_roster_info(self):
        """Update roster info label"""
        if self.roster_df is not None:
            info = f"📋 Roster: {len(self.roster_df)} students | 📄 File: {os.path.basename(self.roster_file) if self.roster_file else 'N/A'}"
            self.roster_info_label.config(text=info, fg='#c7d2fe')
        else:
            self.roster_info_label.config(text="⚠️ No roster loaded", fg='#fbbf24')
    
    def load_roster_dialog(self):
        """Open dialog to load roster file"""
        file_path = filedialog.askopenfilename(
            title="Select Roster File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if file_path:
            self.load_roster(file_path)
    
    def load_roster(self, file_path: Optional[str] = None, show_message: bool = True) -> bool:
        """Load roster file"""
        if not file_path:
            return False
        
        if not os.path.exists(file_path):
            if show_message:
                messagebox.showerror("Error", f"File not found: {file_path}")
            return False
        
        try:
            self.update_status("Loading roster...")
            if file_path.lower().endswith('.csv'):
                self.roster_df = pd.read_csv(file_path, encoding='utf-8')
            else:
                self.roster_df = pd.read_excel(file_path, engine='openpyxl')
            
            self.roster_file = file_path
            self.update_roster_info()
            self.log(f"✓ Roster loaded successfully: {len(self.roster_df)} students", "success")
            self.update_status("Ready")
            return True
        except Exception as e:
            if show_message:
                messagebox.showerror("Error", f"Error loading roster: {str(e)}")
            self.log(f"Error loading roster: {str(e)}", "error")
            self.update_status("Error")
            return False
    
    def process_attendance_dialog(self):
        """Open dialog to process attendance record"""
        if self.roster_df is None:
            messagebox.showwarning("Warning", "Please load a roster file first.")
            return
        
        # Create dialog window
        dialog = tk.Toplevel(self.root)
        dialog.title("Process Attendance Record")
        dialog.geometry("800x700")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(True, True)
        dialog.minsize(750, 650)
        
        # Configure dialog columns and rows
        dialog.columnconfigure(0, weight=1)
        dialog.rowconfigure(0, weight=1)  # Main container expands
        
        # Create main container for scrollable content
        main_container = ttk.Frame(dialog)
        main_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=0, pady=0)
        main_container.columnconfigure(0, weight=1)
        
        # Make dialog modal and visible
        dialog.focus_set()
        dialog.lift()
        
        # Attendance type selection frame
        type_frame = ttk.LabelFrame(main_container, text="Class Type", padding="15")
        type_frame.grid(row=0, column=0, padx=15, pady=15, sticky=(tk.W, tk.E))
        
        attendance_type = tk.StringVar(value="in_person")
        
        ttk.Radiobutton(
            type_frame, 
            text="• In Person", 
            variable=attendance_type, 
            value="in_person"
        ).grid(row=0, column=0, padx=20, pady=10, sticky=tk.W)
        
        ttk.Radiobutton(
            type_frame, 
            text="• Zoom Meeting", 
            variable=attendance_type, 
            value="zoom"
        ).grid(row=0, column=1, padx=20, pady=10, sticky=tk.W)
        
        # Settings frame (will show different options based on type)
        settings_frame = ttk.LabelFrame(main_container, text="Settings", padding="15")
        settings_frame.grid(row=1, column=0, padx=15, pady=15, sticky=(tk.W, tk.E))
        settings_frame.columnconfigure(1, weight=1)
        
        # In-person time settings (initially visible)
        in_person_frame = ttk.Frame(settings_frame)
        in_person_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E))
        
        ttk.Label(in_person_frame, text="Start Time:").grid(row=0, column=0, padx=(0, 10), pady=5, sticky=tk.W)
        start_time_entry = ttk.Entry(in_person_frame, width=15)
        start_time_entry.insert(0, "11:00")
        start_time_entry.grid(row=0, column=1, padx=(0, 20), pady=5, sticky=tk.W)
        
        ttk.Label(in_person_frame, text="End Time:").grid(row=0, column=2, padx=(0, 10), pady=5, sticky=tk.W)
        end_time_entry = ttk.Entry(in_person_frame, width=15)
        end_time_entry.insert(0, "11:35")
        end_time_entry.grid(row=0, column=3, padx=(0, 10), pady=5, sticky=tk.W)
        
        ttk.Label(in_person_frame, text="Between start and end time: 0.6 points | After end time: 0.2 points", 
                 foreground='darkblue').grid(row=1, column=0, columnspan=4, padx=5, pady=(5, 0), sticky=tk.W)
        
        # Zoom cut time settings (initially hidden)
        zoom_frame = ttk.Frame(settings_frame)
        zoom_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E))
        zoom_frame.grid_remove()  # Hide initially
        
        ttk.Label(zoom_frame, text="Cut Time (minutes):").grid(row=0, column=0, padx=(0, 10), pady=5, sticky=tk.W)
        cut_time_entry = ttk.Entry(zoom_frame, width=15)
        cut_time_entry.insert(0, "30")
        cut_time_entry.grid(row=0, column=1, padx=(0, 10), pady=5, sticky=tk.W)
        
        ttk.Label(zoom_frame, text="(Duration threshold in minutes)", foreground='gray').grid(row=1, column=1, padx=(0, 10), pady=(0, 5), sticky=tk.W)
        ttk.Label(zoom_frame, text="Duration ≥ cut time: 0.6 points | Duration < cut time: 0.2 points", 
                 foreground='darkblue').grid(row=2, column=0, columnspan=2, padx=5, pady=(5, 0), sticky=tk.W)
        
        # Function to toggle settings visibility
        def toggle_settings(*args):
            if attendance_type.get() == "in_person":
                in_person_frame.grid()
                zoom_frame.grid_remove()
            else:
                in_person_frame.grid_remove()
                zoom_frame.grid()
        
        attendance_type.trace('w', toggle_settings)
        
        # Date entry frame
        date_frame = ttk.LabelFrame(main_container, text="Meeting Date (Optional)", padding="15")
        date_frame.grid(row=2, column=0, padx=15, pady=15, sticky=(tk.W, tk.E))
        date_frame.columnconfigure(1, weight=1)
        
        ttk.Label(date_frame, text="Date:").grid(row=0, column=0, padx=(0, 10), pady=5, sticky=tk.W)
        date_entry = ttk.Entry(date_frame, width=30)
        date_entry.grid(row=0, column=1, padx=(0, 10), pady=5, sticky=tk.W)
        ttk.Label(date_frame, text="(YYYY-MM-DD format, leave empty for auto-detect)", foreground='gray').grid(row=1, column=1, padx=(0, 10), pady=(0, 5), sticky=tk.W)
        
        # File selection frame
        file_frame = ttk.LabelFrame(main_container, text="Upload Attendance File", padding="15")
        file_frame.grid(row=3, column=0, padx=15, pady=15, sticky=(tk.W, tk.E))
        file_frame.columnconfigure(1, weight=1)
        
        ttk.Label(file_frame, text="File:").grid(row=0, column=0, padx=(0, 10), pady=10, sticky=tk.W)
        file_entry = ttk.Entry(file_frame, width=50)
        file_entry.grid(row=0, column=1, padx=(0, 10), pady=10, sticky=(tk.W, tk.E))
        
        def browse_file():
            initial_dir = ATTENDANCE_FOLDER if os.path.isdir(ATTENDANCE_FOLDER) else os.getcwd()
            file_path = filedialog.askopenfilename(
                title="Select Attendance File",
                initialdir=initial_dir,
                filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("All files", "*.*")]
            )
            if file_path:
                file_entry.delete(0, tk.END)
                file_entry.insert(0, file_path)
        
        browse_btn = ttk.Button(file_frame, text="Browse...", command=browse_file, width=12)
        browse_btn.grid(row=0, column=2, padx=5, pady=10, sticky=tk.E)
        
        # Helper text
        help_text = f"Enter file path or filename (files in 'attendance record' folder can use just filename)"
        ttk.Label(file_frame, text=help_text, foreground='gray').grid(row=1, column=0, columnspan=3, padx=5, pady=(0, 5), sticky=tk.W)
        
        # Buttons frame - always visible at bottom of dialog
        button_frame = ttk.Frame(dialog)
        button_frame.grid(row=1, column=0, pady=20, padx=15, sticky=(tk.E, tk.W))
        dialog.rowconfigure(1, weight=0)  # Buttons row doesn't expand
        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=1)
        
        def process():
            try:
                attendance_file = file_entry.get().strip().strip('"\'')
                date = date_entry.get().strip() or None
                atype = attendance_type.get()
                
                if not attendance_file:
                    messagebox.showwarning("Warning", "Please enter or select an attendance file.")
                    file_entry.focus()
                    return
                
                # Validate based on attendance type
                if atype == "in_person":
                    start_time_str = start_time_entry.get().strip()
                    end_time_str = end_time_entry.get().strip()
                    
                    if not start_time_str or not end_time_str:
                        messagebox.showerror("Error", "Please enter both start time and end time.")
                        return
                    
                    try:
                        datetime.strptime(start_time_str, '%H:%M')
                    except ValueError:
                        messagebox.showerror("Error", f"Invalid start time format: {start_time_str}\n\nPlease use HH:MM format (e.g., 11:00)")
                        start_time_entry.focus()
                        return
                    
                    try:
                        datetime.strptime(end_time_str, '%H:%M')
                    except ValueError:
                        messagebox.showerror("Error", f"Invalid end time format: {end_time_str}\n\nPlease use HH:MM format (e.g., 11:35)")
                        end_time_entry.focus()
                        return
                else:  # zoom
                    cut_time_str = cut_time_entry.get().strip()
                    if not cut_time_str:
                        messagebox.showerror("Error", "Please enter a cut time (minutes).")
                        cut_time_entry.focus()
                        return
                    
                    try:
                        cut_time_minutes = int(cut_time_str)
                        if cut_time_minutes <= 0:
                            raise ValueError("Cut time must be greater than 0")
                    except ValueError as e:
                        messagebox.showerror("Error", f"Invalid cut time: {cut_time_str}\n\nPlease enter a positive number (minutes)")
                        cut_time_entry.focus()
                        return
                
                # Resolve file path
                resolved_path = attendance_file
                if not os.path.exists(attendance_file):
                    if os.path.isdir(ATTENDANCE_FOLDER):
                        candidate = os.path.join(ATTENDANCE_FOLDER, attendance_file)
                        if os.path.exists(candidate):
                            resolved_path = candidate
                        else:
                            messagebox.showerror("Error", f"File not found: {attendance_file}\n\nTried:\n- {attendance_file}\n- {candidate}\n\nPlease check the file path.")
                            file_entry.focus()
                            return
                    else:
                        messagebox.showerror("Error", f"File not found: {attendance_file}\n\nPlease check the file path.")
                        file_entry.focus()
                        return
                
                # Store values before destroying dialog
                if atype == "in_person":
                    final_start_time = start_time_entry.get().strip()
                    final_end_time = end_time_entry.get().strip()
                    final_cut_time = None
                else:
                    final_start_time = None
                    final_end_time = None
                    final_cut_time = int(cut_time_entry.get().strip())
                
                # Close dialog before processing
                dialog.destroy()
                
                # Provide immediate feedback
                self.update_status("Starting attendance processing...")
                self.log(f"Processing attendance file: {resolved_path}", "info")
                
                # Process attendance (runs in background thread)
                if atype == "in_person":
                    self.process_attendance_with_gemini(resolved_path, date, final_start_time, final_end_time)
                else:  # zoom
                    self.process_zoom_attendance(resolved_path, date, final_cut_time)
                    
            except Exception as e:
                import traceback
                error_msg = f"Error in process function: {str(e)}\n\n{traceback.format_exc()}"
                messagebox.showerror("Error", error_msg)
                self.log(f"ERROR: {error_msg}", "error")
        
        def on_process_click():
            """Wrapper to handle button click and provide feedback"""
            try:
                self.log("Process button clicked. Validating inputs...", "info")
                process()
            except Exception as e:
                import traceback
                error_msg = f"Error when clicking Process button: {str(e)}\n\n{traceback.format_exc()}"
                messagebox.showerror("Error", error_msg)
                self.log(f"ERROR: {error_msg}", "error")
        
        process_btn = ttk.Button(button_frame, text="Process", command=on_process_click, width=20)
        process_btn.grid(row=0, column=0, padx=10, pady=5, sticky=(tk.E, tk.W))
        
        cancel_btn = ttk.Button(button_frame, text="Cancel", command=dialog.destroy, width=20)
        cancel_btn.grid(row=0, column=1, padx=10, pady=5, sticky=(tk.E, tk.W))
        
        # Bind Enter key to process
        file_entry.bind('<Return>', lambda e: process())
        date_entry.bind('<Return>', lambda e: process())
        start_time_entry.bind('<Return>', lambda e: process())
        end_time_entry.bind('<Return>', lambda e: process())
        cut_time_entry.bind('<Return>', lambda e: process())
        
        # Center dialog and ensure it's visible
        dialog.update_idletasks()
        
        # Calculate center position
        screen_width = dialog.winfo_screenwidth()
        screen_height = dialog.winfo_screenheight()
        window_width = dialog.winfo_reqwidth()
        window_height = dialog.winfo_reqheight()
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        
        # Ensure window is on screen
        x = max(0, min(x, screen_width - window_width))
        y = max(0, min(y, screen_height - window_height))
        
        dialog.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # Ensure dialog is on top and visible
        dialog.lift()
        dialog.focus_force()
        file_entry.focus()  # Focus on file entry for better UX
    
    def process_attendance_with_gemini(self, attendance_file: str, date: Optional[str] = None, start_time: str = "11:00", end_time: str = "11:35"):
        """Process attendance record directly (bypass DSL)"""
        def process_thread():
            try:
                self.update_status("Processing attendance...")
                self.log(f"Processing attendance file: {attendance_file}")
                
                # Import processing functions from app.py
                import sys
                import importlib.util
                
                # Get the app.py file path
                app_path = os.path.join(os.getcwd(), 'app.py')
                if not os.path.exists(app_path):
                    self.log("Error: app.py not found. Cannot process attendance.", "error")
                    self.update_status("Error")
                    return
                
                # Import required functions from app.py
                spec = importlib.util.spec_from_file_location("app_module", app_path)
                app_module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(app_module)
                
                # Get processing functions
                find_student_in_roster = app_module.find_student_in_roster
                update_roster_with_attendance = app_module.update_roster_with_attendance
                format_date_for_roster = app_module.format_date_for_roster
                find_matching_date_column = app_module.find_matching_date_column
                
                # File path should already be resolved by dialog
                attendance_file_resolved = attendance_file.strip().strip('"\'')
                
                # Read attendance file
                if attendance_file_resolved.lower().endswith('.csv'):
                    try:
                        checkins_df = pd.read_csv(attendance_file_resolved, encoding='utf-8')
                    except (UnicodeDecodeError, UnicodeError):
                        checkins_df = pd.read_csv(attendance_file_resolved, encoding='latin-1')
                else:
                    checkins_df = pd.read_excel(attendance_file_resolved, engine='openpyxl')
                
                self.log(f"Read {len(checkins_df)} rows from attendance file", "info")
                
                # Find Start Date column
                start_date_col = None
                for col in checkins_df.columns:
                    col_lower = str(col).lower().strip()
                    if 'start date' in col_lower or 'startdate' in col_lower.replace(' ', ''):
                        start_date_col = col
                        break
                
                # Find name column
                name_col = None
                excluded_cols = [start_date_col] if start_date_col else []
                
                def is_name_column(col_name, df_col):
                    if col_name in excluded_cols:
                        return False
                    sample_values = df_col.dropna().head(10)
                    if len(sample_values) == 0:
                        return False
                    
                    name_count = 0
                    total = 0
                    for val in sample_values:
                        val_str = str(val).strip()
                        if not val_str or val_str.lower() in ['nan', 'none', '']:
                            continue
                        # Check if it looks like a date
                        if re.match(r'^\d{4}-\d{2}-\d{2}', val_str) or re.match(r'^\d{2}/\d{2}/\d{4}', val_str):
                            continue
                        # Check if it looks like coordinates (latitude/longitude)
                        try:
                            float_val = float(val_str)
                            if -90 <= float_val <= 90 or -180 <= float_val <= 180:
                                continue  # Likely coordinates
                        except:
                            pass
                        # Check if it looks like a name (has letters, contains spaces or commas)
                        if any(c.isalpha() for c in val_str) and (' ' in val_str or ',' in val_str):
                            name_count += 1
                        total += 1
                    
                    return total > 0 and (name_count / total) > 0.5
                
                # Try to find name column - check all columns
                best_name_col = None
                best_name_ratio = 0
                
                for col in checkins_df.columns:
                    if col in excluded_cols:
                        continue
                    
                    col_lower = str(col).lower().strip()
                    
                    # Skip obvious non-name columns
                    if any(skip in col_lower for skip in ['latitude', 'longitude', 'location', 'date', 'time', 'id', 'email']):
                        continue
                    
                    # Check if this column contains name-like data
                    if is_name_column(col, checkins_df[col]):
                        # Calculate name ratio
                        sample_values = checkins_df[col].dropna().head(10)
                        name_like = 0
                        total_valid = 0
                        for val in sample_values:
                            val_str = str(val).strip()
                            if not val_str or val_str.lower() in ['nan', 'none', '']:
                                continue
                            if re.match(r'^\d{4}-\d{2}-\d{2}', val_str) or re.match(r'^\d{2}/\d{2}/\d{4}', val_str):
                                continue
                            try:
                                float_val = float(val_str)
                                if -90 <= float_val <= 90 or -180 <= float_val <= 180:
                                    continue
                            except:
                                pass
                            if any(c.isalpha() for c in val_str) and (' ' in val_str or ',' in val_str):
                                name_like += 1
                            total_valid += 1
                        
                        if total_valid > 0:
                            name_ratio = name_like / total_valid
                            if name_ratio > best_name_ratio:
                                best_name_ratio = name_ratio
                                best_name_col = col
                
                # Also try exact "name" match first (highest priority)
                for col in checkins_df.columns:
                    if str(col).lower().strip() == 'name':
                        if is_name_column(col, checkins_df[col]):
                            name_col = col
                            break
                
                # If exact match not found, use best match
                if name_col is None:
                    name_col = best_name_col
                
                if name_col is None:
                    self.log(f"Error: Name column not found. Available columns: {list(checkins_df.columns)}", "error")
                    self.log("Tip: Qualtrics exports may use question text as column names. Please check the file.", "info")
                    self.update_status("Error")
                    return
                
                self.log(f"Detected columns - Name: {name_col}, Start Date: {start_date_col if start_date_col else 'Not found'}", "info")
                
                # Get time settings from parameters
                try:
                    start_time_obj = datetime.strptime(start_time, '%H:%M').time()
                    end_time_obj = datetime.strptime(end_time, '%H:%M').time()
                except ValueError as e:
                    self.log(f"Error parsing time settings: {str(e)}. Using defaults (11:00-11:35)", "warning")
                    start_time_obj = datetime.strptime('11:00', '%H:%M').time()
                    end_time_obj = datetime.strptime('11:35', '%H:%M').time()
                
                self.log(f"Time window: {start_time} to {end_time} = 0.6 points, after {end_time} = 0.2 points", "info")
                
                # Make a copy of roster for processing
                roster_df = self.roster_df.copy()
                
                processed_count = 0
                errors = []
                
                # Process each check-in
                for idx, row in checkins_df.iterrows():
                    student_name = str(row[name_col]).strip()
                    if not student_name or student_name.lower() in ['nan', 'none', '']:
                        continue
                    
                    # Skip if name looks like a date
                    if re.match(r'^\d{4}-\d{2}-\d{2}', student_name) or re.match(r'^\d{2}/\d{2}/\d{4}', student_name):
                        continue
                    
                    # Get check-in date and time
                    check_in_datetime = None
                    meeting_date = None
                    
                    if start_date_col and pd.notna(row.get(start_date_col)):
                        try:
                            check_in_datetime = pd.to_datetime(row[start_date_col])
                            meeting_date = check_in_datetime.date()
                        except:
                            try:
                                check_in_datetime = datetime.strptime(str(row[start_date_col]), '%Y-%m-%d %H:%M:%S')
                                meeting_date = check_in_datetime.date()
                            except:
                                pass
                    
                    # Use provided date or default to today
                    if meeting_date is None:
                        if date:
                            try:
                                meeting_date = datetime.strptime(date, '%Y-%m-%d').date()
                                check_in_datetime = datetime.combine(meeting_date, datetime.now().time())
                            except:
                                meeting_date = datetime.now().date()
                                check_in_datetime = datetime.now()
                        else:
                            meeting_date = datetime.now().date()
                            check_in_datetime = datetime.now()
                    
                    check_in_time = check_in_datetime.time() if check_in_datetime else datetime.now().time()
                    
                    # Calculate points based on check-in time
                    # If check-in is between start_time and end_time (inclusive): 0.6 points
                    # If check-in is after end_time: 0.2 points
                    # If check-in is before start_time: also 0.6 points (early bird)
                    check_in_dt = datetime.combine(meeting_date, check_in_time)
                    start_dt = datetime.combine(meeting_date, start_time_obj)
                    end_dt = datetime.combine(meeting_date, end_time_obj)
                    
                    if check_in_dt > end_dt:
                        points = 0.2  # Late check-in (after end time)
                    else:
                        points = 0.6  # On-time or early check-in (before or during the window)
                    
                    # Format date and find matching column
                    meeting_datetime = datetime.combine(meeting_date, datetime.min.time())
                    date_str = format_date_for_roster(meeting_datetime)
                    matching_date_col = find_matching_date_column(roster_df, meeting_datetime)
                    if matching_date_col:
                        date_str = matching_date_col
                    
                    # Update roster
                    roster_df, found, confidence, matched_name = update_roster_with_attendance(
                        roster_df, student_name, points, date_str, use_gemini=False
                    )
                    
                    if found:
                        processed_count += 1
                        if confidence < 1.0:
                            self.log(f"✓ {student_name} → {matched_name} ({confidence:.1%} confidence) - {points} points", "info")
                    else:
                        errors.append(f"Could not match: {student_name} (confidence: {confidence:.2f})")
                
                # Update roster in GUI
                self.roster_df = roster_df
                
                # Save roster
                if self._save_roster_internal(roster_df):
                    self.log(f"\n✓ Processed {processed_count} check-ins successfully", "success")
                    if errors:
                        self.log(f"⚠ {len(errors)} unmatched students (see above)", "warning")
                        if len(errors) <= 10:
                            for err in errors:
                                self.log(f"  {err}", "warning")
                        else:
                            for err in errors[:10]:
                                self.log(f"  {err}", "warning")
                            self.log(f"  ... and {len(errors) - 10} more", "warning")
                    
                    self.update_roster_info()
                    self.update_status("Ready")
                    self.log("✓ Roster file updated successfully!", "success")
                else:
                    self.log("Error: Failed to save roster file", "error")
                    self.update_status("Error")
                
            except Exception as e:
                self.log(f"Error: {str(e)}", "error")
                import traceback
                self.log(traceback.format_exc(), "error")
                self.update_status("Error")
        
        threading.Thread(target=process_thread, daemon=True).start()
    
    def process_zoom_attendance(self, zoom_file: str, date: Optional[str] = None, cut_time_minutes: int = 30):
        """Process Zoom attendance record"""
        def process_thread():
            try:
                self.update_status("Processing Zoom attendance...")
                self.log(f"Processing Zoom file: {zoom_file}")
                self.log(f"Cut time: {cut_time_minutes} minutes (≥{cut_time_minutes} min = 0.6 pts, <{cut_time_minutes} min = 0.2 pts)")
                
                # Import processing functions from app.py
                import sys
                import importlib.util
                
                app_path = os.path.join(os.getcwd(), 'app.py')
                if not os.path.exists(app_path):
                    self.log("Error: app.py not found. Cannot process Zoom attendance.", "error")
                    self.update_status("Error")
                    return
                
                spec = importlib.util.spec_from_file_location("app_module", app_path)
                app_module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(app_module)
                
                # Get processing functions
                find_student_in_roster = app_module.find_student_in_roster
                update_roster_with_attendance = app_module.update_roster_with_attendance
                format_date_for_roster = app_module.format_date_for_roster
                find_matching_date_column = app_module.find_matching_date_column
                parse_duration = app_module.parse_duration
                
                # Read Zoom file
                zoom_file_resolved = zoom_file.strip().strip('"\'')
                
                if zoom_file_resolved.lower().endswith('.csv'):
                    try:
                        zoom_df = pd.read_csv(zoom_file_resolved, encoding='utf-8')
                    except (UnicodeDecodeError, UnicodeError):
                        zoom_df = pd.read_csv(zoom_file_resolved, encoding='latin-1')
                else:
                    zoom_df = pd.read_excel(zoom_file_resolved, engine='openpyxl')
                
                self.log(f"Read {len(zoom_df)} rows from Zoom file", "info")
                self.log(f"Columns found: {list(zoom_df.columns)}", "info")
                
                # Find columns by header name
                name_col = None
                duration_col = None
                
                # Find Name column (could be "Name", "Name (original name)", etc.)
                for col in zoom_df.columns:
                    col_str = str(col).lower().strip()
                    if 'name' in col_str and 'guest' not in col_str:
                        name_col = col
                        break
                
                # Find Total duration column (could be "Total duration", "Total duration (minutes)", etc.)
                for col in zoom_df.columns:
                    col_str = str(col).lower().strip()
                    if 'duration' in col_str or 'time' in col_str:
                        duration_col = col
                        break
                
                if name_col is None:
                    self.log("Error: Could not find 'Name' column in Zoom file", "error")
                    self.log(f"Available columns: {list(zoom_df.columns)}", "error")
                    self.update_status("Error")
                    return
                
                if duration_col is None:
                    self.log("Error: Could not find 'Total duration' column in Zoom file", "error")
                    self.log(f"Available columns: {list(zoom_df.columns)}", "error")
                    self.update_status("Error")
                    return
                
                self.log(f"Using columns - Name: {name_col}, Duration: {duration_col}", "info")
                
                # Extract meeting date
                meeting_date = None
                
                # Try form date first
                if date:
                    try:
                        meeting_date = datetime.strptime(date, "%Y-%m-%d")
                    except:
                        pass
                
                if meeting_date is None:
                    meeting_date = datetime.now()
                
                # Format date for roster
                date_str = format_date_for_roster(meeting_date)
                
                # Try to find matching existing date column
                roster_df = self.roster_df.copy()
                matching_date_col = find_matching_date_column(roster_df, meeting_date)
                if matching_date_col:
                    date_str = matching_date_col
                    self.log(f"Found existing date column: {date_str}", "info")
                else:
                    self.log(f"Will create new date column: {date_str}", "info")
                
                # Remove header row if it's in the data (check if first row looks like a header)
                if len(zoom_df) > 0:
                    first_row_name = str(zoom_df.iloc[0].get(name_col, '')).lower().strip()
                    if first_row_name in ['name', 'name (original name)', 'participant']:
                        zoom_df = zoom_df.iloc[1:].reset_index(drop=True)
                        self.log("Removed header row from data", "info")
                
                self.log(f"Processing {len(zoom_df)} student records", "info")
                
                # Create date column if needed
                if date_str not in roster_df.columns:
                    roster_df[date_str] = 0.0
                
                processed_count = 0
                errors = []
                skipped_count = 0
                
                # Process each student
                for idx, row in zoom_df.iterrows():
                    # Get student name
                    name_val = row.get(name_col) if name_col else None
                    if pd.isna(name_val) or str(name_val).strip().lower() in ['nan', '', 'none', 'name', 'participant']:
                        skipped_count += 1
                        continue
                    
                    student_name = str(name_val).strip()
                    
                    # Skip header rows
                    name_lower = student_name.lower().strip()
                    if name_lower in ['name', 'participant', 'total', 'summary', 'meeting', 'zoom', 'class', 'attendance', 'report']:
                        skipped_count += 1
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
                    if duration_minutes >= cut_time_minutes:
                        points = 0.6  # Full attendance
                    elif duration_minutes > 0:
                        points = 0.2  # Partial attendance
                    else:
                        points = 0.0  # No attendance (skip or give 0)
                    
                    if points == 0.0:
                        continue  # Skip students with 0 points
                    
                    # Update roster
                    roster_df, found, confidence, matched_name = update_roster_with_attendance(
                        roster_df, student_name, points, date_str, use_gemini=False
                    )
                    
                    if found:
                        processed_count += 1
                        if confidence < 1.0:
                            self.log(f"✓ {student_name} → {matched_name} ({confidence:.1%} confidence) - {points} pts ({duration_minutes:.1f} min)", "info")
                    else:
                        errors.append(f"Could not match: {student_name} (duration: {duration_minutes:.1f} min, confidence: {confidence:.2f})")
                
                # Update roster in GUI
                self.roster_df = roster_df
                
                # Save roster
                if self._save_roster_internal(roster_df):
                    self.log(f"\n✓ Processed {processed_count} Zoom attendees successfully", "success")
                    if skipped_count > 0:
                        self.log(f"Skipped {skipped_count} rows (headers or invalid data)", "info")
                    if errors:
                        self.log(f"⚠ {len(errors)} unmatched students", "warning")
                        if len(errors) <= 10:
                            for err in errors:
                                self.log(f"  {err}", "warning")
                        else:
                            for err in errors[:10]:
                                self.log(f"  {err}", "warning")
                            self.log(f"  ... and {len(errors) - 10} more", "warning")
                    
                    self.update_roster_info()
                    self.update_status("Ready")
                    self.log("✓ Roster file updated successfully!", "success")
                else:
                    self.log("Error: Failed to save roster file", "error")
                    self.update_status("Error")
                
            except Exception as e:
                self.log(f"Error: {str(e)}", "error")
                import traceback
                self.log(traceback.format_exc(), "error")
                self.update_status("Error")
        
        threading.Thread(target=process_thread, daemon=True).start()
    
    def query_dialog(self):
        """Open dialog for query/view information"""
        if self.roster_df is None:
            messagebox.showwarning("Warning", "Please load a roster file first.")
            return
        
        if not self.gemini_model:
            messagebox.showerror("Error", "Gemini API not configured.")
            return
        
        # Create dialog window
        dialog = tk.Toplevel(self.root)
        dialog.title("Query/View Information")
        dialog.geometry("600x150")
        
        ttk.Label(dialog, text="Enter your query:").grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
        query_entry = ttk.Entry(dialog, width=70)
        query_entry.grid(row=0, column=1, padx=10, pady=10, sticky=(tk.W, tk.E))
        dialog.columnconfigure(1, weight=1)
        
        ttk.Label(dialog, text="Example: 'show late students for November 4'").grid(row=1, column=1, padx=10, sticky=tk.W)
        
        def query():
            query_text = query_entry.get().strip()
            if not query_text:
                messagebox.showwarning("Warning", "Please enter a query.")
                return
            
            dialog.destroy()
            self.query_with_gemini(query_text)
        
        ttk.Button(dialog, text="Query", command=query).grid(row=2, column=1, pady=20)
        ttk.Button(dialog, text="Cancel", command=dialog.destroy).grid(row=2, column=2, pady=20)
    
    def query_with_gemini(self, user_query: str):
        """Use Gemini API to understand user query and generate DSL code"""
        def query_thread():
            try:
                self.update_status("Processing query...")
                self.log(f"Query: {user_query}")
                
                # Get roster info
                roster_info = f"Roster has {len(self.roster_df)} students"
                date_columns = [str(col) for col in self.roster_df.columns 
                              if any(x in str(col).lower() for x in ['nov', 'oct', 'dec', 'jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', '.'])]
                
                # Use prompt template
                try:
                    from dsl.gemini_prompts import create_query_prompt
                    prompt = create_query_prompt(
                        user_query=user_query,
                        roster_info=roster_info,
                        date_columns=date_columns,
                        roster_file=self.roster_file or 'roster_attendance.xlsx'
                    )
                except ImportError:
                    prompt = f"User request: {user_query}\nGenerate DSL code to fulfill this request."
                
                # Call Gemini API
                self.log("Calling Gemini API...")
                response = self.gemini_model.generate_content(prompt)
                dsl_code = response.text.strip()
                
                # Clean up response
                if dsl_code.startswith('```'):
                    lines = dsl_code.split('\n')
                    dsl_code = '\n'.join(lines[1:-1]) if len(lines) > 2 else dsl_code
                
                self.log("\nGenerated DSL Code:", "info")
                self.log("-" * 60, "info")
                self.log(dsl_code, "info")
                self.log("-" * 60, "info")
                
                # Ask for confirmation
                self.root.after(0, lambda: self.confirm_and_execute_dsl(dsl_code, "Execute this DSL code?"))
                
            except Exception as e:
                self.log(f"Error: {str(e)}", "error")
                import traceback
                self.log(traceback.format_exc(), "error")
                self.update_status("Error")
        
        threading.Thread(target=query_thread, daemon=True).start()
    
    def find_student_dialog(self):
        """Open dialog to find student's total points"""
        if self.roster_df is None:
            messagebox.showwarning("Warning", "Please load a roster file first.")
            return
        
        # Create dialog window
        dialog = tk.Toplevel(self.root)
        dialog.title("Find Student's Total Points")
        dialog.geometry("500x120")
        
        ttk.Label(dialog, text="Student Name (Last Name, First Name):").grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
        name_entry = ttk.Entry(dialog, width=50)
        name_entry.grid(row=0, column=1, padx=10, pady=10)
        ttk.Label(dialog, text="Example: Smith, John").grid(row=1, column=1, padx=10, sticky=tk.W)
        
        def find():
            student_name = name_entry.get().strip()
            if not student_name:
                messagebox.showwarning("Warning", "Please enter a student name.")
                return
            
            dialog.destroy()
            self.find_student_points_with_gemini(student_name)
        
        ttk.Button(dialog, text="Find", command=find).grid(row=2, column=1, pady=20)
        ttk.Button(dialog, text="Cancel", command=dialog.destroy).grid(row=2, column=2, pady=20)
    
    def find_student_points_with_gemini(self, student_name: str):
        """Find a student's total points using Gemini API"""
        def find_thread():
            try:
                self.update_status("Finding student...")
                self.log(f"Searching for student: {student_name}")
                
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
                
                # Find student
                matching_students = []
                if name_col:
                    normalized_input = student_name.strip().replace(', ', ',').replace(' ,', ',')
                    
                    for idx, row in self.roster_df.iterrows():
                        roster_name = str(row[name_col]).strip()
                        if not roster_name or roster_name.lower() in ['nan', 'none', '']:
                            continue
                        
                        normalized_roster = roster_name.replace(', ', ',').replace(' ,', ',')
                        input_lower = normalized_input.lower()
                        roster_lower = normalized_roster.lower()
                        
                        if input_lower == roster_lower:
                            matching_students.insert(0, {
                                'index': idx,
                                'name': roster_name,
                                'total_points': self.roster_df.loc[idx, 'Total Points'] if 'Total Points' in self.roster_df.columns else None,
                            })
                        elif input_lower in roster_lower or roster_lower in input_lower:
                            matching_students.append({
                                'index': idx,
                                'name': roster_name,
                                'total_points': self.roster_df.loc[idx, 'Total Points'] if 'Total Points' in self.roster_df.columns else None,
                            })
                        else:
                            input_words = set(input_lower.replace(',', ' ').split())
                            roster_words = set(roster_lower.replace(',', ' ').split())
                            if len(input_words.intersection(roster_words)) >= 2:
                                matching_students.append({
                                    'index': idx,
                                    'name': roster_name,
                                    'total_points': self.roster_df.loc[idx, 'Total Points'] if 'Total Points' in self.roster_df.columns else None,
                                })
                
                # Display results - only show student name and total points
                if matching_students:
                    self.log(f"\n✓ Found {len(matching_students)} matching student(s):\n", "success")
                    for student in matching_students:
                        self.log(f"Student: {student['name']}", "info")
                        if student['total_points'] is not None:
                            self.log(f"Total Points: {student['total_points']}\n", "info")
                        else:
                            # Calculate total from date columns
                            total = 0.0
                            for col in date_columns:
                                val = self.roster_df.loc[student['index'], col]
                                if pd.notna(val):
                                    try:
                                        total += float(val)
                                    except (ValueError, TypeError):
                                        pass
                            self.log(f"Total Points: {total:.1f}\n", "info")
                else:
                    self.log(f"\n✗ No match found for '{student_name}'", "error")
                    self.log("Please check the spelling or format (Last Name, First Name)", "info")
                
                self.update_status("Ready")
                
            except Exception as e:
                self.log(f"Error: {str(e)}", "error")
                import traceback
                self.log(traceback.format_exc(), "error")
                self.update_status("Error")
        
        threading.Thread(target=find_thread, daemon=True).start()
    
    def execute_dsl_dialog(self):
        """Open dialog to execute DSL code manually"""
        if self.roster_df is None:
            messagebox.showwarning("Warning", "Please load a roster file first.")
            return
        
        # Create dialog window
        dialog = tk.Toplevel(self.root)
        dialog.title("Execute DSL Code")
        dialog.geometry("700x500")
        
        ttk.Label(dialog, text="Enter DSL code:").grid(row=0, column=0, padx=10, pady=10, sticky=tk.NW)
        
        dsl_text = scrolledtext.ScrolledText(dialog, wrap=tk.WORD, width=80, height=20, font=('Consolas', 10))
        dsl_text.grid(row=1, column=0, columnspan=2, padx=10, pady=10, sticky=(tk.W, tk.E, tk.N, tk.S))
        dialog.columnconfigure(0, weight=1)
        dialog.rowconfigure(1, weight=1)
        
        def execute():
            dsl_code = dsl_text.get(1.0, tk.END).strip()
            if not dsl_code:
                messagebox.showwarning("Warning", "Please enter DSL code.")
                return
            
            dialog.destroy()
            self.execute_dsl_code(dsl_code)
        
        ttk.Button(dialog, text="Execute", command=execute).grid(row=2, column=0, pady=10)
        ttk.Button(dialog, text="Cancel", command=dialog.destroy).grid(row=2, column=1, pady=10)
    
    def confirm_and_execute_dsl(self, dsl_code: str, message: str = "Execute this DSL code?"):
        """Show confirmation dialog and execute DSL code"""
        result = messagebox.askyesno("Confirm", message)
        if result:
            self.execute_dsl_code(dsl_code)
    
    def execute_dsl_code(self, dsl_code: str) -> bool:
        """Execute DSL code using the DSL executor"""
        if not self.dsl_executor:
            messagebox.showerror("Error", "DSL executor not available.")
            return False
        
        def execute_thread():
            try:
                self.update_status("Executing DSL code...")
                self.log("\nExecuting DSL code...", "info")
                self.log("-" * 60, "info")
                
                # Update executor's roster context
                self.dsl_executor.context['roster'] = self.roster_df
                self.dsl_executor.context['roster_file'] = self.roster_file
                
                result = self.dsl_executor.execute_script(dsl_code)
                
                if result['success']:
                    self.log("\n✓ DSL code executed successfully!", "success")
                    for output in result.get('output', []):
                        if output and output.get('result'):
                            self.log(f"  {output['result']}", "info")
                    
                    # Reload roster if it was modified
                    if self.roster_file:
                        self.load_roster(self.roster_file, show_message=False)
                        self.update_roster_info()
                    
                    self.update_status("Ready")
                else:
                    self.log(f"\n✗ Error executing DSL code: {result.get('error', 'Unknown error')}", "error")
                    if result.get('line_num'):
                        self.log(f"  Line {result.get('line_num')}: {result.get('line', '')}", "error")
                    self.update_status("Error")
                    
            except Exception as e:
                self.log(f"Error executing DSL code: {str(e)}", "error")
                import traceback
                self.log(traceback.format_exc(), "error")
                self.update_status("Error")
        
        threading.Thread(target=execute_thread, daemon=True).start()
        return True


def main():
    """Main entry point"""
    root = tk.Tk()
    app = AttendanceTrackerGUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()

