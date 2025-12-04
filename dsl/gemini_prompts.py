"""
Gemini API Prompt Templates for DSL Code Generation

This module contains prompt templates for generating DSL code using Google's Gemini API.
"""

from typing import Optional

def get_dsl_specification():
    """Load DSL specification from file"""
    try:
        with open('dsl/ATTENDANCE_DSL.md', 'r', encoding='utf-8') as f:
            return f.read()
    except:
        return "DSL specification file not found."


def create_attendance_processing_prompt(roster_sample: str, attendance_sample: str, 
                                       attendance_file: str, date: Optional[str] = None,
                                       roster_file: str = 'roster_attendance.xlsx',
                                       class_type: str = "in_person",
                                       start_time: Optional[str] = None,
                                       end_time: Optional[str] = None,
                                       cut_time: Optional[int] = None):
    """Create prompt for processing attendance records"""
    dsl_spec = get_dsl_specification()
    
    # Build time context based on class type
    time_context = ""
    if class_type == "in_person":
        time_context = f"""
TIME SETTINGS FOR IN-PERSON CLASS:
- Start Time: {start_time or '11:00'}
- End Time: {end_time or '11:35'}
- Points Assignment:
  * 0.6 points: Students who checked in between start time and end time (inclusive)
  * 0.2 points: Students who checked in after end time
  * Early check-ins (before start time) also get 0.6 points"""
    elif class_type == "zoom":
        time_context = f"""
TIME SETTINGS FOR ZOOM MEETING:
- Cut Time: {cut_time or 30} minutes
- Points Assignment:
  * 0.6 points: Students who stayed in the meeting for ≥{cut_time or 30} minutes
  * 0.2 points: Students who stayed in the meeting for <{cut_time or 30} minutes"""
    
    prompt = f"""You are a Teaching Assistant (TA) responsible for manually entering student attendance points into the roster for a large class with over 180 students.

YOUR ROLE AND CONTEXT:
- You are processing attendance records from a class with 180+ students
- Students check in using Qualtrics for in-person classes
- Students often make typos when typing their names
- The roster contains student names in "Last Name, First Name" format (e.g., "Smith, John Michael")
- For in-person Qualtrics check-ins: Students are instructed to type their name in "Last Name, First Name" format, but they often:
  * Include their middle name incorrectly
  * Make spelling mistakes
  * Use different name formats (e.g., "John Smith" instead of "Smith, John")
  * Miss commas or add extra spaces
- For Zoom attendance records: Names are typically in "First Name Last Name" format (e.g., "John Smith")
- You must carefully match attendance names to roster names, handling all these variations

DSL SPECIFICATION:
{dsl_spec[:5000]}

CURRENT ATTENDANCE FILE TO PROCESS:
- File Path: {attendance_file}
- Class Type: {class_type.upper()}
{time_context}
- Meeting Date: {date or 'auto-detect from file'}
- Roster File: {roster_file} (already loaded)

ROSTER SAMPLE (first 10 rows - note the name format):
{roster_sample}

ATTENDANCE RECORD SAMPLE (first 10 rows - observe name variations):
{attendance_sample}

YOUR TASK:
As a TA, you need to generate DSL code that will:

1. Process the attendance file: {attendance_file}
   - The system will automatically match student names from attendance to roster
   - Handle name format differences (Last,First vs First Last)
   - Handle typos and variations (e.g., "Jon" vs "John", "Smith" vs "Smyth")
   - Handle middle name variations (e.g., "John M Smith" vs "Smith, John Michael")
   - Match partial names when full names don't match exactly

2. Assign points based on attendance:
{class_type}_points_instructions

3. Update the roster with attendance points for the date: {date or 'auto-detect'}

4. Save the updated roster

CRITICAL INSTRUCTIONS FOR NAME MATCHING:
- Roster format: "Last Name, First Name Middle Name" (e.g., "Smith, John Michael")
- Attendance may have: "John Smith", "Smith John", "Smith, John", "Jon Smith", "John M. Smith", etc.
- The system handles matching automatically, but you should be aware of these variations
- If a name cannot be matched, it will be flagged for manual review

DSL CODE GENERATION RULES:
- Use exact DSL command syntax from the specification
- One command per line
- Do NOT include LOAD ROSTER command (roster is already loaded)
- Use PROCESS CHECKIN command with the attendance file path
- If date is provided, use DATE parameter
- After processing, ensure the roster is saved

Generate ONLY the DSL code. Do not include explanations, markdown code blocks, or comments outside of the DSL syntax. Return pure DSL commands:
"""
    
    # Add class-type specific point instructions
    if class_type == "in_person":
        prompt = prompt.replace("{class_type}_points_instructions", f"""
   - 0.6 points: Students who checked in between {start_time or '11:00'} and {end_time or '11:35'} (inclusive)
   - 0.6 points: Students who checked in before {start_time or '11:00'} (early bird)
   - 0.2 points: Students who checked in after {end_time or '11:35'} (late)""")
    elif class_type == "zoom":
        prompt = prompt.replace("{class_type}_points_instructions", f"""
   - 0.6 points: Students who stayed in Zoom meeting for ≥{cut_time or 30} minutes
   - 0.2 points: Students who stayed in Zoom meeting for <{cut_time or 30} minutes""")
    else:
        prompt = prompt.replace("{class_type}_points_instructions", """
   - 0.6 points: Full attendance
   - 0.2 points: Partial attendance""")
    
    return prompt


def create_query_prompt(user_query: str, roster_info: str, date_columns: list, roster_file: str):
    """Create prompt for understanding user queries and generating DSL code"""
    dsl_spec = get_dsl_specification()
    
    date_cols_str = ', '.join(date_columns[:15]) if date_columns else 'None'
    if len(date_columns) > 15:
        date_cols_str += f', ... and {len(date_columns) - 15} more'
    
    prompt = f"""You are a Teaching Assistant (TA) managing attendance for a large class with over 180 students.

YOUR ROLE:
- You are helping process and query attendance data for a class
- The roster contains student names in "Last Name, First Name" format
- Students often have typos in their names when checking in via Qualtrics
- You need to generate DSL code to answer questions about attendance

DSL SPECIFICATION:
{dsl_spec[:5000]}

CURRENT SYSTEM STATE:
{roster_info}
- Available date columns: {date_cols_str}
- Roster file: {roster_file}
- Roster format: Student names are in "Last Name, First Name" format (e.g., "Smith, John Michael")

USER REQUEST/QUESTION:
{user_query}

YOUR TASK:
As a TA, analyze the user's request and generate appropriate DSL code to fulfill it. Think step by step:

1. Understand what the user is asking for
2. Identify which DSL commands can accomplish this
3. Match any date references to the exact column names in the roster
4. Generate the DSL code

COMMON REQUESTS AND THEIR DSL EQUIVALENTS:
- "Show late students for [date]" → SHOW LATE STUDENTS DATE [date]
- "Show early/on-time students for [date]" → SHOW EARLY STUDENTS DATE [date]
- "View roster" → VIEW ROSTER
- "Calculate total points" → Use appropriate commands to calculate and save totals
- "Delete date column" → DELETE DATE [column_name]
- "Process check-in" → PROCESS CHECKIN [file] DATE [date]
- "Show attendance for [date]" → View roster or filter by date column
- "Find student [name]" → SHOW STUDENT TOTAL [name] or FIND STUDENT [name]
- "Show total points for [student name]" → SHOW STUDENT TOTAL [name]
- "What are [student name]'s total points?" → SHOW STUDENT TOTAL [name]

DATE FORMAT MATCHING:
- Date columns in roster use formats like: "T,Nov.4", "R,Oct.23", "Nov.4", "10.23"
- When user mentions a date, match it to the closest column name format
- If user says "November 4" or "Nov 4", look for columns like "T,Nov.4" or "Nov.4"
- If user says "October 23" or "Oct 23", look for columns like "R,Oct.23" or "Oct.23"

CRITICAL INSTRUCTIONS:
- Use exact DSL command syntax from the specification
- One command per line
- Do NOT include LOAD ROSTER command (roster is already loaded)
- Match date formats to existing column names exactly
- For date queries, use the exact column name format found in the roster
- If the user wants to see information, use appropriate DSL commands
- Think step by step: What does the user want? → Which DSL command does this? → Generate the code

Generate ONLY the DSL code. Do not include explanations, markdown code blocks, or comments outside of the DSL syntax. Return pure DSL commands:
"""
    return prompt


def create_find_student_prompt(student_name: str, roster_info: str, roster_sample: str, date_columns: list):
    """Create prompt for finding a specific student's information"""
    dsl_spec = get_dsl_specification()
    
    date_cols_str = ', '.join(date_columns[:15]) if date_columns else 'None'
    if len(date_columns) > 15:
        date_cols_str += f', ... and {len(date_columns) - 15} more'
    
    prompt = f"""You are a Teaching Assistant (TA) looking up a specific student's attendance information in a class with over 180 students.

YOUR ROLE AND CONTEXT:
- You are searching for a student in a large roster
- The roster contains student names in "Last Name, First Name" format (e.g., "Smith, John Michael")
- Students often have typos in their names when checking in
- You need to find the student even if there are name format variations

DSL SPECIFICATION:
{dsl_spec[:5000]}

CURRENT SYSTEM STATE:
{roster_info}
- Available date columns: {date_cols_str}
- Roster format: Student names are in "Last Name, First Name" format

STUDENT TO FIND:
User is searching for: "{student_name}"

ROSTER SAMPLE (observe the name format):
{roster_sample}

YOUR TASK:
As a TA, you need to find the student "{student_name}" and retrieve their:
1. Full name (as it appears in the roster)
2. Total points
3. Attendance points for each date column

NAME MATCHING CONSIDERATIONS:
- User input format: "{student_name}" (may be "Last, First" or "First Last")
- Roster format: "Last Name, First Name Middle Name" (e.g., "Smith, John Michael")
- Possible variations:
  * "{student_name}" might match "Smith, John" if user entered "Smith, John"
  * "{student_name}" might match "Smith, John Michael" if user entered "Smith, John"
  * "{student_name}" might have typos (e.g., "Jon" instead of "John", "Smyth" instead of "Smith")
  * Middle names might be included or omitted
  * Name order might be different

IMPORTANT INSTRUCTIONS:
- The roster is already loaded
- Student names in roster are in "Last Name, First Name" format (e.g., "Smith, John Michael")
- User input "{student_name}" may be in various formats
- You need to match the student name even if there are slight variations or typos
- The system will handle name matching automatically, but you should be aware of format differences

NOTE: Since DSL doesn't have a direct "FIND STUDENT" command, the system will perform direct lookup. However, for DSL code generation purposes:
- The system can directly search for the student
- You can suggest using VIEW ROSTER if needed
- Focus on generating code that helps display the student's information

Generate DSL code that helps find and display the student's information. If direct lookup is not available in DSL, suggest using VIEW ROSTER or appropriate commands:
"""
    return prompt


def create_name_matching_prompt(student_names_attendance: list, student_names_roster: list):
    """Create prompt for matching student names between attendance and roster"""
    prompt = f"""You are an expert in matching student names, handling typos and variations.

ATTENDANCE RECORD NAMES (from check-in):
{chr(10).join(f"- {name}" for name in student_names_attendance[:50])}

ROSTER NAMES (from full roster):
{chr(10).join(f"- {name}" for name in student_names_roster[:100])}

TASK:
Match each attendance record name to the best matching roster name, handling:
- Format differences (e.g., "John Smith" vs "Smith,John")
- Typos (e.g., "Jon Smith" vs "John Smith")
- Middle names/initials (e.g., "John M Smith" vs "Smith,John Michael")
- Name order variations (e.g., "Smith John" vs "John Smith")

For each attendance name, provide:
1. The exact roster name that matches

Return as JSON format:
{{
  "matches": [
    {{"attendance_name": "Name from attendance", "roster_name": "Matched roster name", "confidence": "High/Medium/Low"}},
    ...
  ]
}}

Generate ONLY the JSON response:
"""
    return prompt

