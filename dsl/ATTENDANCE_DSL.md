# Attendance Tracker Domain Specific Language (DSL)

## Overview
The Attendance Tracker DSL is a simple, concise language designed to automate attendance tracking operations. Each line contains exactly one command that corresponds to a function in the attendance tracker.

## Syntax Rules
- One command per line
- Commands are case-insensitive
- Whitespace is ignored (except within quoted strings)
- Comments start with `#` and continue to end of line
- String arguments can be in single or double quotes
- Commands can span multiple lines if quoted strings contain newlines

## Commands

### File Operations

#### `LOAD ROSTER <file_path>`
Loads a roster file (Excel or CSV) into the system.
- **Arguments:**
  - `file_path`: Path to the roster file (relative or absolute)
- **Example:** `LOAD ROSTER roster.xlsx`
- **Example:** `LOAD ROSTER "C:\Data\students.csv"`

#### `SAVE ROSTER [<file_path>]`
Saves the current roster to a file. If no path is specified, saves to default location.
- **Arguments:**
  - `file_path` (optional): Destination file path
- **Example:** `SAVE ROSTER`
- **Example:** `SAVE ROSTER updated_roster.xlsx`

#### `DOWNLOAD ROSTER [<file_path>]`
Downloads the current roster as an Excel file.
- **Arguments:**
  - `file_path` (optional): Destination file path
- **Example:** `DOWNLOAD ROSTER`

---

### Check-In Processing

#### `PROCESS CHECKIN <file_path> [DATE <date>] [EARLY_BIRD <time>] [REGULAR <time>]`
Processes check-in data from a Qualtrics export file.
- **Arguments:**
  - `file_path`: Path to Qualtrics export file (CSV or Excel)
  - `DATE` (optional): Meeting date in YYYY-MM-DD format (default: auto-detect from file)
  - `EARLY_BIRD` (optional): Start time for early bird check-in (default: 11:00)
  - `REGULAR` (optional): Start time for regular check-in (default: 11:36)
- **Example:** `PROCESS CHECKIN checkins.csv`
- **Example:** `PROCESS CHECKIN checkins.xlsx DATE 2025-11-04 EARLY_BIRD 11:00 REGULAR 11:36`

#### `SET CHECKIN TIMES EARLY_BIRD <time> REGULAR <time>`
Sets the default check-in time thresholds for point calculation.
- **Arguments:**
  - `EARLY_BIRD`: Time threshold for early bird (0.6 points)
  - `REGULAR`: Time threshold for regular (0.2 points)
- **Example:** `SET CHECKIN TIMES EARLY_BIRD 11:00 REGULAR 11:36`

---

### Zoom Attendance Processing

#### `PROCESS ZOOM <file_path> [DATE <date>]`
Processes Zoom meeting attendance from an Excel/CSV file.
- **Arguments:**
  - `file_path`: Path to Zoom attendance export file
  - `DATE` (optional): Meeting date in YYYY-MM-DD format (default: auto-detect from file)
- **Example:** `PROCESS ZOOM zoom_meeting.xlsx`
- **Example:** `PROCESS ZOOM zoom_oct23.csv DATE 2025-10-23`

---

### Roster Management

#### `VIEW ROSTER`
Displays the current roster with all attendance data.
- **Example:** `VIEW ROSTER`

#### `DELETE DATE <date_column>`
Deletes a specific date column and its attendance data from the roster.
- **Arguments:**
  - `date_column`: Name of the date column to delete (e.g., "R,Oct.23" or "11.4")
- **Example:** `DELETE DATE "R,Oct.23"`
- **Example:** `DELETE DATE 11.4`

#### `SHOW LATE STUDENTS DATE <date>`
Shows the names of students who were late to class (received 0.2 points) for a given date.
- **Arguments:**
  - `DATE`: Meeting date in YYYY-MM-DD, MM/DD/YYYY, or column name format (e.g., "R,Oct.23", "11.4")
- **Example:** `SHOW LATE STUDENTS DATE 2025-11-04`
- **Example:** `SHOW LATE STUDENTS DATE "R,Oct.23"`
- **Example:** `SHOW LATE STUDENTS DATE 11.4`
- **Returns:** A numbered list of student names who received 0.2 points for that date

#### `SHOW EARLY STUDENTS DATE <date>`
Shows the names of students who were on-time/early (received 0.6 points) for a given date.
- **Arguments:**
  - `DATE`: Meeting date in YYYY-MM-DD, MM/DD/YYYY, or column name format (e.g., "R,Oct.23", "11.4")
- **Example:** `SHOW EARLY STUDENTS DATE 2025-11-04`
- **Example:** `SHOW EARLY STUDENTS DATE "R,Oct.23"`
- **Example:** `SHOW EARLY STUDENTS DATE 11.4`
- **Returns:** A numbered list of student names who received 0.6 points for that date

#### `SHOW STUDENT TOTAL <student_name>` or `FIND STUDENT <student_name>`
Shows the total attendance points for a specific student by searching the roster.
- **Arguments:**
  - `student_name`: Student's name (can be in various formats: "Last, First", "First Last", etc.)
- **Example:** `SHOW STUDENT TOTAL "Marco Acosta"`
- **Example:** `SHOW STUDENT TOTAL Acosta, Marco`
- **Example:** `FIND STUDENT "Marco Acosta"`
- **Returns:** The student's name (as found in roster) and their total attendance points
- **Note:** The command searches the roster using flexible name matching, so it handles name format variations and partial matches.

---

### Settings

#### `ENABLE GEMINI`
Enables AI-assisted name matching using Gemini API.
- **Example:** `ENABLE GEMINI`

#### `DISABLE GEMINI`
Disables AI-assisted name matching.
- **Example:** `DISABLE GEMINI`

#### `SET GEMINI KEY <api_key>`
Sets the Gemini API key for AI-assisted matching.
- **Arguments:**
  - `api_key`: Your Gemini API key
- **Example:** `SET GEMINI KEY "your-api-key-here"`

---

### QR Code Generation

#### `GENERATE QR <url> [OUTPUT <file_path>]`
Generates a QR code linking to the specified URL.
- **Arguments:**
  - `url`: URL to encode in the QR code
  - `OUTPUT` (optional): File path to save QR code image (default: displays in console)
- **Example:** `GENERATE QR "https://forms.qualtrics.com/..." OUTPUT qr_code.png`

---

### Utility Commands

#### `ECHO <message>`
Prints a message to the console.
- **Arguments:**
  - `message`: Message to print (can be quoted or unquoted)
- **Example:** `ECHO Processing attendance...`
- **Example:** `ECHO "Starting batch processing"`

#### `WAIT <seconds>`
Pauses execution for the specified number of seconds.
- **Arguments:**
  - `seconds`: Number of seconds to wait (can be decimal)
- **Example:** `WAIT 2`
- **Example:** `WAIT 0.5`

#### `RUN <script_path>`
Executes another DSL script file.
- **Arguments:**
  - `script_path`: Path to another DSL script file
- **Example:** `RUN "scripts\daily_attendance.dsl"`

---

### Batch Processing

#### `BEGIN BATCH`
Starts a batch operation (collects all commands until END BATCH).
- **Example:** `BEGIN BATCH`

#### `END BATCH`
Ends a batch operation and executes all collected commands.
- **Example:** `END BATCH`

---

## File Format
DSL scripts are saved as text files with the extension `.dsl` (Attendance Tracker Language).

## Example Scripts

### Basic Check-In Processing
```dsl
# Load the roster
LOAD ROSTER roster.xlsx

# Process check-ins for November 4
PROCESS CHECKIN "checkins_2025-11-04.csv" DATE 2025-11-04

# Save updated roster
SAVE ROSTER

# View the results
VIEW ROSTER
```

### Weekly Attendance Processing
```dsl
# Setup
LOAD ROSTER roster.xlsx
SET CHECKIN TIMES EARLY_BIRD 11:00 REGULAR 11:36

# Process Monday check-ins
PROCESS CHECKIN "week1_monday.csv" DATE 2025-11-03

# Process Wednesday check-ins
PROCESS CHECKIN "week1_wednesday.csv" DATE 2025-11-05

# Process Friday Zoom meeting
PROCESS ZOOM "week1_friday_zoom.xlsx" DATE 2025-11-07

# Finalize
SAVE ROSTER
DOWNLOAD ROSTER "week1_final_roster.xlsx"
```

### Automated Daily Workflow
```dsl
# Daily attendance processing script
ECHO "Starting daily attendance processing..."

# Load roster
LOAD ROSTER roster.xlsx

# Check today's date and process accordingly
PROCESS CHECKIN "today_checkins.csv"

# Save and backup
SAVE ROSTER
DOWNLOAD ROSTER "backups/roster_backup_$(date).xlsx"

ECHO "Processing complete!"
```

## Error Handling
- If a command fails, execution stops and an error message is displayed
- Use `#` comments to disable commands during debugging
- All file paths are validated before processing
- Missing files generate clear error messages

## Best Practices
1. Always load the roster at the start of a script
2. Save the roster after making changes
3. Use quotes for file paths containing spaces
4. Add comments to document your workflow
5. Test scripts with sample data first
6. Keep backup copies of roster files
