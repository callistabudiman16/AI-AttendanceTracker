# Attendance Tracker DSL (Domain Specific Language)

## Overview

The Attendance Tracker DSL is a simple, concise scripting language designed to automate attendance tracking operations. Each line contains exactly one command that corresponds to a function in the attendance tracker application.

## Quick Start

### 1. Access the DSL Interface

1. Start the Flask app: `python app.py`
2. Navigate to the **DSL Scripts** page in the web interface
3. Enter your DSL commands in the text area
4. Click **Execute Script** to run

### 2. Create a DSL Script File

Create a file with the `.dsl` extension (Attendance Tracker Language):

```dsl
# Example script: daily_checkin.dsl
LOAD ROSTER roster_attendance.xlsx
PROCESS CHECKIN "uploads/today_checkins.csv" DATE 2025-11-04
SAVE ROSTER
VIEW ROSTER
```

### 3. Execute via Command Line

```bash
python dsl_executor.py examples/daily_checkin.dsl
```

## Command Reference

### File Operations

| Command | Description | Example |
|---------|-------------|---------|
| `LOAD ROSTER <file>` | Loads a roster file | `LOAD ROSTER roster.xlsx` |
| `SAVE ROSTER [<file>]` | Saves the current roster | `SAVE ROSTER` |
| `DOWNLOAD ROSTER [<file>]` | Downloads roster as Excel | `DOWNLOAD ROSTER` |

### Check-In Processing

| Command | Description | Example |
|---------|-------------|---------|
| `PROCESS CHECKIN <file> [DATE <date>] [EARLY_BIRD <time>] [REGULAR <time>]` | Processes Qualtrics check-in file | `PROCESS CHECKIN checkins.csv DATE 2025-11-04` |
| `SET CHECKIN TIMES EARLY_BIRD <time> REGULAR <time>` | Sets check-in time thresholds | `SET CHECKIN TIMES EARLY_BIRD 11:00 REGULAR 11:36` |

### Zoom Attendance

| Command | Description | Example |
|---------|-------------|---------|
| `PROCESS ZOOM <file> [DATE <date>]` | Processes Zoom attendance file | `PROCESS ZOOM zoom_meeting.xlsx DATE 2025-10-23` |

### Roster Management

| Command | Description | Example |
|---------|-------------|---------|
| `VIEW ROSTER` | Displays current roster | `VIEW ROSTER` |
| `DELETE DATE <column>` | Deletes a date column | `DELETE DATE "R,Oct.23"` |
| `SHOW LATE STUDENTS DATE <date>` | Shows students who got 0.2 points (late) for a date | `SHOW LATE STUDENTS DATE 11.4` |
| `SHOW EARLY STUDENTS DATE <date>` | Shows students who got 0.6 points (on-time/early) for a date | `SHOW EARLY STUDENTS DATE "T,Nov.4"` |
| `SHOW STUDENT TOTAL <name>` | Shows total attendance points for a specific student | `SHOW STUDENT TOTAL "Marco Acosta"` |
| `FIND STUDENT <name>` | Alias for SHOW STUDENT TOTAL | `FIND STUDENT "Marco Acosta"` |

### Settings

| Command | Description | Example |
|---------|-------------|---------|
| `ENABLE GEMINI` | Enables AI-assisted matching | `ENABLE GEMINI` |
| `DISABLE GEMINI` | Disables AI-assisted matching | `DISABLE GEMINI` |
| `SET GEMINI KEY <key>` | Sets Gemini API key | `SET GEMINI KEY "your-key"` |

### Utilities

| Command | Description | Example |
|---------|-------------|---------|
| `ECHO <message>` | Prints a message | `ECHO "Processing..."` |
| `WAIT <seconds>` | Pauses execution | `WAIT 2` |
| `GENERATE QR <url> [OUTPUT <file>]` | Generates QR code | `GENERATE QR "https://..." OUTPUT qr.png` |

## Example Scripts

### Daily Check-In Workflow

```dsl
# Load roster
LOAD ROSTER roster_attendance.xlsx

# Configure settings
SET CHECKIN TIMES EARLY_BIRD 11:00 REGULAR 11:36

# Process check-ins
PROCESS CHECKIN "uploads/today_checkins.csv" DATE 2025-11-04

# Save and view
SAVE ROSTER
VIEW ROSTER
```

### Weekly Attendance Processing

```dsl
# Setup
LOAD ROSTER roster.xlsx
SET CHECKIN TIMES EARLY_BIRD 11:00 REGULAR 11:36

# Process Monday
ECHO "Processing Monday..."
PROCESS CHECKIN "week1_monday.csv" DATE 2025-11-03

# Process Wednesday
ECHO "Processing Wednesday..."
PROCESS CHECKIN "week1_wednesday.csv" DATE 2025-11-05

# Process Friday Zoom
ECHO "Processing Friday Zoom..."
PROCESS ZOOM "week1_friday_zoom.xlsx" DATE 2025-11-07

# Save and backup
SAVE ROSTER
DOWNLOAD ROSTER "backups/week1_final.xlsx"
```

### Batch Operations

```dsl
# Batch mode allows grouping commands
BEGIN BATCH
LOAD ROSTER roster.xlsx
SET CHECKIN TIMES EARLY_BIRD 11:00 REGULAR 11:36
PROCESS CHECKIN "monday.csv" DATE 2025-11-03
PROCESS CHECKIN "wednesday.csv" DATE 2025-11-05
SAVE ROSTER
END BATCH
```

## Syntax Rules

1. **One command per line** - Each line contains exactly one command
2. **Case-insensitive** - Commands can be uppercase, lowercase, or mixed case
3. **Whitespace ignored** - Extra spaces are ignored (except within quoted strings)
4. **Comments** - Lines starting with `#` are comments
5. **Quoted strings** - Use quotes for file paths with spaces: `"C:\My Files\roster.xlsx"`
6. **Optional arguments** - Optional arguments are in square brackets: `[DATE <date>]`

## Error Handling

- If a command fails, execution stops
- Error messages show the line number and command that failed
- Use `#` to comment out commands during debugging
- File paths are validated before processing

## Files

- **`ATTENDANCE_DSL.md`** - Complete DSL specification and documentation
- **`dsl_executor.py`** - Standalone DSL executor (command-line)
- **`dsl_integrated.py`** - Integrated DSL executor (Flask app)
- **`examples/`** - Example DSL scripts
- **`templates/dsl.html`** - Web interface for DSL execution

## Integration

The DSL executor is integrated into the Flask app:

1. Navigate to `/dsl` in the web interface
2. Enter DSL commands in the text area
3. Click "Execute Script" to run
4. View execution results in the output section

## Best Practices

1. **Always load roster first** - Start scripts with `LOAD ROSTER`
2. **Save after changes** - Use `SAVE ROSTER` after processing
3. **Use quotes for paths** - Quote file paths with spaces: `"my file.csv"`
4. **Add comments** - Document your workflow with `#` comments
5. **Test with samples** - Test scripts with sample data first
6. **Keep backups** - Always backup roster files before batch operations

## Troubleshooting

### Script won't execute
- Check that commands are spelled correctly
- Verify file paths exist
- Ensure roster is loaded before processing

### File not found errors
- Use absolute paths or paths relative to script location
- Check file permissions
- Verify file extensions (.csv, .xlsx)

### Command not recognized
- Check command spelling (case-insensitive but must match exactly)
- See `ATTENDANCE_DSL.md` for complete command list
- Ensure proper spacing between command and arguments

## Future Enhancements

- Variable support
- Conditional statements (IF/ELSE)
- Loops for batch processing
- Function definitions
- Import/export capabilities

