# Attendance Tracker App

A comprehensive attendance tracking application that handles both in-person QR code check-ins and Zoom meeting attendance.

## Features

### 📸 In-Person Check-In
- QR code generation for student check-in
- Manual entry option for student names (format: "Last Name, First Name")
- Automatic point calculation:
  - **0.6 points** for on-time attendance
  - **0.2 points** for late attendance (configurable threshold)
- Real-time roster updates

### 💻 Zoom Meeting Attendance
- Process Zoom attendance reports from Excel files
- Automatic duration-based point calculation:
  - **0.6 points** for 30+ minutes of participation
  - **0.2 points** for less than 30 minutes
  - **0.0 points** for no attendance
- Handles name format differences between in-person and Zoom data

### 📊 Roster Management
- View and export updated roster with attendance points
- Automatic total points calculation
- Export to Excel format

## Installation

1. Install Python 3.8 or higher

2. Install required packages:
```bash
pip install -r requirements.txt
```

## Usage

1. **Start the application:**
```bash
streamlit run app.py
```

**Note:** This app is designed for Professors and Teaching Assistants only. Students do not have access to this application. Students check in by scanning a QR code that links to a Qualtrics form, and the professor/TA exports the Qualtrics responses and imports that data into this app.

2. **Upload Roster File (First Time Only):**
   - Click on the sidebar
   - Upload your student roster file (Excel or CSV format)
   - Ensure the roster has a column with student names
   - The app automatically saves this file and loads it every time you open the app
   - **Note:** After the first upload, the roster is automatically loaded - no need to upload again!

3. **For In-Person Classes:**
   - Go to "📸 In-Person Check-In" tab
   - Generate a QR code that links to your Qualtrics check-in form
   - Display the QR code for students to scan and check in
   - After students check in, export responses from Qualtrics (CSV/Excel)
   - Upload the exported Qualtrics file in the app
   - Or use manual entry to add individual student check-ins
   - Process the check-ins to update the roster with attendance points
   - Configure late threshold and class start time in the sidebar

4. **For Zoom Classes:**
   - Go to "💻 Zoom Attendance" tab
   - Upload the Zoom meeting report (Excel file)
   - The app will automatically:
     - Find name and duration columns
     - Match students to roster (handles "First Last" format from Zoom)
     - Calculate points based on participation duration
   - Select the meeting date and click "Process Zoom Attendance"

5. **View and Export:**
   - Go to "📊 View Roster" tab
   - View the updated roster with all attendance points
   - Download the updated roster as an Excel file

## Name Format Handling

The app automatically handles different name formats:
- **In-person check-in**: "Last Name, First Name" (e.g., "Budiman, Natasha")
- **Zoom Excel files**: "First Name Last Name" (e.g., "Natasha Budiman")
- The app matches these formats to find students in the roster

## File Formats

### Roster File
- Supports Excel (.xlsx, .xls) and CSV formats
- Should contain a column with student names
- Attendance points will be added as new columns (one per date)

### Zoom Attendance File
- Should be an Excel file exported from Zoom
- Must contain:
  - A column with participant names (will be auto-detected)
  - A column with meeting duration (will be auto-detected)
- Duration can be in formats like:
  - "1:30:45" (hours:minutes:seconds)
  - "90:30" (minutes:seconds)
  - "90" (minutes)

## Configuration

### Late Threshold
Set the number of minutes after class start when a student is considered late (default: 15 minutes).

### Class Start Time
Set the time when the class starts (used to determine if check-ins are late).

## Troubleshooting

- **Student not found in roster**: Check that the name format matches. The app tries to match variations, but ensure names are consistent.
- **Zoom file not processing**: Ensure the Zoom Excel file has recognizable "name" and "duration" columns.
- **Webcam not working**: QR code scanning via webcam requires additional permissions. Use manual entry as an alternative.

## Notes

- **Persistent Roster File:** The app automatically saves your roster to `roster_attendance.xlsx` and loads it each time you open the app
- **No Re-uploading Required:** Once you upload your roster the first time, it's automatically loaded on subsequent uses
- **All Meetings in One File:** All attendance dates (e.g., 2024-10-10, 2024-10-12, 2024-10-17) are stored in the same roster file
- Each date gets its own column for attendance points
- The app preserves all existing roster data while adding attendance columns
- Total points are calculated automatically across all attendance dates
- The roster is automatically saved after each attendance update (check-ins or Zoom processing)
