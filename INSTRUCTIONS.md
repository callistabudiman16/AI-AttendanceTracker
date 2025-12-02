# Instructions for Professors and Teaching Assistants

## Overview

This attendance tracker app helps you manage student attendance for both in-person and online (Zoom) classes. The app updates your Excel roster file with attendance points automatically.

**Important:** This app is for Professors and Teaching Assistants only. Students do not have access to this application.

## Getting Started

### 1. First-Time Setup

1. **Upload your roster file** in the sidebar (first time only)
   - Supports Excel (.xlsx, .xls) or CSV files
   - Must contain a column with student names
   - Can include other columns (ID, Email, etc.) - they will be preserved
   - The app automatically saves this file as `roster_attendance.xlsx`
   - **After the first upload, the roster loads automatically - no need to upload again!**

2. **Configure class settings** in the sidebar:
   - **Late Threshold**: Minutes after class start when a student is considered late (default: 15 minutes)
   - **Class Start Time**: The time your class begins

### 2. Setting Up Student Check-In

For in-person classes, you have two options:

#### Option A: Qualtrics Form (Recommended)
1. Create a check-in form in Qualtrics:
   - Set up a text field where students enter their name
   - Format: **"Last Name, First Name"** (e.g., "Budiman, Natasha")
   - Qualtrics automatically collects timestamps for responses

2. In the app:
   - Go to "📸 In-Person Check-In" tab
   - Enter your Qualtrics form URL in the "QR Code URL" field
   - Display the generated QR code in class for students to scan

#### Option B: Manual Entry
- Use the manual entry option in the app to add individual student check-ins as needed

## Using the App

### For In-Person Classes

1. **Generate and Display QR Code**
   - In "📸 In-Person Check-In" tab, enter your check-in form URL
   - QR code is generated automatically
   - Display it on screen or print it for students

2. **After Class: Import Check-In Data**
   - In Qualtrics, go to Data & Analysis → Export & Import → Export Data
   - Export responses as CSV or Excel file
   - In the app, click "Upload Qualtrics Export File"
   - Select the exported file from Qualtrics

3. **Process Check-Ins**
   - Review the preview of check-ins
   - Click "Process Check-Ins and Update Roster"
   - The app will:
     - Match students to the roster (handles name format variations)
     - Calculate points based on check-in time (0.6 for on-time, 0.2 for late)
     - Update the roster file

4. **Manual Entry (if needed)**
   - Enter student name: "Last Name, First Name"
   - Click "Add Check-In"
   - Points are calculated automatically based on current time

### For Zoom Classes

1. **After the Zoom Meeting**
   - Export the Zoom meeting report as an Excel file
   - Go to Zoom Reports → Usage Reports → Meeting Participants

2. **Process Zoom Attendance**
   - Go to "💻 Zoom Attendance" tab
   - Click "Upload Zoom Attendance Excel File"
   - Select the exported Zoom report
   - Select the meeting date
   - Click "Process Zoom Attendance"

3. **Points Calculation**
   - The app automatically:
     - Finds participant names and durations
     - Calculates points:
       - **0.6 points** for 30+ minutes of participation
       - **0.2 points** for less than 30 minutes
       - **0.0 points** for no attendance
     - Matches names to roster (handles "First Last" format from Zoom)
     - Updates the roster file

### View and Export Updated Roster

1. Go to "📊 View Roster" tab
2. View the updated roster with:
   - All attendance columns (one per date)
   - Total points calculated automatically
3. Click "📥 Download Updated Roster" to save as Excel file

## Point System

### In-Person Attendance
- **0.6 points**: Student checked in on time (within the late threshold)
- **0.2 points**: Student checked in late (after the late threshold)

### Zoom Attendance
- **0.6 points**: Student participated for 30+ minutes
- **0.2 points**: Student participated for less than 30 minutes
- **0.0 points**: Student did not attend

## Name Format Handling

The app automatically handles different name formats:

- **In-person check-ins**: "Last Name, First Name" (e.g., "Budiman, Natasha")
- **Zoom Excel files**: "First Name Last Name" (e.g., "Natasha Budiman")
- **Roster file**: Can be either format - the app will match automatically

## Tips

- **Persistent Roster:** Your roster is automatically saved and loaded each time you use the app
- **All Meetings in One File:** All class dates (e.g., 2024-10-10, 2024-10-12, 2024-10-17, etc.) are stored in the same roster file
- Each class date gets its own column in the roster
- Total points are calculated automatically across all dates
- The roster is automatically saved after each attendance update - no manual saving needed!
- You can replace the roster file in the sidebar if needed (existing attendance data will be preserved if names match)
- The app handles duplicate check-ins (keeps the higher point value)
- Check-in data files are saved as `checkins_YYYY-MM-DD.csv` for your records

## Troubleshooting

**Student not found in roster:**
- Check that the name format matches (app tries to match variations automatically)
- Ensure names are spelled correctly

**Zoom file not processing:**
- Ensure the Excel file has columns with "name" and "duration" keywords
- Check that duration is in a recognizable format (HH:MM:SS, MM:SS, or minutes)

**Check-in file not loading:**
- Ensure the file has a "Name" column
- Optional: Include a "Time" column for accurate late detection

## Exporting and Saving

- **Automatic Saving:** The roster is automatically saved to `roster_attendance.xlsx` after each attendance update
- **No Manual Saving Needed:** The file persists between sessions - just reopen the app and your roster loads automatically
- **Backup Downloads:** Use the "Download Backup Copy" button if you want an additional backup
- Check-in data files are automatically saved for your records
