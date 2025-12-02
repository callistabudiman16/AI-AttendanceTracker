# Zoom Meeting Attendance Processing Script

ECHO "Processing Zoom meeting attendance..."

# Load roster
LOAD ROSTER roster_attendance.xlsx

# Process Zoom meeting (date auto-detected from file)
PROCESS ZOOM "uploads/zoom_meeting_oct23.xlsx"

# Save updated roster
SAVE ROSTER

ECHO "Zoom attendance processed successfully!"

