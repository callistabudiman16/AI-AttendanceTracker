# Weekly Attendance Processing Script
# Processes attendance for a full week

ECHO "Processing weekly attendance..."

# Load roster
LOAD ROSTER roster_attendance.xlsx

# Configure settings
SET CHECKIN TIMES EARLY_BIRD 11:00 REGULAR 11:36

# Monday - In-person class
ECHO "Processing Monday check-ins..."
PROCESS CHECKIN "uploads/week1_monday.csv" DATE 2025-11-03

# Wednesday - In-person class
ECHO "Processing Wednesday check-ins..."
PROCESS CHECKIN "uploads/week1_wednesday.csv" DATE 2025-11-05

# Friday - Zoom meeting
ECHO "Processing Friday Zoom attendance..."
PROCESS ZOOM "uploads/week1_friday_zoom.xlsx" DATE 2025-11-07

# Save and backup
SAVE ROSTER
DOWNLOAD ROSTER "backups/week1_final_roster.xlsx"

ECHO "Weekly processing complete!"

