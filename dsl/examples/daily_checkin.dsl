# Daily Check-In Processing Script
# This script processes daily check-ins from Qualtrics

ECHO "Starting daily check-in processing..."

# Load the roster
LOAD ROSTER roster_attendance.xlsx

# Set check-in time thresholds
SET CHECKIN TIMES EARLY_BIRD 11:00 REGULAR 11:36

# Process today's check-ins (date will be auto-detected from file)
PROCESS CHECKIN "uploads/today_checkins.csv"

# Save the updated roster
SAVE ROSTER

ECHO "Daily check-in processing complete!"

