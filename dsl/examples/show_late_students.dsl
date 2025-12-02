# Show Late Students Example Script
# This script demonstrates how to find students who were late (0.2 points) for a specific date

ECHO "Loading roster..."

# Load the roster
LOAD ROSTER roster_attendance.xlsx

ECHO "Finding late students for November 4, 2025..."

# Show late students for a specific date
# Use the roster column format: T,Nov.4 (Tuesday, November 4)
SHOW LATE STUDENTS DATE "T,Nov.4"

ECHO "Done!"

