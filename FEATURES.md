# Features Summary

## ✅ Implemented Features

### 1. In-Person Attendance Tracking
- ✅ QR code generation that links to Qualtrics check-in forms
- ✅ Import check-in data from Qualtrics exports (CSV/Excel files)
- ✅ Manual entry option for professors/TAs
- ✅ Automatic point calculation:
  - **0.6 points** for on-time attendance (within threshold)
  - **0.2 points** for late attendance (after threshold)
- ✅ Configurable late threshold (default: 15 minutes)
- ✅ Configurable class start time
- ✅ Real-time roster updates
- ✅ Check-in records saved to CSV files for record keeping
- ✅ App is for Professors/TAs only - students check in via external forms

### 2. Zoom Meeting Attendance Tracking
- ✅ Excel file upload for Zoom meeting reports
- ✅ Automatic detection of name and duration columns
- ✅ Duration parsing (supports multiple formats):
  - HH:MM:SS (hours:minutes:seconds)
  - MM:SS (minutes:seconds)
  - Minutes (numeric)
- ✅ Point calculation based on participation duration:
  - **0.6 points** for 30+ minutes
  - **0.2 points** for less than 30 minutes
  - **0.0 points** for no attendance
- ✅ Date selection for meeting date
- ✅ Automatic roster updates

### 3. Name Format Handling
- ✅ Supports "Last Name, First Name" format (in-person check-in)
- ✅ Supports "First Name Last Name" format (Zoom Excel files)
- ✅ Automatic name matching between different formats
- ✅ Flexible roster name formats

### 4. Roster Management
- ✅ Excel and CSV file support
- ✅ Automatic attendance column creation (one per date)
- ✅ Total points calculation across all dates
- ✅ Export updated roster to Excel
- ✅ View attendance history
- ✅ Preserve all existing roster data

### 5. User Interface
- ✅ Modern Streamlit web interface
- ✅ Tab-based navigation:
  - In-Person Check-In
  - Zoom Attendance
  - View Roster
- ✅ Sidebar with settings and file uploads
- ✅ Real-time feedback and error messages
- ✅ Data visualization with pandas DataFrames

### 6. Data Management
- ✅ Session state management
- ✅ CSV file generation for check-ins
- ✅ Excel file export for roster
- ✅ Duplicate prevention
- ✅ Data validation

## 📋 Usage Workflow

### In-Person Class
1. Upload roster file (download from OneDrive and upload to app)
2. Set class start time and late threshold
3. Create Qualtrics check-in form
4. Generate QR code in app linking to Qualtrics form
5. Display QR code for students to scan and check in
6. Export responses from Qualtrics (CSV/Excel)
7. Import Qualtrics export file into app
8. Process check-ins to update roster with points
9. Points are assigned based on check-in time (0.6 on-time, 0.2 late)

### Zoom Class
1. Export Zoom meeting report as Excel
2. Upload Zoom Excel file
3. Select meeting date
4. Process attendance automatically
5. Points are assigned based on participation duration

### View and Export
1. View updated roster with all attendance points
2. See total points per student
3. Export roster to Excel file

## 🎯 Key Requirements Met

✅ QR code check-in for students
✅ Name format: "Last Name, First Name" for in-person
✅ Name format: "First Name Last Name" for Zoom
✅ Point system: 0.6 (on-time/full participation), 0.2 (late/partial)
✅ CSV/Excel file recording
✅ Roster file updates
✅ Zoom duration tracking (30+ minutes = 0.6, <30 = 0.2, no show = 0)
✅ Automatic name matching between formats
