# DSL (Domain Specific Language) Module

This folder contains all DSL-related files for the Attendance Tracker application.

## Structure

```
dsl/
├── __init__.py              # Python package initialization
├── ATTENDANCE_DSL.md        # Complete DSL specification and documentation
├── DSL_README.md            # Quick start guide for DSL
├── dsl_executor.py          # Standalone command-line DSL executor
├── dsl_integrated.py        # Flask-integrated DSL executor
├── examples/                # Example DSL scripts
│   ├── daily_checkin.dsl
│   ├── weekly_attendance.dsl
│   └── zoom_processing.dsl
└── README.md                # This file
```

## Files

### Documentation
- **ATTENDANCE_DSL.md**: Complete specification of all DSL commands, syntax rules, and examples
- **DSL_README.md**: Quick start guide with common use cases
- **README.md**: This file - folder structure overview

### Executors
- **dsl_executor.py**: Standalone executor for running DSL scripts from the command line
- **dsl_integrated.py**: Integrated executor that works with the Flask app and can access app functions

### Examples
- **examples/**: Contains example DSL scripts demonstrating common workflows

## Usage

### Command Line
```bash
python dsl/dsl_executor.py dsl/examples/daily_checkin.dsl
```

### Flask Integration
The DSL executor is integrated into the Flask app and can be accessed via the web interface at `/dsl`.

### Import in Python
```python
from dsl import IntegratedDSLExecutor, DSLExecutor
```

## Notes

- The `templates/dsl.html` file remains in the templates folder as it's required by Flask
- All DSL-related code and documentation is now centralized in this folder
- The examples folder contains sample scripts you can use as templates

