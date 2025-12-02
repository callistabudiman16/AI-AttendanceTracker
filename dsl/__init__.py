"""
DSL (Domain Specific Language) Module for Attendance Tracker

This package contains:
- dsl_executor.py: Standalone DSL executor for command-line use
- dsl_integrated.py: Integrated DSL executor for Flask app
- examples/: Example DSL scripts
- ATTENDANCE_DSL.md: Complete DSL specification
- DSL_README.md: Quick start guide
"""

from .dsl_integrated import IntegratedDSLExecutor
from .dsl_executor import DSLExecutor

__all__ = ['IntegratedDSLExecutor', 'DSLExecutor']
