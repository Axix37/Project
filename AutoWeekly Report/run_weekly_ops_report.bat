@echo off
:: ============================================================
::  Weekly Operations Report
::  Schedule this file in Windows Task Scheduler to run weekly.
::  Builds Weekly_Ops_Report_YYYY-MM-DD.pptx (see PPTX_OUTPUT_DIR in
::  config.env). The script refreshes the Rework rerun workbook itself
::  before reading it; refresh the Power BI add-in slides in PowerPoint.
::
::  Trigger:  Weekly, Monday, 06:00
::  Action:   Start a program
::    Program:   python
::    Arguments: weekly_ops_report.py
::    Start in:  %~dp0
:: ============================================================
cd /d "%~dp0"
python weekly_ops_report.py --config config.env
