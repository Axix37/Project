@echo off

cd /d "%~dp0"
python weekly_ops_report.py --config config.env
