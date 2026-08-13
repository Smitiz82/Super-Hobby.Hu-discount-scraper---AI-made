@echo off
cd /d "%~dp0"
py tracker.py >> tracker.log 2>&1
