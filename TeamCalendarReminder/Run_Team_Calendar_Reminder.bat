@echo off
REM Double-click this file to start the Team Calendar & Reminder tool.
REM It runs the PowerShell script without changing any system-wide
REM PowerShell security settings (the -ExecutionPolicy Bypass only
REM applies to this one window) and does not require admin rights.
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Team_Calendar_Reminder.ps1"
