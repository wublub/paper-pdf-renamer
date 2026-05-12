@echo off
chcp 65001 >nul
cd /d "%~dp0"
python pdf_renamer.py
if errorlevel 1 pause
