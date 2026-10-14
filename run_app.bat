@echo off
title Employee Reimbursement AI Automation System
echo ===================================================
echo Starting Employee Reimbursement Automation System...
echo ===================================================

where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Python is not found in PATH! Please install Python 3.10+ from python.org.
    pause
    exit /b
)

echo Installing / Checking required libraries...
python -m pip install streamlit pandas openpyxl pdfplumber pypdf pypdfium2 pytesseract pillow rapidfuzz

echo Launching Streamlit interface...
python -m streamlit run app.py

pause
