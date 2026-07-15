@echo off
cd /d "%~dp0"
python -m pip install -r requirements.txt
if errorlevel 1 goto :error
python -m streamlit run app.py
goto :eof

:error
echo.
echo Setup failed. Confirm that Python 3.10 or newer is installed, then run this file again.
pause
