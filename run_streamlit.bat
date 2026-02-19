@echo off
REM SEMRUSH Data Extractor - Streamlit App Launcher

echo.
echo ========================================
echo   SEMRUSH Data Extractor - Streamlit
echo ========================================
echo.

REM Check if virtual environment exists
if not exist ".venv" (
    echo Creating virtual environment...
    python -m venv .venv
)

echo Activating virtual environment...
call .\.venv\Scripts\activate.bat

echo Installing dependencies...
pip install -r requirements.txt --quiet

echo.
echo Launching Streamlit app...
echo The app should open in your default browser at http://localhost:8501
echo.
echo Press Ctrl+C to stop the server
echo.

streamlit run streamlit_app.py

pause
