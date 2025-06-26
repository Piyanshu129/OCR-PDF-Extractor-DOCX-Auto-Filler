@echo off
echo Starting the OCR DOCX Filler App...
cd /d %~dp0

:: Create virtual env if not already
if not exist venv (
    echo Creating virtual environment...
    python -m venv venv
    call venv\Scripts\activate
    pip install -r requirements.txt
) else (
    call venv\Scripts\activate
)

streamlit run app.py
pause



orr


@echo off
cd /d "C:\Users\YourName\Projects\ocr_app"
call venv\Scripts\activate.bat
streamlit run app.py

@echo off
cd /d "C:\Users\YourName\PycharmProjects\pdftoexcel"
call venv\Scripts\activate.bat
streamlit run app.py
pause
