@echo off
REM Inicia el servidor Uvicorn (Windows)
python -m pip install -r "%~dp0requirements.txt"
python "%~dp0run_uvicorn.py"
pause
