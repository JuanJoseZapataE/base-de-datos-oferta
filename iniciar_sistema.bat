@echo off
REM Script para iniciar backend (FastAPI) y abrir frontend.
REM Modificado para mantener la ventana abierta y ver errores.

REM 1) Ir a la carpeta raiz del proyecto
cd /d "%~dp0"

REM 2) Abrir el frontend (index.html) en el navegador predeterminado
start "" "%~dp0frontend\index.html"

REM 3) Lanzar PowerShell de forma VISIBLE y mantener la sesión activa
REM    - NoExit: Evita que la consola se cierre si el comando falla o termina.
REM    - Se eliminó WindowStyle Hidden para que puedas ver el log de Uvicorn.
powershell -NoProfile -NoExit -Command "Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass; & '.\fastapi_app\oferta\Scripts\Activate.ps1'; python -m uvicorn fastapi_app.main:app --reload"

REM El comando 'pause' sirve por si PowerShell no llega a lanzarse, 
REM así la ventana de CMD te avisará del problema.
pause