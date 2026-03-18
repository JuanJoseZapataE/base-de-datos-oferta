@echo off
REM Script para iniciar backend (FastAPI) y abrir frontend automaticamente.
REM Ejecutar este archivo haciendo doble clic.

REM 1) Ir a la carpeta raiz del proyecto
cd /d "%~dp0"

REM 2) Abrir el frontend (index.html) en el navegador predeterminado
start "" "%~dp0frontend\index.html"

REM 3) Lanzar PowerShell con los mismos pasos de activacion.txt en SEGUNDO PLANO
REM    para que no se muestre ninguna terminal ni icono en la barra de tareas.
REM    Comandos equivalentes a:
REM      Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
REM      .\fastapi_app\oferta\Scripts\Activate.ps1
REM      python -m uvicorn fastapi_app.main:app --reload
start "" powershell -WindowStyle Hidden -NoProfile -Command "Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass; & '.\fastapi_app\oferta\Scripts\Activate.ps1'; python -m uvicorn fastapi_app.main:app --reload"

REM Cerrar esta ventana de .bat; el backend seguira corriendo en segundo plano.
exit
