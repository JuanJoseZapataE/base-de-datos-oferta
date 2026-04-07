@echo off
REM Script para crear el entorno virtual "oferta" e instalar dependencias
REM Ejecutar este archivo haciendo doble clic.

REM 1) Ir a la carpeta raiz del proyecto
cd /d "%~dp0"

REM 2) Verificar si ya existe la carpeta del entorno
IF EXIST "fastapi_app\oferta" (
    echo El entorno virtual "oferta" ya existe.
) ELSE (
    echo Creando entorno virtual "oferta"...
    python -m venv "fastapi_app\oferta"
    IF ERRORLEVEL 1 (
        echo Error al crear el entorno virtual. Asegurate de tener Python instalado en PATH.
        pause
        exit /b 1
    )
)

REM 3) Activar entorno y instalar requirements (usar los de la carpeta fastapi_app)
call "fastapi_app\oferta\Scripts\activate.bat"
IF ERRORLEVEL 1 (
    echo No se pudo activar el entorno virtual.
    pause
    exit /b 1
)

echo Instalando dependencias desde fastapi_app\requirements.txt ...
pip install --upgrade pip
pip install -r "fastapi_app\requirements.txt"
IF ERRORLEVEL 1 (
    echo Hubo un error instalando las dependencias.
    pause
    exit /b 1
)

echo.
echo Entorno "oferta" listo con dependencias instaladas.
echo Ahora puedes usar iniciar_sistema.bat para arrancar el sistema.
pause
exit /b 0
