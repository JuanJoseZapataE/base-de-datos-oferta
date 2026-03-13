# PowerShell script para iniciar el servidor
pip install -r "$PSScriptRoot\requirements.txt"
python "$PSScriptRoot\run_uvicorn.py"
Read-Host -Prompt "Presiona Enter para cerrar"
