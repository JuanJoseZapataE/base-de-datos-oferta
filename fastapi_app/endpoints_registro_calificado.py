"""
Endpoints para Registro Calificado - Se importan en main.py
"""
from fastapi import UploadFile, File, HTTPException
from fastapi.responses import JSONResponse
from registro_calificado import (
    process_registro_calificado, 
    read_excel_registro_calificado, 
    get_registro_calificado_data
)


def setup_registro_calificado_endpoints(app):
    """Registra los endpoints de Registro Calificado en la app FastAPI"""

    @app.post('/registro-calificado/upload-excel')
    async def upload_registro_calificado(file: UploadFile = File(...)):
        """Sube un archivo Excel de Registro Calificado."""
        if not file.filename.lower().endswith(('.xls', '.xlsx')):
            raise HTTPException(status_code=400, detail='El archivo debe ser .xls o .xlsx')

        content = await file.read()
        
        try:
            df = read_excel_registro_calificado(content)
            result = await process_registro_calificado(df)
            return JSONResponse(result)
        except HTTPException:
            raise
        except Exception as e:
            raise HTTPException(status_code=500, detail=f'Error al procesar archivo: {str(e)}')


    @app.get('/registro-calificado/data')
    async def get_registro_calificado():
        """Obtiene los datos cargados de Registro Calificado."""
        try:
            data = await get_registro_calificado_data()
            return JSONResponse({'items': data})
        except Exception as e:
            return JSONResponse({'items': [], 'error': str(e)}, status_code=500)
