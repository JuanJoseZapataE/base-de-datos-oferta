"""
Módulo para gestionar la carga de archivos de Registro Calificado
"""
import pandas as pd
import io
from fastapi import HTTPException
from sqlalchemy import create_engine, text
import os
from dotenv import load_dotenv

load_dotenv(os.path.join(os.path.dirname(__file__), '.env'))
DATABASE_URL = os.getenv("DATABASE_URL", "mysql+pymysql://root@127.0.0.1/Oferta")
engine = create_engine(DATABASE_URL)


def read_excel_registro_calificado(content):
    """Lee un archivo Excel de registro calificado"""
    try:
        df = pd.read_excel(io.BytesIO(content), sheet_name=0, dtype=str)
        return df
    except Exception as e:
        raise HTTPException(
            status_code=400,
            detail=f'Error al leer el archivo Excel: {str(e)}'
        )


def normalize_col_name(col_name: str) -> str:
    """Normaliza nombres de columna para búsqueda flexible"""
    return str(col_name).lower().strip().replace(' ', '_').replace('-', '_')


async def process_registro_calificado(df: pd.DataFrame):
    """
    Procesa el DataFrame de registro calificado.
    Se espera que tenga columnas como: numero_ficha, estado, fecha, etc.
    """
    if df.empty:
        raise HTTPException(status_code=400, detail='El Excel no contiene filas')

    # Normalizar nombres de columnas
    df.columns = [normalize_col_name(col) for col in df.columns]

    # Posibles nombres de columna para identificar ficha
    ficha_aliases = [
        'numero_ficha', 'numero_de_ficha', 'identificador_ficha', 
        'codigo_ficha', 'cod_ficha', 'ficha'
    ]
    
    # Encontrar la columna de ficha
    ficha_col = None
    for alias in ficha_aliases:
        if alias in df.columns:
            ficha_col = alias
            break
    
    if not ficha_col:
        raise HTTPException(
            status_code=400,
            detail=f'No se encontró columna de número de ficha. Esperadas: {", ".join(ficha_aliases)}'
        )

    # Validar que haya datos
    df = df.dropna(subset=[ficha_col])
    if df.empty:
        raise HTTPException(status_code=400, detail='No hay fichas válidas en el archivo')

    # Insertar/Actualizar en la tabla registro_calificado
    try:
        with engine.connect() as conn:
            # Crear tabla si no existe
            create_table_sql = """
            CREATE TABLE IF NOT EXISTS registro_calificado (
                id INT AUTO_INCREMENT PRIMARY KEY,
                numero_ficha INT NOT NULL,
                estado VARCHAR(100),
                fecha_carga DATETIME DEFAULT CURRENT_TIMESTAMP,
                datos_adicionales LONGTEXT,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                updated_at DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                UNIQUE KEY unique_ficha (numero_ficha)
            )
            """
            conn.execute(text(create_table_sql))
            conn.commit()

            # Insertar registros
            for idx, row in df.iterrows():
                numero_ficha = str(row[ficha_col]).strip()
                if not numero_ficha or numero_ficha == 'nan':
                    continue

                try:
                    numero_ficha_int = int(float(numero_ficha))
                except (ValueError, TypeError):
                    continue

                # Serializar resto de datos
                datos_dict = {k: v for k, v in row.to_dict().items() if k != ficha_col}
                datos_json = str(datos_dict)

                insert_sql = """
                INSERT INTO registro_calificado (numero_ficha, estado, datos_adicionales)
                VALUES (:numero_ficha, 'Cargado', :datos_adicionales)
                ON DUPLICATE KEY UPDATE
                    datos_adicionales = :datos_adicionales,
                    updated_at = CURRENT_TIMESTAMP
                """
                conn.execute(
                    text(insert_sql),
                    {
                        'numero_ficha': numero_ficha_int,
                        'datos_adicionales': datos_json
                    }
                )
            
            conn.commit()

        return {
            'status': 'success',
            'message': f'Se cargaron {len(df)} registros correctamente',
            'rows_processed': len(df)
        }

    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f'Error al guardar en base de datos: {str(e)}'
        )


async def get_registro_calificado_data():
    """Obtiene los datos de registro calificado de la base de datos"""
    try:
        with engine.connect() as conn:
            result = conn.execute(text("""
                SELECT id, numero_ficha, estado, fecha_carga, updated_at
                FROM registro_calificado
                ORDER BY updated_at DESC
                LIMIT 100
            """))
            rows = result.fetchall()
            return [dict(row._mapping) for row in rows]
    except Exception as e:
        # Si la tabla no existe, retornar lista vacía
        return []
