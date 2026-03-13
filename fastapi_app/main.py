from fastapi import FastAPI, UploadFile, File, HTTPException, Form
from fastapi.responses import JSONResponse, StreamingResponse
from fastapi.encoders import jsonable_encoder
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List, Optional
import pandas as pd
import io
import os
from datetime import datetime
import re
import unicodedata
from dotenv import load_dotenv
# Cargar .env desde la carpeta del paquete (asegura carga aunque el cwd sea el padre)
load_dotenv(os.path.join(os.path.dirname(__file__), '.env'))
from sqlalchemy import create_engine, text, bindparam
from openpyxl.styles import Alignment, Font
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

# URL de la base de datos: editar o usar la variable de entorno DATABASE_URL
# Ejemplo: mysql+pymysql://root:password@localhost/sena_oferta
DATABASE_URL = os.getenv("DATABASE_URL", "mysql+pymysql://root@127.0.0.1/Oferta")

engine = create_engine(DATABASE_URL)
app = FastAPI(title="Importador Excel -> MySQL (sena_oferta)")

# Habilitar CORS para permitir peticiones desde el frontend local
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)




@app.get('/')
def root():
    return {'message': 'API running. Usa /docs para ver los endpoints.'}

EXPECTED_COLUMNS = [
    'cod_regional', 'regional', 'cod_municipio', 'municipio', 'cod_centro', 'centro_formacion',
    'cod_programa', 'denominacion_programa', 'cod_ficha', 'estado_ficha', 'jornada', 'nivel_formacion',
    'cupo', 'inscritos_primera_opcion', 'inscritos_segunda_opcion', 'oferta', 'tipo', 'perfil_ingreso', 'periodo'
]

PROGRAMAS_COLUMNS = [
    'centro_formacion',
    'numero_ficha',
    'ciudad_municipio',
    'fecha_inicio',
    'fecha_fin',
    'nivel_formacion',
    'denominacion_programa',
    # Antes: estrato_programa. Ahora se usa como "estrategia del programa".
    'estrategia_programa',
    'convenio',
    'cupos',
    'aprendices_activos',
    'certificado',
    'tipo_formacion',
    'estado_curso',
    'fecha_corte',
]


def ensure_programas_table():
    create_sql = """
    CREATE TABLE IF NOT EXISTS programas_formacion (
        id BIGINT NOT NULL AUTO_INCREMENT PRIMARY KEY,
        centro_formacion VARCHAR(200) NULL,
        numero_ficha BIGINT NULL,
        ciudad_municipio VARCHAR(150) NULL,
        fecha_inicio DATE NULL,
        fecha_fin DATE NULL,
        nivel_formacion VARCHAR(100) NULL,
        denominacion_programa VARCHAR(255) NULL,
        estrategia_programa VARCHAR(255) NULL,
        convenio VARCHAR(255) NULL,
        cupos INT NULL,
        aprendices_activos INT NULL,
        certificado VARCHAR(255) NULL,
        tipo_formacion VARCHAR(100) NULL,
        estado_curso VARCHAR(100) NULL,
        fecha_corte DATE NULL,
        INDEX idx_programas_fecha_corte (fecha_corte),
        INDEX idx_programas_municipio (ciudad_municipio),
        INDEX idx_programas_numero_ficha (numero_ficha)
    )
    """
    with engine.begin() as conn:
        conn.execute(text(create_sql))
        # Ajuste de esquema para instalaciones previas donde convenio quedo corto.
        try:
            conn.execute(text('ALTER TABLE programas_formacion MODIFY COLUMN convenio VARCHAR(255) NULL'))
        except Exception:
            pass
        # Migracion suave: instalaciones antiguas usaban estrato_programa.
        # Renombrar a estrategia_programa y ampliar longitud si existe.
        try:
            conn.execute(text('ALTER TABLE programas_formacion CHANGE COLUMN estrato_programa estrategia_programa VARCHAR(255) NULL'))
        except Exception:
            pass
        # Asegurar columna estado_curso para instalaciones previas.
        try:
            conn.execute(text('ALTER TABLE programas_formacion ADD COLUMN estado_curso VARCHAR(100) NULL'))
        except Exception:
            pass


ensure_programas_table()


def normalize_cols(cols):
    return [
        normalize_col_name(c) if isinstance(c, str) else c
        for c in cols
    ]


def normalize_col_name(value: str) -> str:
    if not isinstance(value, str):
        return value
    s = unicodedata.normalize('NFKD', value).encode('ascii', 'ignore').decode('ascii')
    s = s.strip().lower()
    s = s.replace(' ', '_').replace('.', '').replace('-', '_').replace('/', '_')
    s = s.replace('(', '').replace(')', '')
    return s


def looks_like_expected_headers(columns) -> bool:
    normalized = set(normalize_cols(columns))
    expected = set(EXPECTED_COLUMNS)
    matches = len(normalized.intersection(expected))
    return matches >= 4 and ('cod_ficha' in normalized or 'cod_regional' in normalized)


def detect_header_row(df_raw: pd.DataFrame, max_scan_rows: int = 30) -> Optional[int]:
    expected = set(EXPECTED_COLUMNS)
    scan_limit = min(max_scan_rows, len(df_raw.index))
    best_row = None
    best_score = 0

    for idx in range(scan_limit):
        row_values = [value for value in df_raw.iloc[idx].tolist() if pd.notna(value)]
        normalized_row = normalize_cols([str(value) for value in row_values])
        normalized_set = set(normalized_row)
        score = len(normalized_set.intersection(expected))

        if 'cod_ficha' in normalized_set:
            score += 2

        if score > best_score:
            best_score = score
            best_row = idx

    if best_row is not None and best_score >= 4:
        return int(best_row)
    return None


def read_excel_with_header_detection(content: bytes) -> pd.DataFrame:
    read_attempts = [
        {'engine': 'openpyxl'},
        {},
    ]
    last_error = None

    for kwargs in read_attempts:
        try:
            df_default = pd.read_excel(io.BytesIO(content), **kwargs)
            if looks_like_expected_headers(df_default.columns):
                return df_default

            df_raw = pd.read_excel(io.BytesIO(content), header=None, **kwargs)
            header_row = detect_header_row(df_raw)
            if header_row is not None:
                return pd.read_excel(io.BytesIO(content), header=header_row, **kwargs)

            return df_default
        except Exception as e:
            last_error = e

    raise HTTPException(status_code=400, detail=f'No se pudo leer el Excel: {last_error}')


def normalize_tipo(value: str) -> str:
    if not isinstance(value, str):
        return value
    v = value.strip().lower()
    if 'presencial' in v and ('distancia' in v or 'a distancia' in v):
        return 'PRESENCIAL Y A DISTANCIA'
    if 'presencial' in v:
        return 'PRESENCIAL'
    if 'distancia' in v or 'a distancia' in v:
        return 'A DISTANCIA'
    if 'virtual' in v:
        return 'VIRTUAL'
    return value.upper()


def normalize_oferta(value) -> str:
    if value is None:
        return None
    s = str(value).strip().upper()
    # Prefer to return a single-character code compatible with CHAR(1) in the DB
    # Map digits
    if s.isdigit():
        if s[-1] in '1234':
            return s[-1]
    # Map common roman numerals I..IV
    roman_map = {'I': '1', 'II': '2', 'III': '3', 'IV': '4'}
    if s in roman_map:
        return roman_map[s]
    # Map by keywords: VIRTUAL -> 4, PRESENCIAL or DISTANCIA -> 1
    if 'VIRTUAL' in s:
        return '4'
    if 'PRESENCIAL' in s or 'DISTANCIA' in s or 'A DISTANCIA' in s:
        return '1'
    # If contains a digit anywhere, take last digit
    for ch in reversed(s):
        if ch in '1234':
            return ch
    # Fallback: take first character (trimmed) to avoid length errors
    return s[0]


def export_header_label(column_name: str) -> str:
    """Convierte nombres técnicos a encabezados legibles para Excel."""
    if not column_name:
        return ''
    s = str(column_name).strip().replace('_', ' ')
    s = s.replace(' cod ', ' codigo ')
    if s.startswith('cod '):
        s = 'codigo ' + s[4:]
    if s == 'cod':
        s = 'codigo'
    return ' '.join(word.capitalize() for word in s.split())


def get_first_existing_column(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    existing = set(df.columns)
    for c in candidates:
        if c in existing:
            return c
    return None


def get_column_by_keywords(df: pd.DataFrame, keyword_groups: List[List[str]]) -> Optional[str]:
    """Busca una columna cuyo nombre contenga todos los tokens de algun grupo."""
    cols = [str(c) for c in df.columns]
    for group in keyword_groups:
        for col in cols:
            if all(token in col for token in group):
                return col
    return None


def clean_optional_text(v):
    if pd.isna(v):
        return None
    s = str(v).strip()
    if s == '':
        return None
    if s.lower() in {'nan', 'none', 'null', 'nat', '<na>'}:
        return None
    return s


def read_excel_basic(content: bytes) -> pd.DataFrame:
    """Lee un Excel simple desde bytes.

    - Primero intenta como .xlsx con openpyxl.
    - Si falla y el backend intenta usar otro engine que no está instalado
      (por ejemplo xlrd para .xls), se devuelve un error 400 entendible
      en lugar de romper con 500 Internal Server Error.
    """
    # 1) Intentar siempre como .xlsx (openpyxl)
    try:
        return pd.read_excel(io.BytesIO(content), engine='openpyxl')
    except Exception as e_openpyxl:
        # 2) Fallback genérico de pandas. Si el archivo es .xls y está instalado xlrd,
        #    pandas usará ese engine de forma automática.
        try:
            return pd.read_excel(io.BytesIO(content))
        except ImportError:
            # Caso típico: archivo .xls pero xlrd no está instalado.
            raise HTTPException(
                status_code=400,
                detail=(
                    'No se pudo leer el Excel porque falta soporte para archivos .xls. '
                    'Vuelve a ejecutar la instalacion de requisitos para habilitarlo '
                    'o convierte el archivo a .xlsx.'
                ),
            )
        except Exception:
            # Si tampoco se puede leer aquí, reportar error de formato de archivo.
            raise HTTPException(
                status_code=400,
                detail='No se pudo leer el Excel. Verifica que sea un archivo de Excel valido (.xls o .xlsx).',
            ) from e_openpyxl


def read_excel_no_header(content: bytes) -> pd.DataFrame:
    try:
        return pd.read_excel(io.BytesIO(content), header=None, engine='openpyxl')
    except Exception:
        return pd.read_excel(io.BytesIO(content), header=None)


def read_excel_with_header_row(content: bytes, header_row: int) -> pd.DataFrame:
    try:
        return pd.read_excel(io.BytesIO(content), header=header_row, engine='openpyxl')
    except Exception:
        return pd.read_excel(io.BytesIO(content), header=header_row)


def extract_fecha_corte_from_filename(filename: str):
    """Extrae fecha de corte desde nombre tipo PE-04_20260306_15+55.xlsx -> 2026-03-06."""
    if not filename:
        return None
    match = re.search(r'(\d{8})', filename)
    if not match:
        return None
    raw = match.group(1)
    try:
        return datetime.strptime(raw, '%Y%m%d').date()
    except Exception:
        return None


@app.post('/upload-excel')
async def upload_excel(file: UploadFile = File(...), periodo: Optional[int] = Form(None), oferta: Optional[str] = Form(None), tipo: Optional[str] = Form(None)):
    if not file.filename.lower().endswith(('.xls', '.xlsx')):
        raise HTTPException(status_code=400, detail='El archivo debe ser .xls o .xlsx')

    content = await file.read()
    df = read_excel_with_header_detection(content)

    # Normalizar nombres de columnas
    df.columns = normalize_cols(df.columns)

    # Preparar dataframe para insertar: asegurarnos de que existan todas las columnas
    # Si faltan columnas opcionales (oferta, tipo, perfil_ingreso) no impediremos la subida;
    # las creamos y las rellenamos desde los placeholders del formulario más abajo.
    df_to_insert = df.copy()
    # Si el Excel tiene exactamente el número esperado de columnas, asignar nombres por posición
    if df.shape[1] == len(EXPECTED_COLUMNS):
        df_to_insert.columns = EXPECTED_COLUMNS
    # Añadir las columnas faltantes (las inicializamos con None)
    for col in EXPECTED_COLUMNS:
        if col not in df_to_insert.columns:
            df_to_insert[col] = None
    # Reordenar columnas para consistencia
    df_to_insert = df_to_insert[EXPECTED_COLUMNS].copy()

    # Intentar convertir columnas numéricas; esto eliminará filas que sean en realidad
    # encabezados leídos como datos (p.ej. 'COD_REGIONAL') porque se convertirán a NaN.
    int_cols = ['cod_regional', 'cod_municipio', 'cod_centro', 'cod_programa', 'cod_ficha',
                'cupo', 'inscritos_primera_opcion', 'inscritos_segunda_opcion', 'periodo']
    for col in int_cols:
        if col in df_to_insert.columns:
            df_to_insert[col] = pd.to_numeric(df_to_insert[col], errors='coerce')

    # Quitar filas que no tengan un cod_ficha válido (clave primaria necesaria)
    if 'cod_ficha' in df_to_insert.columns:
        before = len(df_to_insert)
        df_to_insert = df_to_insert[df_to_insert['cod_ficha'].notna()].copy()
        removed = before - len(df_to_insert)
        if removed:
            print(f'Removed {removed} rows that looked like headers or had invalid cod_ficha')

    # Ahora convertir periodo y otras columnas a enteros donde aplique
    try:
        if 'periodo' in df_to_insert.columns:
            df_to_insert['periodo'] = pd.to_numeric(df_to_insert['periodo'], errors='coerce').astype('Int64')
    except Exception:
        pass

    # Normalizar columnas si vienen en el Excel
    if 'oferta' in df_to_insert.columns:
        df_to_insert['oferta'] = df_to_insert['oferta'].apply(lambda v: normalize_oferta(v) if pd.notna(v) else v)
    if 'tipo' in df_to_insert.columns:
        df_to_insert['tipo'] = df_to_insert['tipo'].apply(lambda v: normalize_tipo(v) if pd.notna(v) else v)

    # Comportamiento requerido:
    # - Si la columna existe y al menos una fila tiene valor, permitimos subir.
    #   - Si el formulario provee un valor, rellenamos los nulos con el valor del formulario.
    #   - Si el formulario no provee valor, dejamos nulos donde existan.
    # - Si la columna existe pero todas las filas están vacías, la tratamos como "ausente":
    #   requerimos el valor en el formulario para rellenar toda la columna.
    # - Si la columna no existe, requerimos el valor en el formulario.

    # periodo
    if 'periodo' in df_to_insert.columns:
        # ya intentamos convertir a numérico más arriba
        has_any = df_to_insert['periodo'].notna().any()
        if has_any:
            if periodo is not None:
                try:
                    periodo_val = int(periodo)
                    df_to_insert['periodo'] = df_to_insert['periodo'].fillna(periodo_val)
                except Exception:
                    raise HTTPException(status_code=400, detail='Periodo inválido')
            # si periodo no se provee, permitimos nulos tal como vienen en el Excel
        else:
            # columna existe pero vacía en todas las filas -> necesitamos formulario
            if periodo is None:
                raise HTTPException(status_code=400, detail='Periodo requerido (ni en Excel ni en el formulario)')
            try:
                periodo_val = int(periodo)
            except Exception:
                raise HTTPException(status_code=400, detail='Periodo inválido')
            df_to_insert['periodo'] = periodo_val
    else:
        # columna ausente
        if periodo is None:
            raise HTTPException(status_code=400, detail='Periodo requerido (ni en Excel ni en el formulario)')
        try:
            periodo_val = int(periodo)
        except Exception:
            raise HTTPException(status_code=400, detail='Periodo inválido')
        df_to_insert['periodo'] = periodo_val

    # oferta
    # Permitir que el Excel no tenga la columna 'oferta'. Si existe, rellenar nulos con el formulario cuando se provea.
    if 'oferta' in df_to_insert.columns:
        has_any = df_to_insert['oferta'].notna().any()
        if has_any:
            if oferta:
                oferta_norm = normalize_oferta(oferta)
                df_to_insert['oferta'] = df_to_insert['oferta'].fillna(oferta_norm)
            # si no se provee oferta en el formulario, dejamos nulos donde existan
        else:
            # columna existe pero vacía en todas las filas -> si el formulario tiene valor, usarlo, si no dejar nulos
            if oferta:
                oferta_norm = normalize_oferta(oferta)
                df_to_insert['oferta'] = oferta_norm
            else:
                df_to_insert['oferta'] = df_to_insert['oferta']
    else:
        # columna ausente -> crearla y rellenar con el valor del formulario si está, o dejar None
        oferta_norm = normalize_oferta(oferta) if oferta else None
        df_to_insert['oferta'] = oferta_norm

    # tipo
    # Permitir que el Excel no tenga la columna 'tipo'. Si existe, rellenar nulos con el formulario cuando se provea.
    if 'tipo' in df_to_insert.columns:
        has_any = df_to_insert['tipo'].notna().any()
        if has_any:
            if tipo:
                tipo_norm = normalize_tipo(tipo)
                df_to_insert['tipo'] = df_to_insert['tipo'].fillna(tipo_norm)
            # si no se provee tipo en el formulario, dejamos nulos donde existan
        else:
            if tipo:
                tipo_norm = normalize_tipo(tipo)
                df_to_insert['tipo'] = tipo_norm
            else:
                df_to_insert['tipo'] = df_to_insert['tipo']
    else:
        tipo_norm = normalize_tipo(tipo) if tipo else None
        df_to_insert['tipo'] = tipo_norm

    # perfil_ingreso: si la columna no existe simplemente crearla (se puede dejar vacía)
    if 'perfil_ingreso' not in df_to_insert.columns:
        df_to_insert['perfil_ingreso'] = None

    # Insertar en la base de datos usando pandas.to_sql (append)
    try:
        df_to_insert.to_sql('fichas_formacion', con=engine, if_exists='append', index=False)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'Error al insertar en la base de datos: {e}')

    return JSONResponse({'inserted': len(df_to_insert)})


@app.get('/fichas')
def get_fichas(
    periodo: Optional[int] = None,
    oferta: Optional[str] = None,
    tipo: Optional[str] = None,
    page: int = 1,
    per_page: int = 50,
):
    """Devuelve paginado: 50 por página por defecto. Respuesta JSON con items y metadatos.
    """
    # validar parámetros de paginación
    try:
        page = int(page)
    except Exception:
        page = 1
    try:
        per_page = int(per_page)
    except Exception:
        per_page = 50
    if page < 1:
        page = 1
    if per_page < 1 or per_page > 1000:
        per_page = 50

    clauses = []
    params = {}
    if periodo is not None:
        clauses.append('periodo = :periodo')
        params['periodo'] = int(periodo)
    if oferta is not None:
        clauses.append('oferta = :oferta')
        params['oferta'] = normalize_oferta(oferta)
    if tipo is not None:
        clauses.append('UPPER(tipo) = :tipo')
        params['tipo'] = normalize_tipo(tipo)

    where_sql = ''
    if clauses:
        where_sql = ' WHERE ' + ' AND '.join(clauses)

    count_sql = 'SELECT COUNT(*) AS total FROM fichas_formacion' + where_sql

    try:
        with engine.connect() as conn:
            total = conn.execute(text(count_sql), params).scalar() or 0
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'Error contando registros: {e}')

    offset = (page - 1) * per_page
    # Order by periodo asc, then oferta asc, then cod_ficha
    data_sql = f'SELECT * FROM fichas_formacion{where_sql} ORDER BY periodo ASC, oferta ASC, cod_ficha ASC LIMIT :limit OFFSET :offset'
    params2 = dict(params)
    params2['limit'] = per_page
    params2['offset'] = offset

    try:
        df = pd.read_sql(data_sql, con=engine, params=params2)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'Error al leer la base de datos: {e}')

    return JSONResponse({
        'items': df.to_dict(orient='records'),
        'total': int(total),
        'page': page,
        'per_page': per_page,
    })


@app.delete('/fichas/{cod_ficha}')
def delete_ficha(cod_ficha: int):
    """Eliminar una ficha por su `cod_ficha`. Devuelve 204 si fue eliminada, 404 si no existe."""
    try:
        with engine.begin() as conn:
            # Verificar existencia antes de eliminar
            exists = conn.execute(text('SELECT COUNT(*) FROM fichas_formacion WHERE cod_ficha = :id'), {'id': int(cod_ficha)}).scalar() or 0
            if int(exists) == 0:
                raise HTTPException(status_code=404, detail='Ficha no encontrada')
            # Ejecutar borrado (commit al salir del context manager)
            conn.execute(text('DELETE FROM fichas_formacion WHERE cod_ficha = :id'), {'id': int(cod_ficha)})
            return JSONResponse(status_code=204, content={})
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'Error eliminando ficha: {e}')
    # Si rowcount no indicó eliminación pero no hubo error, devolver 204
    return JSONResponse(status_code=204, content={})


@app.get('/fichas/count')
def fichas_count():
    """Endpoint diagnóstico: devuelve el total de filas y hasta 5 filas de ejemplo."""
    try:
        with engine.connect() as conn:
            total = conn.execute(text('SELECT COUNT(*) FROM fichas_formacion')).scalar() or 0
        sample = pd.read_sql('SELECT * FROM fichas_formacion LIMIT 5', con=engine)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'Error al consultar la base de datos: {e}')

    return JSONResponse({'total': int(total), 'sample': sample.to_dict(orient='records')})


@app.get('/fichas/all')
def fichas_all():
    """Devuelve todos los registros de la tabla `fichas_formacion` sin paginación."""
    try:
        # Orden por periodo asc (años más antiguos primero), luego por oferta asc y cod_ficha
        df = pd.read_sql('SELECT * FROM fichas_formacion ORDER BY periodo ASC, oferta ASC, cod_ficha ASC', con=engine)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'Error al leer la base de datos: {e}')

    return JSONResponse(df.to_dict(orient='records'))


@app.get('/fichas/export')
def export_fichas_excel(
    centro: Optional[str] = None,
    oferta: Optional[str] = None,
    estado: Optional[str] = None,
    tipo: Optional[str] = None,
    nivel: Optional[str] = None,
    periodo: Optional[int] = None,
    search: Optional[str] = None,
):
    """Exporta Excel con los filtros activos (mismos criterios del frontend)."""
    clauses = []
    params = {}

    if centro:
        clauses.append('LOWER(TRIM(centro_formacion)) = :centro')
        params['centro'] = centro.strip().lower()
    if oferta:
        clauses.append('oferta = :oferta')
        params['oferta'] = normalize_oferta(oferta)
    if estado:
        clauses.append('LOWER(TRIM(estado_ficha)) = :estado')
        params['estado'] = estado.strip().lower()
    if tipo:
        clauses.append('LOWER(TRIM(tipo)) = :tipo')
        params['tipo'] = normalize_tipo(tipo).strip().lower()
    if nivel:
        clauses.append('LOWER(TRIM(nivel_formacion)) = :nivel')
        params['nivel'] = nivel.strip().lower()
    if periodo is not None:
        clauses.append('periodo = :periodo')
        params['periodo'] = int(periodo)
    if search:
        clauses.append('LOWER(COALESCE(denominacion_programa, "")) LIKE :search')
        params['search'] = f"%{search.strip().lower()}%"

    where_sql = ''
    if clauses:
        where_sql = ' WHERE ' + ' AND '.join(clauses)

    sql = f'SELECT * FROM fichas_formacion{where_sql} ORDER BY periodo ASC, oferta ASC, cod_ficha ASC'

    try:
        df = pd.read_sql(text(sql), con=engine, params=params)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'Error al exportar desde la base de datos: {e}')

    # Excluir columna no requerida en exportacion.
    df_export = df.copy()
    if 'perfil_ingreso' in df_export.columns:
        df_export = df_export.drop(columns=['perfil_ingreso'])

    # Encabezados legibles para Excel: sin guion bajo y con formato titulo.
    df_export.columns = [export_header_label(col) for col in df_export.columns]

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name='fichas')

        ws = writer.book['fichas']
        max_row = ws.max_row
        max_col = ws.max_column

        # Ajuste de texto en todas las celdas (encabezado y datos).
        wrap_alignment = Alignment(wrap_text=True, vertical='top')
        for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
            for cell in row:
                cell.alignment = wrap_alignment

        # Encabezados en negrita para mejorar lectura.
        for cell in ws[1]:
            cell.font = Font(bold=True)

        # Ajustar ancho por contenido: texto largo -> columna mas ancha, texto corto -> mas angosta.
        min_width = 12
        max_width = 60
        for col_idx in range(1, max_col + 1):
            col_letter = get_column_letter(col_idx)
            max_len = 0
            for row_idx in range(1, max_row + 1):
                value = ws.cell(row=row_idx, column=col_idx).value
                cell_text = '' if value is None else str(value)
                if len(cell_text) > max_len:
                    max_len = len(cell_text)
            adjusted = min(max(max_len + 2, min_width), max_width)
            ws.column_dimensions[col_letter].width = adjusted

        # Crear una tabla de Excel para aplicar formato de tabla.
        if max_col >= 1 and max_row >= 1:
            last_col_letter = get_column_letter(max_col)
            table_ref = f'A1:{last_col_letter}{max_row}'
            table = Table(displayName='FichasExport', ref=table_ref)
            style = TableStyleInfo(
                name='TableStyleMedium9',
                showFirstColumn=False,
                showLastColumn=False,
                showRowStripes=True,
                showColumnStripes=False,
            )
            table.tableStyleInfo = style
            ws.add_table(table)
    output.seek(0)

    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f'fichas_export_{ts}.xlsx'

    return StreamingResponse(
        output,
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={'Content-Disposition': f'attachment; filename="{filename}"'},
    )


@app.post('/programas/upload-excel')
async def upload_programas_excel(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(('.xls', '.xlsx')):
        raise HTTPException(status_code=400, detail='El archivo debe ser .xls o .xlsx')

    content = await file.read()
    fecha_corte_file = extract_fecha_corte_from_filename(file.filename or '')
    if not fecha_corte_file:
        raise HTTPException(
            status_code=400,
            detail='No se pudo extraer fecha_corte del nombre del archivo. Usa formato como PE-04_20260306_15+55.xlsx',
        )

    # Solo mapear los campos definidos para la tabla programas_formacion.
    # Se incluyen variantes que suelen venir en archivos tipo SOFIA/planeacion.
    col_map = {
        'centro_formacion': ['centro_formacion', 'centro_de_formacion', 'nombre_centro', 'nombre_centro_formacion'],
        'numero_ficha': ['numero_ficha', 'numero_de_ficha', 'n_ficha', 'codigo_ficha', 'cod_ficha', 'identificador_ficha'],
        'ciudad_municipio': ['ciudad_municipio', 'ciudad_o_municipio', 'nombre_ciudad', 'municipio', 'nombre_municipio_curso'],
        'fecha_inicio': ['fecha_inicio', 'fecha_de_inicio', 'inicio_ficha', 'fecha_inicio_ficha'],
        'fecha_fin': ['fecha_fin', 'fecha_de_fin', 'fin_ficha', 'fecha_fin_ficha', 'fecha_terminacion_ficha'],
        'nivel_formacion': ['nivel_formacion', 'nivel_de_formacion', 'nombre_nivel_formacion'],
        'denominacion_programa': ['denominacion_programa', 'denominacion_del_programa', 'nombre_curso', 'nombre_programa', 'nombre_programa_formacion'],
        # Nuevo uso: estrategia del programa; se alimenta principalmente desde
        # encabezados tipo NOMBRE_PROGRAMA_ESPECIAL.
        'estrategia_programa': [
            'estrategia_programa',
            'estrategia_del_programa',
            'nombre_programa_especial',
        ],
        # Estado del curso, viene tipicamente como ESTADO_CURSO.
        'estado_curso': [
            'estado_curso',
            'estado_del_curso',
        ],
        'convenio': ['convenio', 'nombre_convenio', 'tipo_convenio'],
        'cupos': ['cupos', 'cupo', 'meta_cupo', 'meta_cupos', 'total_aprendices'],
        'aprendices_activos': ['aprendices_activos', 'total_aprendices_activos'],
        'certificado': ['certificado'],
        'tipo_formacion': ['tipo_formacion', 'tipo_de_formacion', 'nombre_tipo_formacion'],
    }

    # Detectar si los encabezados reales no estan en la primera fila (caso tipico: primera fila con titulo PE-04_...)
    alias_pool = set()
    for aliases in col_map.values():
        for alias in aliases:
            alias_pool.add(normalize_col_name(alias))

    df = read_excel_basic(content)
    if df.empty:
        raise HTTPException(status_code=400, detail='El Excel no contiene filas')

    norm_default_headers = [normalize_col_name(str(c)) for c in list(df.columns)]
    default_score = len(set(norm_default_headers).intersection(alias_pool))

    # Si detecta pocos encabezados utiles o muchos "unnamed", intenta encontrar la fila de encabezado correcta.
    unnamed_count = sum(1 for c in norm_default_headers if c.startswith('unnamed:'))
    if default_score < 3 or unnamed_count >= 5:
        df_raw = read_excel_no_header(content)
        scan_limit = min(40, len(df_raw.index))
        best_idx = None
        best_score = -1

        for idx in range(scan_limit):
            row_values = [str(v) for v in df_raw.iloc[idx].tolist() if pd.notna(v)]
            row_norm = set(normalize_cols(row_values))
            score = len(row_norm.intersection(alias_pool))
            if score > best_score:
                best_score = score
                best_idx = idx

        if best_idx is not None and best_score >= 3:
            df = read_excel_with_header_row(content, int(best_idx))

    df.columns = normalize_cols(df.columns)

    keyword_map = {
        'centro_formacion': [['centro'], ['nombre', 'centro']],
        'numero_ficha': [['ficha'], ['codigo', 'ficha']],
        'ciudad_municipio': [['ciudad'], ['municipio']],
        'fecha_inicio': [['fecha', 'inicio'], ['inicio']],
        'fecha_fin': [['fecha', 'fin'], ['fin']],
        'nivel_formacion': [['nivel', 'formacion'], ['nivel']],
        'denominacion_programa': [['nombre', 'curso'], ['denominacion', 'programa'], ['programa']],
        # Para estrategia_programa no usamos heuristica de palabras clave; se
        # confia en el mapeo explicito de col_map (NOMBRE_PROGRAMA_ESPECIAL).
        'convenio': [['convenio']],
        'cupos': [['cupo'], ['cupos']],
        'aprendices_activos': [['aprendices', 'activos'], ['activos']],
        'certificado': [['certificado']],
        'tipo_formacion': [['tipo', 'formacion']],
    }

    df_out = pd.DataFrame()
    mapped_sources = {}
    for target in PROGRAMAS_COLUMNS:
        aliases = [normalize_col_name(a) for a in col_map.get(target, [target])]
        source_col = get_first_existing_column(df, aliases)
        if not source_col and target in keyword_map:
            source_col = get_column_by_keywords(df, keyword_map[target])
        if source_col:
            df_out[target] = df[source_col]
            mapped_sources[target] = source_col
        else:
            df_out[target] = None

    # fecha_corte siempre se toma del nombre del archivo.
    df_out['fecha_corte'] = fecha_corte_file

    # Normalizacion de tipos
    for dcol in ['fecha_inicio', 'fecha_fin']:
        df_out[dcol] = pd.to_datetime(df_out[dcol], errors='coerce').dt.date

    for ncol in ['numero_ficha', 'cupos', 'aprendices_activos']:
        df_out[ncol] = pd.to_numeric(df_out[ncol], errors='coerce').astype('Int64')

    for scol in ['centro_formacion', 'ciudad_municipio', 'nivel_formacion', 'denominacion_programa', 'estrategia_programa', 'convenio', 'certificado', 'tipo_formacion', 'estado_curso']:
        df_out[scol] = df_out[scol].apply(clean_optional_text)

    total_rows_before_filter = len(df_out)
    num_ficha_with_value = int(df_out['numero_ficha'].notna().sum())
    denom_with_value = int(df_out['denominacion_programa'].fillna('').astype(str).str.strip().ne('').sum())

    # Eliminar filas completamente vacias en campos clave
    key_fields = ['numero_ficha', 'denominacion_programa']
    df_out = df_out[~df_out[key_fields].isna().all(axis=1)].copy()
    if df_out.empty:
        mapped_resume = ', '.join([f'{k}->{v}' for k, v in mapped_sources.items()]) if mapped_sources else 'ninguna'
        detected_headers = ', '.join([str(c) for c in list(df.columns)[:20]])
        raise HTTPException(
            status_code=400,
            detail=(
                'No se pudo insertar porque ninguna fila tiene datos en los campos clave '
                '(numero_ficha o denominacion_programa). '
                f'Filas leidas: {total_rows_before_filter}. '
                f'Filas con numero_ficha: {num_ficha_with_value}. '
                f'Filas con denominacion_programa: {denom_with_value}. '
                f'Mapeo detectado: {mapped_resume}. '
                f'Encabezados detectados (primeros 20): {detected_headers}'
            ),
        )

    # Validacion explicita para convenio (evita errores SQL opacos).
    convenio_max_len = 255
    if 'convenio' in df_out.columns:
        convenio_lengths = df_out['convenio'].dropna().astype(str).str.len()
        if not convenio_lengths.empty and int(convenio_lengths.max()) > convenio_max_len:
            bad_idx = convenio_lengths.idxmax()
            bad_value = str(df_out.loc[bad_idx, 'convenio'])
            preview = bad_value[:180]
            raise HTTPException(
                status_code=400,
                detail=(
                    f'El campo convenio supera el tamano permitido ({convenio_max_len}) en al menos una fila. '
                    f'Fila aproximada: {int(bad_idx) + 2}. '
                    f'Longitud encontrada: {int(convenio_lengths.max())}. '
                    f'Valor (preview): {preview}'
                ),
            )

    try:
        # Logica de persistencia:
        # - Si la ficha (numero_ficha) no existe aun en la tabla, se inserta
        #   el registro completo.
        # - Si la ficha ya existe, no se inserta una fila nueva; solo se
        #   actualiza la columna estrategia_programa (que viene del Excel).

        # Normalizar lista de fichas del Excel (no nulas)
        ficha_series = df_out['numero_ficha'].dropna() if 'numero_ficha' in df_out.columns else pd.Series([], dtype='Int64')
        ficha_ids = [int(x) for x in ficha_series.tolist()]

        existing_ids: set[int] = set()
        if ficha_ids:
            check_sql = text('SELECT DISTINCT numero_ficha FROM programas_formacion WHERE numero_ficha IN :ids').bindparams(bindparam('ids', expanding=True))
            with engine.connect() as conn:
                rows = conn.execute(check_sql, {'ids': ficha_ids}).fetchall()
            existing_ids = {int(r[0]) for r in rows if r and r[0] is not None}

        # Filas nuevas (fichas que aun no existen en la tabla)
        df_new = df_out[~df_out['numero_ficha'].isin(existing_ids)].copy() if 'numero_ficha' in df_out.columns else df_out.copy()
        if not df_new.empty:
            df_new.to_sql('programas_formacion', con=engine, if_exists='append', index=False)

        # Filas existentes: actualizar campos "suaves" si vienen valores
        df_update = df_out[df_out['numero_ficha'].isin(existing_ids)].copy() if existing_ids else pd.DataFrame(columns=df_out.columns)
        if not df_update.empty:
            # Nos interesan principalmente estrategia_programa y estado_curso
            df_update = df_update[
                (df_update.get('estrategia_programa').notna() if 'estrategia_programa' in df_update.columns else False)
                | (df_update.get('estado_curso').notna() if 'estado_curso' in df_update.columns else False)
            ]

        updated_fichas = 0
        if not df_update.empty:
            update_params = [
                {
                    'numero_ficha': int(row['numero_ficha']),
                    'estrategia_programa': clean_optional_text(row['estrategia_programa']) if 'estrategia_programa' in df_update.columns else None,
                    'estado_curso': clean_optional_text(row['estado_curso']) if 'estado_curso' in df_update.columns else None,
                }
                for _, row in df_update.iterrows()
                if pd.notna(row['numero_ficha'])
            ]
            if update_params:
                update_sql = text(
                    'UPDATE programas_formacion '
                    'SET '
                    '    estrategia_programa = COALESCE(:estrategia_programa, estrategia_programa), '
                    '    estado_curso = COALESCE(:estado_curso, estado_curso) '
                    'WHERE numero_ficha = :numero_ficha'
                )
                with engine.begin() as conn:
                    result = conn.execute(update_sql, update_params)
                updated_fichas = int(result.rowcount or 0)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'Error insertando/actualizando programas: {e}')

    return JSONResponse({
        'inserted': int(len(df_new)) if 'df_new' in locals() else 0,
        'updated_fichas': int(updated_fichas),
        'fecha_corte': str(fecha_corte_file),
    })


@app.post('/programas/upload-certificados')
async def upload_programas_certificados(file: UploadFile = File(...)):
    """Actualiza campo `certificado` en programas_formacion usando un Excel complementario y cruce por numero_ficha."""
    if not file.filename.lower().endswith(('.xls', '.xlsx')):
        raise HTTPException(status_code=400, detail='El archivo debe ser .xls o .xlsx')

    content = await file.read()
    df = read_excel_basic(content)
    if df.empty:
        raise HTTPException(status_code=400, detail='El Excel complementario no contiene filas')

    # Posibles encabezados para cruce y valor de certificados.
    ficha_aliases = [
        'numero_ficha', 'numero_de_ficha', 'identificador_ficha', 'codigo_ficha', 'cod_ficha', 'ficha'
    ]
    certificados_aliases = [
        'certificado', 'certificados', 'aprendices_certificados', 'total_aprendices_certificados', 'total_certificados'
    ]

    alias_pool = set(normalize_col_name(x) for x in ficha_aliases + certificados_aliases)
    norm_default_headers = [normalize_col_name(str(c)) for c in list(df.columns)]
    default_score = len(set(norm_default_headers).intersection(alias_pool))
    unnamed_count = sum(1 for c in norm_default_headers if c.startswith('unnamed:'))

    if default_score < 2 or unnamed_count >= 5:
        df_raw = read_excel_no_header(content)
        scan_limit = min(40, len(df_raw.index))
        best_idx = None
        best_score = -1

        # Para certificados nos basta encontrar la fila donde aparezca
        # claramente la columna de ficha (por ejemplo, "Ficha"). Usamos
        # solo los aliases de ficha para que no dependa de tener tambien
        # encabezados de certificados en esa fila.
        ficha_alias_norm = set(normalize_col_name(x) for x in ficha_aliases)
        for idx in range(scan_limit):
            row_values = [str(v) for v in df_raw.iloc[idx].tolist() if pd.notna(v)]
            row_norm = set(normalize_cols(row_values))
            # puntuacion basada SOLO en coincidencias con encabezados de ficha
            score = len(row_norm.intersection(ficha_alias_norm))
            if score > best_score:
                best_score = score
                best_idx = idx
        # Si encontramos al menos una coincidencia con los aliases de ficha,
        # usamos esa fila como encabezado real.
        if best_idx is not None and best_score >= 1:
            df = read_excel_with_header_row(content, int(best_idx))

    df.columns = normalize_cols(df.columns)

    ficha_col = get_first_existing_column(df, [normalize_col_name(x) for x in ficha_aliases])
    cert_col = get_first_existing_column(df, [normalize_col_name(x) for x in certificados_aliases])

    # Se requiere SIEMPRE una columna de ficha. Si no existe, no podemos cruzar.
    if not ficha_col:
        headers = ', '.join([str(c) for c in list(df.columns)[:25]])
        raise HTTPException(
            status_code=400,
            detail=(
                'No se encontro ninguna columna de ficha en el Excel complementario. '
                f'Se esperaban encabezados similares a: {ficha_aliases}. '
                f'Encabezados detectados: {headers}'
            ),
        )

    # Si no hay columna explicita de cantidad de certificados, asumimos que
    # cada fila representa 1 certificado por ficha. Esto encaja con archivos
    # donde viene una fila por aprendiz certificado.
    if not cert_col:
        df_cert = pd.DataFrame({
            'numero_ficha': pd.to_numeric(df[ficha_col], errors='coerce').astype('Int64'),
            'certificado': 1,
        })
    else:
        df_cert = pd.DataFrame({
            'numero_ficha': pd.to_numeric(df[ficha_col], errors='coerce').astype('Int64'),
            'certificado': pd.to_numeric(df[cert_col], errors='coerce').astype('Int64'),
        })

    df_cert = df_cert[df_cert['numero_ficha'].notna()].copy()
    if df_cert.empty:
        raise HTTPException(status_code=400, detail='No hay filas con numero de ficha valido en el Excel complementario')

    # Si hay fichas repetidas en el archivo, sumar certificados para consolidar.
    df_cert = df_cert.groupby('numero_ficha', as_index=False)['certificado'].sum(min_count=1)
    df_cert['certificado'] = df_cert['certificado'].fillna(0).astype('Int64')

    ficha_ids = [int(x) for x in df_cert['numero_ficha'].dropna().tolist()]
    if not ficha_ids:
        raise HTTPException(status_code=400, detail='No se pudieron obtener fichas para actualizar')

    check_sql = text('SELECT DISTINCT numero_ficha FROM programas_formacion WHERE numero_ficha IN :ids').bindparams(bindparam('ids', expanding=True))
    with engine.connect() as conn:
        existing = conn.execute(check_sql, {'ids': ficha_ids}).fetchall()
    existing_ids = set(int(r[0]) for r in existing if r and r[0] is not None)

    df_to_update = df_cert[df_cert['numero_ficha'].isin(existing_ids)].copy()
    unmatched_ids = [fid for fid in ficha_ids if fid not in existing_ids]

    if df_to_update.empty:
        raise HTTPException(
            status_code=400,
            detail='Ninguna ficha del archivo complementario coincide con la tabla programas_formacion',
        )

    update_params = [
        {
            'numero_ficha': int(row['numero_ficha']),
            'certificado': str(int(row['certificado'])) if pd.notna(row['certificado']) else None,
        }
        for _, row in df_to_update.iterrows()
    ]

    update_sql = text('UPDATE programas_formacion SET certificado = :certificado WHERE numero_ficha = :numero_ficha')
    with engine.begin() as conn:
        result = conn.execute(update_sql, update_params)

    return JSONResponse({
        'updated_rows': int(result.rowcount or 0),
        'updated_fichas': int(len(df_to_update)),
        'unmatched_fichas': int(len(unmatched_ids)),
        'unmatched_sample': unmatched_ids[:20],
    })


@app.get('/programas')
def get_programas(
    year: Optional[int] = None,
    municipio: Optional[str] = None,
    estrategia: Optional[str] = None,
    convenio: Optional[str] = None,
    page: int = 1,
    per_page: int = 30,
):
    # Paginacion: maximo 30 registros por pagina
    try:
        page = int(page)
    except Exception:
        page = 1
    try:
        per_page = int(per_page)
    except Exception:
        per_page = 30
    if page < 1:
        page = 1
    if per_page < 1:
        per_page = 30
    if per_page > 30:
        per_page = 30

    clauses = []
    params: dict = {}

    if year is not None:
        clauses.append('YEAR(fecha_corte) = :year')
        params['year'] = int(year)
    if municipio:
        clauses.append('LOWER(TRIM(ciudad_municipio)) = :municipio')
        params['municipio'] = municipio.strip().lower()
    if estrategia:
        clauses.append('LOWER(TRIM(estrategia_programa)) = :estrategia')
        params['estrategia'] = estrategia.strip().lower()
    if convenio:
        clauses.append('LOWER(TRIM(convenio)) = :convenio')
        params['convenio'] = convenio.strip().lower()

    where_sql = ''
    if clauses:
        where_sql = ' WHERE ' + ' AND '.join(clauses)

    count_sql = f'SELECT COUNT(*) AS total FROM programas_formacion{where_sql}'
    try:
        with engine.connect() as conn:
            total = conn.execute(text(count_sql), params).scalar() or 0
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'Error contando programas: {e}')

    offset = (page - 1) * per_page
    data_sql = (
        'SELECT * FROM programas_formacion'
        f'{where_sql} '
        'ORDER BY fecha_corte DESC, numero_ficha ASC, id ASC '
        'LIMIT :limit OFFSET :offset'
    )
    params_data = dict(params)
    params_data['limit'] = per_page
    params_data['offset'] = offset

    try:
        df = pd.read_sql(text(data_sql), con=engine, params=params_data)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'Error consultando programas: {e}')

    # JSON no soporta NaN/NaT/inf; convertir a None para serializar correctamente.
    if not df.empty:
        df = df.replace([float('inf'), float('-inf')], pd.NA)
        df = df.where(pd.notna(df), None)

    fecha_corte = None
    if not df.empty and 'fecha_corte' in df.columns:
        valid = pd.to_datetime(df['fecha_corte'], errors='coerce').dropna()
        if not valid.empty:
            fecha_corte = valid.max().date().isoformat()

    items = df.to_dict(orient='records')
    for row in items:
        for key, value in list(row.items()):
            try:
                if pd.isna(value):
                    row[key] = None
                    continue
            except Exception:
                pass
            if hasattr(value, 'isoformat') and not isinstance(value, str):
                try:
                    row[key] = value.isoformat()
                except Exception:
                    pass

    payload = {
        'items': items,
        'total': int(total),
        'fecha_corte': fecha_corte,
        'page': page,
        'per_page': per_page,
    }
    return JSONResponse(content=jsonable_encoder(payload))


class UpdateRequest(BaseModel):
    cod_fichas: List[int]
    periodo: Optional[int] = None
    oferta: Optional[str] = None
    tipo: Optional[str] = None


class FichaUpdate(BaseModel):
    cod_regional: Optional[int] = None
    regional: Optional[str] = None
    cod_municipio: Optional[int] = None
    municipio: Optional[str] = None
    cod_centro: Optional[int] = None
    centro_formacion: Optional[str] = None
    cod_programa: Optional[int] = None
    denominacion_programa: Optional[str] = None
    cod_ficha: Optional[int] = None
    estado_ficha: Optional[str] = None
    jornada: Optional[str] = None
    nivel_formacion: Optional[str] = None
    cupo: Optional[int] = None
    inscritos_primera_opcion: Optional[int] = None
    inscritos_segunda_opcion: Optional[int] = None
    oferta: Optional[str] = None
    tipo: Optional[str] = None
    perfil_ingreso: Optional[str] = None
    periodo: Optional[int] = None


@app.get('/fichas/{cod_ficha}')
def get_ficha(cod_ficha: int):
    try:
        # Construir la consulta con el id como entero para evitar problemas de parámetros con pymysql
        df = pd.read_sql(f"SELECT * FROM fichas_formacion WHERE cod_ficha = {int(cod_ficha)}", con=engine)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'Error al leer la base de datos: {e}')
    if df.empty:
        raise HTTPException(status_code=404, detail='Ficha no encontrada')
    return JSONResponse(df.iloc[0].to_dict())


@app.put('/fichas/{cod_ficha}')
def update_ficha(cod_ficha: int, payload: FichaUpdate):
    data = payload.dict(exclude_unset=True)
    if not data:
        raise HTTPException(status_code=400, detail='No hay campos para actualizar')

    # No permitir cambiar la PK cod_ficha a otro valor desde aquí
    if 'cod_ficha' in data:
        data.pop('cod_ficha')

    updates = {}
    for k, v in data.items():
        if v is None:
            updates[k] = None
        elif k == 'oferta':
            updates['oferta'] = normalize_oferta(v)
        elif k == 'tipo':
            updates['tipo'] = normalize_tipo(v)
        else:
            updates[k] = v

    set_parts = []
    params = {}
    for i, (k, v) in enumerate(updates.items()):
        param_name = f'val_{i}'
        set_parts.append(f"{k} = :{param_name}")
        params[param_name] = v

    params['id'] = cod_ficha
    sql = text(f"UPDATE fichas_formacion SET {', '.join(set_parts)} WHERE cod_ficha = :id")

    try:
        with engine.begin() as conn:
            result = conn.execute(sql, params)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'Error al actualizar la ficha: {e}')

    return JSONResponse({'updated_rows': result.rowcount})


@app.post('/fichas/update')
def update_fichas(req: UpdateRequest):
    if not req.cod_fichas:
        raise HTTPException(status_code=400, detail='Se requiere al menos un cod_ficha')

    updates = {}
    if req.periodo is not None:
        updates['periodo'] = int(req.periodo)
    if req.oferta is not None:
        updates['oferta'] = normalize_oferta(req.oferta)
    if req.tipo is not None:
        updates['tipo'] = normalize_tipo(req.tipo)

    if not updates:
        raise HTTPException(status_code=400, detail='No hay campos para actualizar')

    set_parts = []
    params = {}
    for i, (k, v) in enumerate(updates.items()):
        param_name = f'val_{i}'
        set_parts.append(f"{k} = :{param_name}")
        params[param_name] = v

    params['ids'] = req.cod_fichas
    sql = text(f"UPDATE fichas_formacion SET {', '.join(set_parts)} WHERE cod_ficha IN :ids").bindparams(bindparam('ids', expanding=True))

    try:
        with engine.begin() as conn:
            result = conn.execute(sql, params)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'Error al actualizar registros: {e}')

    return JSONResponse({'updated_rows': result.rowcount})
