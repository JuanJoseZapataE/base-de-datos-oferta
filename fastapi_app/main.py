# ...existing code...

from fastapi import FastAPI, UploadFile, File, HTTPException, Form
from fastapi.responses import JSONResponse, StreamingResponse
from fastapi.encoders import jsonable_encoder
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List, Optional
import pandas as pd
import io
import os
import math
from datetime import datetime, date, time
import re
import unicodedata
import xml.etree.ElementTree as ET
from dotenv import load_dotenv
# Cargar .env desde la carpeta del paquete (asegura carga aunque el cwd sea el padre)
load_dotenv(os.path.join(os.path.dirname(__file__), '.env'))
from sqlalchemy import create_engine, text, bindparam
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter




# URL de la base de datos: editar o usar la variable de entorno DATABASE_URL
# Ejemplo: mysql+pymysql://root:password@localhost/sena_oferta
DATABASE_URL = os.getenv("DATABASE_URL", "mysql+pymysql://root@127.0.0.1/Oferta")

engine = create_engine(DATABASE_URL)
app = FastAPI(title="Importador Excel -> MySQL (sena_oferta)")



# Endpoint para traer todos los programas filtrados (sin paginación)
@app.get('/programas/all')
def programas_all(
    year: Optional[str] = None,
    municipio: Optional[str] = None,
    centro: Optional[str] = None,
    nivel: Optional[str] = None,
    estrategia: Optional[str] = None,
    convenio: Optional[str] = None,
    vigencia: Optional[str] = None,
    numero_ficha: Optional[int] = None,
    search: Optional[str] = None,
    solo_certificados: Optional[str] = None,
):
    clauses = []
    params: dict = {}
    if year is not None:
        years = [y.strip() for y in str(year).split(',') if y.strip()]
        if years:
            if len(years) == 1:
                clauses.append('YEAR(fecha_corte) = :year_0')
            else:
                in_keys = []
                for i, val in enumerate(years):
                    key = f'year_{i}'
                    in_keys.append(f':{key}')
                    params[key] = int(val)
                clauses.append('YEAR(fecha_corte) IN (' + ','.join(in_keys) + ')')
            if 'year_0' not in params and years:
                params['year_0'] = int(years[0])
    if municipio:
        municipios = [m.strip().lower() for m in str(municipio).split(',') if m.strip()]
        if municipios:
            if len(municipios) == 1:
                clauses.append('LOWER(TRIM(ciudad_municipio)) = :municipio_0')
            else:
                in_keys = []
                for i, val in enumerate(municipios):
                    key = f'municipio_{i}'
                    in_keys.append(f':{key}')
                    params[key] = val
                clauses.append('LOWER(TRIM(ciudad_municipio)) IN (' + ','.join(in_keys) + ')')
            if 'municipio_0' not in params and municipios:
                params['municipio_0'] = municipios[0]
    if centro:
        centros = [c.strip().lower() for c in str(centro).split(',') if c.strip()]
        if centros:
            if len(centros) == 1:
                clauses.append('LOWER(TRIM(centro_formacion)) = :centro_0')
            else:
                in_keys = []
                for i, val in enumerate(centros):
                    key = f'centro_{i}'
                    in_keys.append(f':{key}')
                    params[key] = val
                clauses.append('LOWER(TRIM(centro_formacion)) IN (' + ','.join(in_keys) + ')')
            if 'centro_0' not in params and centros:
                params['centro_0'] = centros[0]
    if nivel:
        niveles = [n.strip().lower() for n in str(nivel).split(',') if n.strip()]
        if niveles:
            if len(niveles) == 1:
                clauses.append('LOWER(TRIM(nivel_formacion)) = :nivel_0')
            else:
                in_keys = []
                for i, val in enumerate(niveles):
                    key = f'nivel_{i}'
                    in_keys.append(f':{key}')
                    params[key] = val
                clauses.append('LOWER(TRIM(nivel_formacion)) IN (' + ','.join(in_keys) + ')')
            if 'nivel_0' not in params and niveles:
                params['nivel_0'] = niveles[0]
    if estrategia:
        estrategias = [e.strip().lower() for e in str(estrategia).split(',') if e.strip()]
        if estrategias:
            if len(estrategias) == 1:
                clauses.append('LOWER(TRIM(estrategia_programa)) = :estrategia_0')
            else:
                in_keys = []
                for i, val in enumerate(estrategias):
                    key = f'estrategia_{i}'
                    in_keys.append(f':{key}')
                    params[key] = val
                clauses.append('LOWER(TRIM(estrategia_programa)) IN (' + ','.join(in_keys) + ')')
            if 'estrategia_0' not in params and estrategias:
                params['estrategia_0'] = estrategias[0]
    if convenio:
        convenios = [c.strip().lower() for c in str(convenio).split(',') if c.strip()]
        if convenios:
            if len(convenios) == 1:
                clauses.append('LOWER(TRIM(convenio)) = :convenio_0')
            else:
                in_keys = []
                for i, val in enumerate(convenios):
                    key = f'convenio_{i}'
                    in_keys.append(f':{key}')
                    params[key] = val
                clauses.append('LOWER(TRIM(convenio)) IN (' + ','.join(in_keys) + ')')
            if 'convenio_0' not in params and convenios:
                params['convenio_0'] = convenios[0]
    if vigencia is not None:
        vigencias = [v.strip() for v in str(vigencia).split(',') if v.strip()]
        if vigencias:
            if len(vigencias) == 1:
                clauses.append('YEAR(fecha_inicio) = :vigencia_0')
            else:
                in_keys = []
                for i, val in enumerate(vigencias):
                    key = f'vigencia_{i}'
                    in_keys.append(f':{key}')
                    params[key] = int(val)
                clauses.append('YEAR(fecha_inicio) IN (' + ','.join(in_keys) + ')')
            if 'vigencia_0' not in params and vigencias:
                params['vigencia_0'] = int(vigencias[0])
    if numero_ficha is not None:
        clauses.append('numero_ficha = :numero_ficha')
        params['numero_ficha'] = int(numero_ficha)
    if search:
        s = str(search).strip().lower()
        if s:
            clauses.append('LOWER(TRIM(denominacion_programa)) LIKE :search')
            params['search'] = f'%{s}%'
    if solo_certificados and str(solo_certificados).strip().lower() not in {'0', 'false', 'no'}:
        clauses.append('(certificado IS NOT NULL AND certificado <> 0)')

    where_sql = ''
    if clauses:
        where_sql = ' WHERE ' + ' AND '.join(clauses)

    sql = f'SELECT * FROM programas_formacion{where_sql} ORDER BY fecha_corte DESC, numero_ficha ASC, id ASC'
    try:
        df = pd.read_sql(text(sql), con=engine, params=params)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'Error consultando programas: {e}')

    data = []
    if not df.empty:
        # Limpiar infinitos primero
        df = df.replace([float('inf'), float('-inf')], pd.NA)

        # Convertir columnas de fecha/tiempo a cadenas ISO para que sean JSON serializables
        for col in ['fecha_inicio', 'fecha_fin', 'fecha_corte']:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
                df[col] = df[col].apply(
                    lambda v: v.isoformat() if hasattr(v, 'isoformat') else v
                )

        # Pasar a lista de dicts y reemplazar NaN/inf por None para que JSON los acepte
        raw_records = df.to_dict(orient='records')
        cleaned_records = []
        for row in raw_records:
            for key, value in list(row.items()):
                if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
                    row[key] = None
            cleaned_records.append(row)
        data = cleaned_records

    return JSONResponse(data)




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


INDICATIVA_COLUMNS = [
    'id_indicativa',
    'regional',
    'codigo_de_centro',
    'nombre_sede',
    'vigencia',
    'periodo_oferta',
    'codigo_programa',
    'version',
    'codigo_version',
    'nombre_programa',
    'nivel_de_formacion',
    'modalidad',
    'mes_inicio',
    'cupos',
    'ano_termina',
    'departamento_formacion',
    'codigo_dane_departamento',
    'municipio_formacion',
    'codigo_dane_municipio',
    'gira_tecnica',
    'programa_fic',
    'tipo_de_oferta',
    'persona_registra',
    'fecha_de_registro',
    'tipo_de_institucion',
    'nivel_institucion',
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


def ensure_indicativa_table():
    create_sql = """
    CREATE TABLE IF NOT EXISTS indicativa (
        id BIGINT NOT NULL AUTO_INCREMENT PRIMARY KEY,
        id_indicativa BIGINT NULL,
        regional VARCHAR(150) NULL,
        codigo_de_centro INT NULL,
        nombre_sede VARCHAR(255) NULL,
        vigencia INT NULL,
        periodo_oferta VARCHAR(100) NULL,
        codigo_programa BIGINT NULL,
        version INT NULL,
        codigo_version VARCHAR(50) NULL,
        nombre_programa VARCHAR(255) NULL,
        nivel_de_formacion VARCHAR(150) NULL,
        modalidad VARCHAR(150) NULL,
        mes_inicio VARCHAR(50) NULL,
        cupos INT NULL,
        ano_termina INT NULL,
        departamento_formacion VARCHAR(150) NULL,
        codigo_dane_departamento VARCHAR(20) NULL,
        municipio_formacion VARCHAR(150) NULL,
        codigo_dane_municipio VARCHAR(20) NULL,
        gira_tecnica VARCHAR(50) NULL,
        programa_fic VARCHAR(50) NULL,
        tipo_de_oferta VARCHAR(150) NULL,
        persona_registra VARCHAR(150) NULL,
        fecha_de_registro DATETIME NULL,
        tipo_de_institucion VARCHAR(150) NULL,
        nivel_institucion VARCHAR(150) NULL,
        INDEX idx_indicativa_vigencia (vigencia),
        INDEX idx_indicativa_periodo (periodo_oferta),
        INDEX idx_indicativa_centro (nombre_sede)
    )
    """
    with engine.begin() as conn:
        conn.execute(text(create_sql))


ensure_indicativa_table()


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


def export_header_label_indicativa(column_name: str) -> str:
    if not column_name:
        return ''
    mapping = {
        'nombre_sede': 'Centro de formacion',
        'nivel_de_formacion': 'Nivel de formacion',
        'nombre_programa': 'Denominacion del programa',
        'periodo_oferta': 'Periodo oferta',
        'tipo_de_oferta': 'Tipo oferta',
    }
    if column_name in mapping:
        return mapping[column_name]
    return export_header_label(column_name)


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


def read_spreadsheetml_xml(content: bytes) -> pd.DataFrame:
    """Lee un archivo XML de Excel 2003 (SpreadsheetML) como tabla.

    Extrae la primera hoja (Worksheet/Table) y construye un DataFrame usando
    la primera fila como encabezados y el resto como filas de datos.
    """
    try:
        root = ET.fromstring(content)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f'XML de Excel invalido: {e}')

    ws = root.find('.//{*}Worksheet')
    if ws is None:
        raise HTTPException(status_code=400, detail='No se encontro ningun Worksheet en el XML de Excel.')

    table = ws.find('.//{*}Table')
    if table is None:
        raise HTTPException(status_code=400, detail='No se encontro ninguna tabla (Table) en el XML de Excel.')

    rows_raw = []
    for row in table.findall('.//{*}Row'):
        # Soportar celdas con atributo ss:Index (saltos de columnas).
        cells: list[str] = []
        col_pos = 0
        for cell in row.findall('.//{*}Cell'):
            idx_attr = None
            for attr_name, attr_val in cell.attrib.items():
                if attr_name.endswith('Index'):
                    idx_attr = attr_val
                    break
            if idx_attr is not None:
                try:
                    col_pos = int(idx_attr) - 1
                except Exception:
                    pass

            data_el = cell.find('.//{*}Data')
            text = '' if data_el is None or data_el.text is None else str(data_el.text)

            if len(cells) <= col_pos:
                cells.extend([''] * (col_pos - len(cells)))
                cells.append(text)
            else:
                cells[col_pos] = text

            col_pos += 1

        # Ignorar filas completamente vacias
        if any(val.strip() for val in cells):
            rows_raw.append(cells)

    if not rows_raw:
        return pd.DataFrame()

    # Detectar la fila que realmente contiene los encabezados (no el titulo tipo "PE-04_").
    # Usamos palabras clave tipicas de tus archivos: IDENTIFICADOR_FICHA, NUMERO_FICHA,
    # NOMBRE_PROGRAMA_FORMACION, etc.
    header_aliases = [
        'identificador_ficha', 'numero_ficha', 'n_ficha', 'codigo_ficha', 'cod_ficha',
        'nombre_programa_formacion', 'denominacion_programa',
        'centro_formacion', 'ciudad_municipio', 'nivel_formacion',
    ]
    alias_pool = set(normalize_col_name(a) for a in header_aliases)

    best_idx = 0
    best_score = -1
    scan_limit = min(40, len(rows_raw))
    for idx in range(scan_limit):
        row = rows_raw[idx]
        # normalizar cada celda como si fuera nombre de columna
        norm_cells = set(normalize_col_name(c) for c in row if c is not None)
        score = len(norm_cells.intersection(alias_pool))
        if score > best_score:
            best_score = score
            best_idx = idx

    header = rows_raw[best_idx]
    data_rows = rows_raw[best_idx + 1 :]
    num_cols = len(header)

    normalized_rows = []
    for r in data_rows:
        if len(r) < num_cols:
            r = r + [''] * (num_cols - len(r))
        elif len(r) > num_cols:
            r = r[:num_cols]
        normalized_rows.append(r)

    df = pd.DataFrame(normalized_rows, columns=header)
    return df


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


def extract_fecha_corte_from_excel_content(content: bytes):
    """Intenta obtener fecha de corte desde el contenido del archivo Excel.

    Caso de uso: archivos que no traen fecha en el nombre pero si en una
    celda fija, por ejemplo A3 (fila 3, columna 1).
    """
    # Leer el libro con openpyxl para acceder directamente a la celda A3
    try:
        wb = load_workbook(io.BytesIO(content), data_only=True)
    except Exception:
        return None

    try:
        ws = wb.active
    except Exception:
        return None

    try:
        cell = ws['A3']
    except Exception:
        return None

    value = cell.value
    if value is None:
        return None

    # 1) Si ya viene como datetime/fecha de Excel, usarla directamente.
    if isinstance(value, datetime):
        return value.date()

    # 2) Intentar parseo generico con pandas (por si es texto tipo "2022-12-01").
    s = str(value).strip()
    try:
        ts = pd.to_datetime(s, dayfirst=True, errors='coerce')
    except Exception:
        ts = None
    if ts is not None and not pd.isna(ts):
        return ts.date()

    # 3) Buscar patrones numericos dentro del texto.
    #    Primero AAAAMMDD (8 digitos), luego AAAAMM (6 digitos).
    m8 = re.search(r'(\d{8})', s)
    if m8:
        raw = m8.group(1)
        try:
            dt = datetime.strptime(raw, '%Y%m%d')
            return dt.date()
        except Exception:
            pass

    m6 = re.search(r'(\d{6})', s)
    if m6:
        raw = m6.group(1)
        try:
            year = int(raw[:4])
            month = int(raw[4:6])
            dt = datetime(year, month, 1)
            return dt.date()
        except Exception:
            pass
    return None


def _parse_excel_fecha_value(value) -> Optional[date]:
    """Intenta convertir un valor de celda de Excel a date, asumiendo formato dia/mes/año cuando es texto."""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    s = str(value).strip()
    if not s:
        return None
    # Intentar formato explicito dd/mm/yyyy primero (requerido por el usuario)
    for fmt in ('%d/%m/%Y', '%d-%m-%Y', '%Y-%m-%d', '%Y/%m/%d'):
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            continue
    # Fallback generico usando pandas (acepta mas variantes)
    try:
        ts = pd.to_datetime(s, dayfirst=True, errors='coerce')
        if ts is not None and not pd.isna(ts):
            return ts.date()
    except Exception:
        pass
    # Patrones numericos: ddmmyyyy o yyyymmdd dentro del texto
    m8 = re.search(r'(\d{8})', s)
    if m8:
        raw = m8.group(1)
        for fmt in ('%d%m%Y', '%Y%m%d'):
            try:
                return datetime.strptime(raw, fmt).date()
            except Exception:
                continue
    return None


def _parse_excel_hora_value(value) -> Optional[time]:
    """Intenta convertir un valor de celda de Excel a time (HH:MM[:SS])."""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.time().replace(microsecond=0)
    if isinstance(value, time):
        return value.replace(microsecond=0)
    s = str(value).strip()
    if not s:
        return None
    # Intentar formatos comunes de hora
    for fmt in ('%H:%M', '%H:%M:%S'):
        try:
            return datetime.strptime(s, fmt).time().replace(microsecond=0)
        except Exception:
            continue
    # Fallback: si viene como numero de Excel (fraccion del dia), intentar convertir
    try:
        # Excel almacena horas como fraccion del dia; multiplicar por 24 para horas
        num = float(s)
        total_seconds = int(round(num * 24 * 3600))
        hh = (total_seconds // 3600) % 24
        mm = (total_seconds % 3600) // 60
        ss = total_seconds % 60
        return time(hour=hh, minute=mm, second=ss)
    except Exception:
        return None


def extract_fecha_reporte_from_filename_fichas(filename: str) -> Optional[date]:
    """Extrae fecha de reporte desde nombre tipo CCX_17032026.xlsx o CCX_17-03-2026.xlsx.

    Estructura esperada: (siglas_centro)_(fecha_reporte)
    """
    if not filename:
        return None
    name = os.path.splitext(os.path.basename(filename))[0]
    parts = name.split('_')
    candidate = None
    if len(parts) >= 2:
        candidate = parts[1]
    else:
        # Si no hay guion bajo, buscar bloque de 8 digitos en todo el nombre
        m = re.search(r'(\d{8})', name)
        if m:
            candidate = m.group(1)
    if not candidate:
        return None
    s = str(candidate).strip()
    # Normalizar separadores
    s_norm = s.replace('-', '/').replace('.', '/').replace(' ', '/')
    # Intentar dd/mm/yyyy
    try:
        return datetime.strptime(s_norm, '%d/%m/%Y').date()
    except Exception:
        pass
    # Intentar ddmmyyyy (sin separadores)
    if re.fullmatch(r'\d{8}', s):
        for fmt in ('%d%m%Y', '%Y%m%d'):
            try:
                return datetime.strptime(s, fmt).date()
            except Exception:
                continue
    return None


def extract_fecha_hora_reporte_fichas(content: bytes, filename: str):
    """Obtiene fecha (B4) y hora (B5) del Excel de fichas o, si falta la fecha, del nombre del archivo.

    - B4: fecha de reporte en formato dia/mes/año (preferido).
    - B5: hora de reporte (HH:MM u hora de Excel).
    - Si B4 no tiene valor interpretable, se intenta extraer la fecha desde el nombre.
    """
    fecha: Optional[date] = None
    hora: Optional[time] = None

    # Primero intentar leer directamente desde el contenido del Excel
    try:
        wb = load_workbook(io.BytesIO(content), data_only=True)
        ws = wb.active
        try:
            fecha_val = ws['B4'].value
            hora_val = ws['B5'].value
        except Exception:
            fecha_val = None
            hora_val = None
        fecha = _parse_excel_fecha_value(fecha_val)
        hora = _parse_excel_hora_value(hora_val)
    except Exception:
        # Si no se puede abrir el libro, se intentara solo por nombre
        pass

    # Si la fecha sigue sin definirse, usar nombre del archivo
    if fecha is None:
        fecha = extract_fecha_reporte_from_filename_fichas(filename or '')

    return fecha, hora


@app.post('/upload-excel')
async def upload_excel(file: UploadFile = File(...), periodo: Optional[int] = Form(None), oferta: Optional[str] = Form(None), tipo: Optional[str] = Form(None)):
    if not file.filename.lower().endswith(('.xls', '.xlsx', '.xml')):
        raise HTTPException(status_code=400, detail='El archivo debe ser .xls, .xlsx o .xml')

    content = await file.read()

    # Extraer fecha y hora de reporte desde el Excel (B4/B5) o, si falta la fecha,
    # desde el nombre del archivo. Esto permite saber de que corte es el archivo
    # que se esta subiendo en el modulo de fichas.
    fecha_reporte, hora_reporte = extract_fecha_hora_reporte_fichas(content, file.filename or '')

    # Si es XML, intentar leerlo como tabla antes de aplicar la logica de deteccion
    # de encabezados pensada para Excel.
    if file.filename.lower().endswith('.xml'):
        try:
            df = pd.read_xml(io.BytesIO(content))
        except Exception as e:
            raise HTTPException(
                status_code=400,
                detail=f'No se pudo leer el XML como tabla: {e}',
            )
    else:
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

    # Incluir en la respuesta la fecha y hora de reporte detectadas para
    # que el frontend pueda mostrar de que archivo/corte se trata.
    fecha_str = fecha_reporte.strftime('%d/%m/%Y') if isinstance(fecha_reporte, date) else None
    hora_str = hora_reporte.strftime('%H:%M:%S') if isinstance(hora_reporte, time) else None

    return JSONResponse({'inserted': len(df_to_insert), 'fecha_reporte': fecha_str, 'hora_reporte': hora_str})


@app.post('/indicativa/upload-excel')
async def upload_indicativa_excel(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(('.xls', '.xlsx', '.xml')):
        raise HTTPException(status_code=400, detail='El archivo debe ser .xls, .xlsx o .xml')

    content = await file.read()

    # Permitir XML, igual que en otros modulos
    if file.filename.lower().endswith('.xml'):
        try:
            df = pd.read_xml(io.BytesIO(content))
        except Exception as e:
            raise HTTPException(status_code=400, detail=f'No se pudo leer el XML de indicativa como tabla: {e}')
    else:
        df = read_excel_basic(content)

    if df.empty:
        raise HTTPException(status_code=400, detail='El Excel no contiene filas')

    # Normalizar nombres de columnas (quita acentos, pasa a minusculas, reemplaza espacios por _)
    df.columns = normalize_cols(df.columns)

    # Asegurar todas las columnas esperadas
    df_to_insert = pd.DataFrame()
    for col in INDICATIVA_COLUMNS:
        if col in df.columns:
            df_to_insert[col] = df[col]
        else:
            df_to_insert[col] = None

    # Tipos basicos
    for col in ['codigo_de_centro', 'vigencia', 'version', 'cupos', 'ano_termina', 'codigo_programa', 'id_indicativa']:
        if col in df_to_insert.columns:
            df_to_insert[col] = pd.to_numeric(df_to_insert[col], errors='coerce').astype('Int64')

    # fecha_de_registro puede venir como texto o fecha de Excel
    if 'fecha_de_registro' in df_to_insert.columns:
        try:
            df_to_insert['fecha_de_registro'] = pd.to_datetime(df_to_insert['fecha_de_registro'], errors='coerce')
        except Exception:
            pass

    # Eliminar filas completamente vacias en campos clave basicos (nombre_sede y nombre_programa)
    key_fields = ['nombre_sede', 'nombre_programa']
    df_to_insert = df_to_insert[~df_to_insert[key_fields].isna().all(axis=1)].copy()
    if df_to_insert.empty:
        raise HTTPException(status_code=400, detail='No se encontraron filas validas para insertar en indicativa')

    try:
        df_to_insert.to_sql('indicativa', con=engine, if_exists='append', index=False)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'Error al insertar en la tabla indicativa: {e}')

    return JSONResponse({'inserted': int(len(df_to_insert))})


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


@app.get('/indicativa')
def get_indicativa(
    page: int = 1,
    per_page: int = 50,
    centro: Optional[str] = None,
    nivel: Optional[str] = None,
    periodo_oferta: Optional[str] = None,
    municipio: Optional[str] = None,
    search: Optional[str] = None,
):
    """Listado paginado de la tabla indicativa para el frontend, con filtros opcionales."""
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
    if per_page < 1 or per_page > 200:
        per_page = 50

    # Construir filtros
    clauses = []
    params: dict = {}
    if centro:
        centros = [c.strip().lower() for c in str(centro).split(',') if c.strip()]
        if centros:
            if len(centros) == 1:
                clauses.append('LOWER(TRIM(nombre_sede)) = :centro_0')
            else:
                in_keys = []
                for i, val in enumerate(centros):
                    key = f'centro_{i}'
                    in_keys.append(f':{key}')
                    params[key] = val
                clauses.append('LOWER(TRIM(nombre_sede)) IN (' + ','.join(in_keys) + ')')
            if 'centro_0' not in params and centros:
                params['centro_0'] = centros[0]
    if nivel:
        niveles = [n.strip().lower() for n in str(nivel).split(',') if n.strip()]
        if niveles:
            if len(niveles) == 1:
                clauses.append('LOWER(TRIM(nivel_de_formacion)) = :nivel_0')
            else:
                in_keys = []
                for i, val in enumerate(niveles):
                    key = f'nivel_{i}'
                    in_keys.append(f':{key}')
                    params[key] = val
                clauses.append('LOWER(TRIM(nivel_de_formacion)) IN (' + ','.join(in_keys) + ')')
            if 'nivel_0' not in params and niveles:
                params['nivel_0'] = niveles[0]
    if periodo_oferta:
        periodos = [p.strip().lower() for p in str(periodo_oferta).split(',') if p.strip()]
        if periodos:
            if len(periodos) == 1:
                clauses.append('LOWER(TRIM(periodo_oferta)) = :periodo_oferta_0')
            else:
                in_keys = []
                for i, val in enumerate(periodos):
                    key = f'periodo_oferta_{i}'
                    in_keys.append(f':{key}')
                    params[key] = val
                clauses.append('LOWER(TRIM(periodo_oferta)) IN (' + ','.join(in_keys) + ')')
            if 'periodo_oferta_0' not in params and periodos:
                params['periodo_oferta_0'] = periodos[0]
    if municipio:
        municipios = [m.strip().lower() for m in str(municipio).split(',') if m.strip()]
        if municipios:
            if len(municipios) == 1:
                clauses.append('LOWER(TRIM(municipio_formacion)) = :municipio_0')
            else:
                in_keys = []
                for i, val in enumerate(municipios):
                    key = f'municipio_{i}'
                    in_keys.append(f':{key}')
                    params[key] = val
                clauses.append('LOWER(TRIM(municipio_formacion)) IN (' + ','.join(in_keys) + ')')
            if 'municipio_0' not in params and municipios:
                params['municipio_0'] = municipios[0]
    if search:
        s = str(search).strip().lower()
        if s:
            clauses.append('LOWER(TRIM(nombre_programa)) LIKE :search')
            params['search'] = f'%{s}%'

    where_sql = ''
    if clauses:
        where_sql = ' WHERE ' + ' AND '.join(clauses)

    # Contar total
    count_sql = f'SELECT COUNT(*) FROM indicativa{where_sql}'
    try:
        with engine.connect() as conn:
            total = conn.execute(text(count_sql), params).scalar() or 0
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'Error contando registros de indicativa: {e}')

    offset = (page - 1) * per_page
    sql = (
        'SELECT id, nombre_sede, municipio_formacion, nivel_de_formacion, nombre_programa, '
        'periodo_oferta, tipo_de_oferta '
        'FROM indicativa'
        f'{where_sql} '
        'ORDER BY vigencia DESC, periodo_oferta ASC, nombre_sede ASC '
        'LIMIT :limit OFFSET :offset'
    )
    params_data = dict(params)
    params_data['limit'] = per_page
    params_data['offset'] = offset
    try:
        df = pd.read_sql(text(sql), con=engine, params=params_data)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'Error consultando indicativa: {e}')

    if not df.empty:
        df = df.replace([float('inf'), float('-inf')], pd.NA)
        df = df.where(pd.notna(df), None)

    items = df.to_dict(orient='records') if not df.empty else []

    # Renombrar claves para que ya vayan con los nombres que usara el frontend
    mapped_items = []
    for row in items:
        mapped_items.append(
            {
                'id': row.get('id'),
                'centro_formacion': row.get('nombre_sede'),
                'municipio_formacion': row.get('municipio_formacion'),
                'nivel_formacion': row.get('nivel_de_formacion'),
                'denominacion_programa': row.get('nombre_programa'),
                'periodo_oferta': row.get('periodo_oferta'),
                'tipo_oferta': row.get('tipo_de_oferta'),
            }
        )

    # Asegurar que no queden NaN/inf en los datos antes de serializar a JSON
    cleaned_items = []
    for row in mapped_items:
        for key, value in list(row.items()):
            if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
                row[key] = None
        cleaned_items.append(row)

    return JSONResponse(
        content=jsonable_encoder(
            {
                'items': cleaned_items,
                'total': int(total),
                'page': page,
                'per_page': per_page,
            }
        )
    )


@app.get('/indicativa/export')
def export_indicativa_excel(
    centro: Optional[str] = None,
    nivel: Optional[str] = None,
    periodo_oferta: Optional[str] = None,
    municipio: Optional[str] = None,
    search: Optional[str] = None,
):
    """Exporta Excel de la tabla indicativa respetando los filtros activos."""
    clauses = []
    params: dict = {}

    if centro:
        centros = [c.strip().lower() for c in str(centro).split(',') if c.strip()]
        if centros:
            if len(centros) == 1:
                clauses.append('LOWER(TRIM(nombre_sede)) = :centro_0')
            else:
                in_keys = []
                for i, val in enumerate(centros):
                    key = f'centro_{i}'
                    in_keys.append(f':{key}')
                    params[key] = val
                clauses.append('LOWER(TRIM(nombre_sede)) IN (' + ','.join(in_keys) + ')')
            if 'centro_0' not in params and centros:
                params['centro_0'] = centros[0]
    if nivel:
        niveles = [n.strip().lower() for n in str(nivel).split(',') if n.strip()]
        if niveles:
            if len(niveles) == 1:
                clauses.append('LOWER(TRIM(nivel_de_formacion)) = :nivel_0')
            else:
                in_keys = []
                for i, val in enumerate(niveles):
                    key = f'nivel_{i}'
                    in_keys.append(f':{key}')
                    params[key] = val
                clauses.append('LOWER(TRIM(nivel_de_formacion)) IN (' + ','.join(in_keys) + ')')
            if 'nivel_0' not in params and niveles:
                params['nivel_0'] = niveles[0]
    if periodo_oferta:
        periodos = [p.strip().lower() for p in str(periodo_oferta).split(',') if p.strip()]
        if periodos:
            if len(periodos) == 1:
                clauses.append('LOWER(TRIM(periodo_oferta)) = :periodo_oferta_0')
            else:
                in_keys = []
                for i, val in enumerate(periodos):
                    key = f'periodo_oferta_{i}'
                    in_keys.append(f':{key}')
                    params[key] = val
                clauses.append('LOWER(TRIM(periodo_oferta)) IN (' + ','.join(in_keys) + ')')
            if 'periodo_oferta_0' not in params and periodos:
                params['periodo_oferta_0'] = periodos[0]
    if municipio:
        municipios = [m.strip().lower() for m in str(municipio).split(',') if m.strip()]
        if municipios:
            if len(municipios) == 1:
                clauses.append('LOWER(TRIM(municipio_formacion)) = :municipio_0')
            else:
                in_keys = []
                for i, val in enumerate(municipios):
                    key = f'municipio_{i}'
                    in_keys.append(f':{key}')
                    params[key] = val
                clauses.append('LOWER(TRIM(municipio_formacion)) IN (' + ','.join(in_keys) + ')')
            if 'municipio_0' not in params and municipios:
                params['municipio_0'] = municipios[0]
    if search:
        s = str(search).strip().lower()
        if s:
            clauses.append('LOWER(TRIM(nombre_programa)) LIKE :search')
            params['search'] = f'%{s}%'

    where_sql = ''
    if clauses:
        where_sql = ' WHERE ' + ' AND '.join(clauses)

    # Exportar siempre todas las columnas de la tabla indicativa.
    sql = (
        'SELECT * FROM indicativa'
        f'{where_sql} '
        'ORDER BY vigencia DESC, periodo_oferta ASC, nombre_sede ASC'
    )

    try:
        df = pd.read_sql(text(sql), con=engine, params=params)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'Error al exportar indicativa: {e}')

    df_export = df.copy()

    # Encabezados legibles para Excel.
    original_cols = list(df_export.columns)
    df_export.columns = [export_header_label_indicativa(col) for col in df_export.columns]

    # Columnas que SI se ven en el frontend (resto deben quedar ocultas en Excel).
    visible_db_cols = {
        'nombre_sede',
        'nivel_de_formacion',
        'nombre_programa',
        'periodo_oferta',
        'tipo_de_oferta',
    }
    hidden_db_cols = {c for c in original_cols if c not in visible_db_cols}
    hidden_header_labels = {export_header_label_indicativa(c) for c in hidden_db_cols}

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name='indicativa')

        ws = writer.book['indicativa']
        max_row = ws.max_row
        max_col = ws.max_column

        wrap_alignment = Alignment(wrap_text=True, vertical='top')
        for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
            for cell in row:
                cell.alignment = wrap_alignment

        for cell in ws[1]:
            cell.font = Font(bold=True)

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

            # Ocultar en Excel las columnas que no se ven en la tabla del frontend.
            header_value = ws.cell(row=1, column=col_idx).value
            if header_value in hidden_header_labels:
                ws.column_dimensions[col_letter].hidden = True

        if max_col >= 1 and max_row >= 1:
            last_col_letter = get_column_letter(max_col)
            table_ref = f'A1:{last_col_letter}{max_row}'
            table = Table(displayName='IndicativaExport', ref=table_ref)
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
    filename = f'indicativa_export_{ts}.xlsx'

    return StreamingResponse(
        output,
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={'Content-Disposition': f'attachment; filename="{filename}"'},
    )


@app.get('/indicativa/filters')
def get_indicativa_filters():
    """Devuelve valores distintos para los filtros de indicativa."""
    try:
        with engine.connect() as conn:
            centros = [
                str(r[0])
                for r in conn.execute(
                    text('SELECT DISTINCT nombre_sede FROM indicativa WHERE nombre_sede IS NOT NULL ORDER BY nombre_sede ASC')
                ).fetchall()
                if r[0] is not None
            ]
            niveles = [
                str(r[0])
                for r in conn.execute(
                    text('SELECT DISTINCT nivel_de_formacion FROM indicativa WHERE nivel_de_formacion IS NOT NULL ORDER BY nivel_de_formacion ASC')
                ).fetchall()
                if r[0] is not None
            ]
            periodos = [
                str(r[0])
                for r in conn.execute(
                    text('SELECT DISTINCT periodo_oferta FROM indicativa WHERE periodo_oferta IS NOT NULL ORDER BY periodo_oferta ASC')
                ).fetchall()
                if r[0] is not None
            ]
            municipios = [
                str(r[0])
                for r in conn.execute(
                    text('SELECT DISTINCT municipio_formacion FROM indicativa WHERE municipio_formacion IS NOT NULL ORDER BY municipio_formacion ASC')
                ).fetchall()
                if r[0] is not None
            ]
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'Error obteniendo filtros de indicativa: {e}')

    return JSONResponse(
        content=jsonable_encoder(
            {
                'centros': centros,
                'niveles': niveles,
                'periodos_oferta': periodos,
                'municipios': municipios,
            }
        )
    )


@app.delete('/indicativa/delete-all')
def delete_indicativa_all():
    """Elimina todos los registros de la tabla indicativa."""
    try:
        with engine.begin() as conn:
            result = conn.execute(text('DELETE FROM indicativa'))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'Error eliminando todos los registros de indicativa: {e}')

    return JSONResponse({'deleted_rows': int(result.rowcount or 0)})


@app.delete('/indicativa/{indicativa_id}')
def delete_indicativa_by_id(indicativa_id: int):
    """Elimina un registro de indicativa por su id."""
    try:
        with engine.begin() as conn:
            exists = conn.execute(
                text('SELECT COUNT(*) FROM indicativa WHERE id = :id'),
                {'id': int(indicativa_id)},
            ).scalar() or 0
            if int(exists) == 0:
                raise HTTPException(status_code=404, detail='Registro de indicativa no encontrado')

            result = conn.execute(
                text('DELETE FROM indicativa WHERE id = :id'),
                {'id': int(indicativa_id)},
            )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'Error eliminando registro de indicativa: {e}')

    return JSONResponse({'deleted_rows': int(result.rowcount or 0), 'id': int(indicativa_id)})


@app.get('/fichas/export')
def export_fichas_excel(
    centro: Optional[str] = None,
    oferta: Optional[str] = None,
    estado: Optional[str] = None,
    tipo: Optional[str] = None,
    nivel: Optional[str] = None,
    periodo: Optional[str] = None,
    search: Optional[str] = None,
):
    """Exporta Excel con los filtros activos (mismos criterios del frontend)."""
    clauses = []
    params = {}

    if centro:
        centros = [c.strip().lower() for c in str(centro).split(',') if c.strip()]
        if centros:
            if len(centros) == 1:
                clauses.append('LOWER(TRIM(centro_formacion)) = :centro_0')
            else:
                in_keys = []
                for i, val in enumerate(centros):
                    key = f'centro_{i}'
                    in_keys.append(f':{key}')
                    params[key] = val
                clauses.append('LOWER(TRIM(centro_formacion)) IN (' + ','.join(in_keys) + ')')
            if 'centro_0' not in params and centros:
                params['centro_0'] = centros[0]
    if oferta:
        ofertas = [normalize_oferta(o) for o in str(oferta).split(',') if str(o).strip()]
        if ofertas:
            if len(ofertas) == 1:
                clauses.append('oferta = :oferta_0')
            else:
                in_keys = []
                for i, val in enumerate(ofertas):
                    key = f'oferta_{i}'
                    in_keys.append(f':{key}')
                    params[key] = val
                clauses.append('oferta IN (' + ','.join(in_keys) + ')')
            if 'oferta_0' not in params and ofertas:
                params['oferta_0'] = ofertas[0]
    if estado:
        estados = [e.strip().lower() for e in str(estado).split(',') if e.strip()]
        if estados:
            if len(estados) == 1:
                clauses.append('LOWER(TRIM(estado_ficha)) = :estado_0')
            else:
                in_keys = []
                for i, val in enumerate(estados):
                    key = f'estado_{i}'
                    in_keys.append(f':{key}')
                    params[key] = val
                clauses.append('LOWER(TRIM(estado_ficha)) IN (' + ','.join(in_keys) + ')')
            if 'estado_0' not in params and estados:
                params['estado_0'] = estados[0]
    if tipo:
        tipos = [normalize_tipo(t).strip().lower() for t in str(tipo).split(',') if t.strip()]
        if tipos:
            if len(tipos) == 1:
                clauses.append('LOWER(TRIM(tipo)) = :tipo_0')
            else:
                in_keys = []
                for i, val in enumerate(tipos):
                    key = f'tipo_{i}'
                    in_keys.append(f':{key}')
                    params[key] = val
                clauses.append('LOWER(TRIM(tipo)) IN (' + ','.join(in_keys) + ')')
            if 'tipo_0' not in params and tipos:
                params['tipo_0'] = tipos[0]
    if nivel:
        niveles = [n.strip().lower() for n in str(nivel).split(',') if n.strip()]
        if niveles:
            if len(niveles) == 1:
                clauses.append('LOWER(TRIM(nivel_formacion)) = :nivel_0')
            else:
                in_keys = []
                for i, val in enumerate(niveles):
                    key = f'nivel_{i}'
                    in_keys.append(f':{key}')
                    params[key] = val
                clauses.append('LOWER(TRIM(nivel_formacion)) IN (' + ','.join(in_keys) + ')')
            if 'nivel_0' not in params and niveles:
                params['nivel_0'] = niveles[0]
    if periodo is not None:
        periodos = [p.strip() for p in str(periodo).split(',') if p.strip()]
        if periodos:
            if len(periodos) == 1:
                clauses.append('periodo = :periodo_0')
            else:
                in_keys = []
                for i, val in enumerate(periodos):
                    key = f'periodo_{i}'
                    in_keys.append(f':{key}')
                    params[key] = int(val)
                clauses.append('periodo IN (' + ','.join(in_keys) + ')')
            if 'periodo_0' not in params and periodos:
                params['periodo_0'] = int(periodos[0])
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

    # Exportar todas las columnas de la tabla fichas_formacion.
    df_export = df.copy()

    # Conservar los nombres originales para poder decidir que columnas ocultar.
    original_cols_fichas = list(df_export.columns)

    # Encabezados legibles para Excel: sin guion bajo y con formato titulo.
    df_export.columns = [export_header_label(col) for col in df_export.columns]

    # Columnas que NO se ven en la tabla del frontend (se ocultaran en Excel).
    hidden_fichas_db_cols = {'cod_municipio', 'cod_regional', 'cod_centro', 'perfil_ingreso'}
    hidden_fichas_headers = {
        export_header_label(c) for c in hidden_fichas_db_cols if c in original_cols_fichas
    }

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

            # Ocultar columnas que no se muestran en la tabla del frontend.
            header_value = ws.cell(row=1, column=col_idx).value
            if header_value in hidden_fichas_headers:
                ws.column_dimensions[col_letter].hidden = True

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
async def upload_programas_excel(
    file: UploadFile = File(...),
    fecha_corte_manual: Optional[date] = Form(None),
):
    """Subida normal de programas.

    Prioriza la fecha de corte manual enviada por el frontend.
    Si no se envia, usa compatibilidad con nombre del archivo o celda A3.
    """
    if not file.filename.lower().endswith(('.xls', '.xlsx', '.xml')):
        raise HTTPException(status_code=400, detail='El archivo debe ser .xls, .xlsx o .xml')

    content = await file.read()
    fecha_corte_file = fecha_corte_manual
    if not fecha_corte_file:
        fecha_corte_file = extract_fecha_corte_from_filename(file.filename or '')
    # Si el nombre no trae fecha, intentar leerla desde el contenido (A3)
    if not fecha_corte_file and file.filename.lower().endswith(('.xls', '.xlsx')):
        fecha_corte_file = extract_fecha_corte_from_excel_content(content)
    if not fecha_corte_file:
        raise HTTPException(
            status_code=400,
            detail='No se pudo obtener fecha_corte. Envia fecha_corte_manual o usa un archivo con fecha en nombre/celda A3.',
        )

    stats = _process_programas_excel(content=content, filename=file.filename or '', fecha_corte_file=fecha_corte_file)
    return JSONResponse(stats)


@app.post('/programas/upload-excel-historico')
async def upload_programas_excel_historico(file: UploadFile = File(...), year: int = Form(...)):
    """Subida de archivos historicos de programas.

    Estos archivos no traen fecha de corte explicita, solo el anio. Se toma
    como fecha_corte el 31/12 de ese anio para que los filtros por anio
    funcionen igual que con los archivos normales.
    """
    if not file.filename.lower().endswith(('.xls', '.xlsx', '.xml')):
        raise HTTPException(status_code=400, detail='El archivo debe ser .xls, .xlsx o .xml')

    try:
        year_int = int(year)
    except Exception:
        raise HTTPException(status_code=400, detail='El anio historico es invalido')

    if year_int < 1900 or year_int > 2100:
        raise HTTPException(status_code=400, detail='El anio historico debe estar entre 1900 y 2100')

    content = await file.read()
    # Usamos como fecha_corte el ultimo dia del anio para que YEAR(fecha_corte)
    # coincida con el anio historico que selecciona el usuario en los filtros.
    fecha_corte_file = date(year_int, 12, 31)

    stats = _process_programas_excel(content=content, filename=file.filename or '', fecha_corte_file=fecha_corte_file)
    return JSONResponse(stats)


def _process_programas_excel(*, content: bytes, filename: str, fecha_corte_file: date) -> dict:
    """Logica comun para importar programas desde un Excel.

    Se usa tanto para la subida normal como para la subida historica.
    """
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
        # Ahora tipo_formacion se toma principalmente desde MODALIDAD_FORMACION
        # (normalizado a modalidad_formacion), manteniendo aliases anteriores
        # como compatibilidad por si vienen otros formatos viejos.
        'tipo_formacion': ['modalidad_formacion', 'tipo_formacion', 'tipo_de_formacion', 'nombre_tipo_formacion'],
    }

    # Detectar si los encabezados reales no estan en la primera fila (caso tipico: primera fila con titulo PE-04_...)
    alias_pool = set()
    for aliases in col_map.values():
        for alias in aliases:
            alias_pool.add(normalize_col_name(alias))

    # Permitir XML: si la extension es .xml, leerlo como tabla antes de aplicar
    # la logica de deteccion de encabezados propia de Excel.
    if filename.lower().endswith('.xml'):
        # Para programas, los XML suelen ser archivos de Excel 2003 (SpreadsheetML).
        # Los leemos como hoja de calculo, no como tabla generica.
        try:
            df = read_spreadsheetml_xml(content)
        except HTTPException:
            raise
        except Exception as e:
            raise HTTPException(
                status_code=400,
                detail=f'No se pudo leer el XML de programas como Excel: {e}',
            )
    else:
        df = read_excel_basic(content)
    if df.empty:
        raise HTTPException(status_code=400, detail='El Excel no contiene filas')

    norm_default_headers = [normalize_col_name(str(c)) for c in list(df.columns)]
    default_score = len(set(norm_default_headers).intersection(alias_pool))

    # Si detecta pocos encabezados utiles o muchos "unnamed", intenta encontrar la fila de encabezado correcta.
    unnamed_count = sum(1 for c in norm_default_headers if c.startswith('unnamed:'))
    if (default_score < 3 or unnamed_count >= 5) and not filename.lower().endswith('.xml'):
        # Solo intentamos la deteccion avanzada de fila de encabezado cuando el
        # archivo es realmente un Excel (xls/xlsx). Para XML ya tenemos el
        # DataFrame correcto desde pd.read_xml y reintentar leerlo como Excel
        # provoca errores de formato.
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

    # fecha_corte se recibe ya calculada (normal o historica).
    df_out['fecha_corte'] = fecha_corte_file

    # Normalizacion de tipos
    for dcol in ['fecha_inicio', 'fecha_fin']:
        # En archivos historicos las fechas suelen venir como dia/mes/anio
        # (por ejemplo 31/12/2018), por eso usamos dayfirst=True.
        df_out[dcol] = pd.to_datetime(df_out[dcol], errors='coerce', dayfirst=True).dt.date

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
        # - Si la ficha ya existe, no se inserta una fila nueva; se actualizan
        #   algunos campos segun la fecha_corte:
        #   - aprendices_activos siempre se actualiza con el valor mas reciente
        #     disponible (si viene en el Excel).
        #   - "aprendices matriculados" (antes cupos) debe representar el valor
        #     del corte mas antiguo disponible para esa ficha. Es decir:
        #       * Si el nuevo archivo tiene fecha_corte mas reciente que la
        #         ya almacenada, NO se toca cupos.
        #       * Si el nuevo archivo tiene fecha_corte mas antigua, se
        #         actualizan cupos y fecha_corte con ese corte mas viejo.
        #   - estrategia_programa y estado_curso se siguen actualizando de
        #     forma "suave" solo cuando vienen valores.

        # Normalizar lista de fichas del Excel (no nulas)
        ficha_series = df_out['numero_ficha'].dropna() if 'numero_ficha' in df_out.columns else pd.Series([], dtype='Int64')
        ficha_ids = [int(x) for x in ficha_series.tolist()]

        existing_ids: set[int] = set()
        existing_fechas: dict[int, Optional[date]] = {}
        if ficha_ids:
            check_sql = text('SELECT numero_ficha, fecha_corte FROM programas_formacion WHERE numero_ficha IN :ids').bindparams(bindparam('ids', expanding=True))
            with engine.connect() as conn:
                rows = conn.execute(check_sql, {'ids': ficha_ids}).fetchall()
            for r in rows:
                if not r or r[0] is None:
                    continue
                num = int(r[0])
                existing_ids.add(num)
                fecha_val = r[1] if len(r) > 1 else None
                if isinstance(fecha_val, datetime):
                    existing_fechas[num] = fecha_val.date()
                elif isinstance(fecha_val, date):
                    existing_fechas[num] = fecha_val
                else:
                    existing_fechas[num] = None

        # Estadisticas de duplicados (respecto a la tabla existente)
        duplicate_fichas = len(existing_ids)
        duplicate_rows_total = int(df_out['numero_ficha'].isin(existing_ids).sum()) if existing_ids and 'numero_ficha' in df_out.columns else 0

        # Filas nuevas (fichas que aun no existen en la tabla)
        df_new = df_out[~df_out['numero_ficha'].isin(existing_ids)].copy() if 'numero_ficha' in df_out.columns else df_out.copy()
        if not df_new.empty:
            df_new.to_sql('programas_formacion', con=engine, if_exists='append', index=False)

        # Filas existentes: actualizar segun reglas descritas arriba
        df_update = df_out[df_out['numero_ficha'].isin(existing_ids)].copy() if existing_ids else pd.DataFrame(columns=df_out.columns)

        updated_fichas = 0
        if not df_update.empty:
            update_params = []
            for _, row in df_update.iterrows():
                if pd.isna(row.get('numero_ficha')):
                    continue
                num_ficha = int(row['numero_ficha'])

                # Fecha de corte actualmente almacenada para esta ficha (si existe)
                old_fecha = existing_fechas.get(num_ficha)
                new_fecha = fecha_corte_file

                # cupos (aprendices matriculados): solo se actualiza cuando el
                # nuevo archivo tiene una fecha_corte mas antigua que la
                # almacenada o cuando no hay fecha previa.
                cupos_val = row.get('cupos') if 'cupos' in df_update.columns else None
                cupos_param = None
                fecha_param = None
                if pd.notna(cupos_val):
                    try:
                        cupos_int = int(cupos_val)
                    except Exception:
                        cupos_int = None
                    if cupos_int is not None:
                        if old_fecha is None or new_fecha is None or new_fecha < old_fecha:
                            cupos_param = cupos_int
                            fecha_param = new_fecha

                # aprendices_activos: siempre se actualiza si viene algun valor
                activos_val = row.get('aprendices_activos') if 'aprendices_activos' in df_update.columns else None
                activos_param = None
                if pd.notna(activos_val):
                    try:
                        activos_param = int(activos_val)
                    except Exception:
                        activos_param = None

                # estrategia_programa y estado_curso: comportamiento "suave"
                est_param = clean_optional_text(row['estrategia_programa']) if 'estrategia_programa' in df_update.columns else None
                estado_param = clean_optional_text(row['estado_curso']) if 'estado_curso' in df_update.columns else None

                update_params.append(
                    {
                        'numero_ficha': num_ficha,
                        'estrategia_programa': est_param,
                        'estado_curso': estado_param,
                        'cupos': cupos_param,
                        'aprendices_activos': activos_param,
                        'fecha_corte': fecha_param,
                    }
                )

            if update_params:
                update_sql = text(
                    'UPDATE programas_formacion '
                    'SET '
                    '    estrategia_programa = COALESCE(:estrategia_programa, estrategia_programa), '
                    '    estado_curso = COALESCE(:estado_curso, estado_curso), '
                    '    cupos = COALESCE(:cupos, cupos), '
                    '    aprendices_activos = COALESCE(:aprendices_activos, aprendices_activos), '
                    '    fecha_corte = COALESCE(:fecha_corte, fecha_corte) '
                    'WHERE numero_ficha = :numero_ficha'
                )
                with engine.begin() as conn:
                    result = conn.execute(update_sql, update_params)
                updated_fichas = int(result.rowcount or 0)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'Error insertando/actualizando programas: {e}')

    return {
        'inserted': int(len(df_new)) if 'df_new' in locals() else 0,
        'updated_fichas': int(updated_fichas),
        'duplicate_fichas': int(duplicate_fichas) if 'duplicate_fichas' in locals() else 0,
        'duplicate_rows': int(duplicate_rows_total) if 'duplicate_rows_total' in locals() else 0,
        'fecha_corte': str(fecha_corte_file),
    }


@app.post('/programas/upload-certificados')
async def upload_programas_certificados(file: UploadFile = File(...)):
    """Actualiza campo `certificado` en programas_formacion usando un Excel complementario y cruce por numero_ficha."""
    if not file.filename.lower().endswith(('.xls', '.xlsx', '.xml')):
        raise HTTPException(status_code=400, detail='El archivo debe ser .xls, .xlsx o .xml')

    content = await file.read()
    is_xml = file.filename.lower().endswith('.xml')

    if is_xml:
        try:
            df = pd.read_xml(io.BytesIO(content))
        except Exception as e:
            raise HTTPException(
                status_code=400,
                detail=f'No se pudo leer el XML complementario de certificados como tabla: {e}',
            )
    else:
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

    if (default_score < 2 or unnamed_count >= 5) and not is_xml:
        # Igual que en programas: solo intentamos la deteccion de encabezado
        # avanzada cuando el archivo es un Excel real. Para XML confiamos en
        # el DataFrame devuelto por pd.read_xml.
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
    year: Optional[str] = None,
    municipio: Optional[str] = None,
    centro: Optional[str] = None,
    nivel: Optional[str] = None,
    estrategia: Optional[str] = None,
    convenio: Optional[str] = None,
    vigencia: Optional[str] = None,
    numero_ficha: Optional[int] = None,
    search: Optional[str] = None,
    solo_certificados: Optional[str] = None,
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
        per_page = 20
    if page < 1:
        page = 1
    if per_page < 1:
        per_page = 20
    if per_page > 20:
        per_page = 20

    clauses = []
    params: dict = {}

    if year is not None:
        years = [y.strip() for y in str(year).split(',') if y.strip()]
        if years:
            if len(years) == 1:
                clauses.append('YEAR(fecha_corte) = :year_0')
            else:
                in_keys = []
                for i, val in enumerate(years):
                    key = f'year_{i}'
                    in_keys.append(f':{key}')
                    params[key] = int(val)
                clauses.append('YEAR(fecha_corte) IN (' + ','.join(in_keys) + ')')
            if 'year_0' not in params and years:
                params['year_0'] = int(years[0])
    if municipio:
        municipios = [m.strip().lower() for m in str(municipio).split(',') if m.strip()]
        if municipios:
            if len(municipios) == 1:
                clauses.append('LOWER(TRIM(ciudad_municipio)) = :municipio_0')
            else:
                in_keys = []
                for i, val in enumerate(municipios):
                    key = f'municipio_{i}'
                    in_keys.append(f':{key}')
                    params[key] = val
                clauses.append('LOWER(TRIM(ciudad_municipio)) IN (' + ','.join(in_keys) + ')')
            if 'municipio_0' not in params and municipios:
                params['municipio_0'] = municipios[0]
    if centro:
        centros = [c.strip().lower() for c in str(centro).split(',') if c.strip()]
        if centros:
            if len(centros) == 1:
                clauses.append('LOWER(TRIM(centro_formacion)) = :centro_0')
            else:
                in_keys = []
                for i, val in enumerate(centros):
                    key = f'centro_{i}'
                    in_keys.append(f':{key}')
                    params[key] = val
                clauses.append('LOWER(TRIM(centro_formacion)) IN (' + ','.join(in_keys) + ')')
            if 'centro_0' not in params and centros:
                params['centro_0'] = centros[0]
    if nivel:
        niveles = [n.strip().lower() for n in str(nivel).split(',') if n.strip()]
        if niveles:
            if len(niveles) == 1:
                clauses.append('LOWER(TRIM(nivel_formacion)) = :nivel_0')
            else:
                in_keys = []
                for i, val in enumerate(niveles):
                    key = f'nivel_{i}'
                    in_keys.append(f':{key}')
                    params[key] = val
                clauses.append('LOWER(TRIM(nivel_formacion)) IN (' + ','.join(in_keys) + ')')
            if 'nivel_0' not in params and niveles:
                params['nivel_0'] = niveles[0]
    if estrategia:
        estrategias = [e.strip().lower() for e in str(estrategia).split(',') if e.strip()]
        if estrategias:
            if len(estrategias) == 1:
                clauses.append('LOWER(TRIM(estrategia_programa)) = :estrategia_0')
            else:
                in_keys = []
                for i, val in enumerate(estrategias):
                    key = f'estrategia_{i}'
                    in_keys.append(f':{key}')
                    params[key] = val
                clauses.append('LOWER(TRIM(estrategia_programa)) IN (' + ','.join(in_keys) + ')')
            if 'estrategia_0' not in params and estrategias:
                params['estrategia_0'] = estrategias[0]
    if convenio:
        convenios = [c.strip().lower() for c in str(convenio).split(',') if c.strip()]
        if convenios:
            if len(convenios) == 1:
                clauses.append('LOWER(TRIM(convenio)) = :convenio_0')
            else:
                in_keys = []
                for i, val in enumerate(convenios):
                    key = f'convenio_{i}'
                    in_keys.append(f':{key}')
                    params[key] = val
                clauses.append('LOWER(TRIM(convenio)) IN (' + ','.join(in_keys) + ')')
            if 'convenio_0' not in params and convenios:
                params['convenio_0'] = convenios[0]
    if vigencia is not None:
        vigencias = [v.strip() for v in str(vigencia).split(',') if v.strip()]
        if vigencias:
            if len(vigencias) == 1:
                clauses.append('YEAR(fecha_inicio) = :vigencia_0')
            else:
                in_keys = []
                for i, val in enumerate(vigencias):
                    key = f'vigencia_{i}'
                    in_keys.append(f':{key}')
                    params[key] = int(val)
                clauses.append('YEAR(fecha_inicio) IN (' + ','.join(in_keys) + ')')
            if 'vigencia_0' not in params and vigencias:
                params['vigencia_0'] = int(vigencias[0])
    if numero_ficha is not None:
        clauses.append('numero_ficha = :numero_ficha')
        params['numero_ficha'] = int(numero_ficha)
    if search:
        s = str(search).strip().lower()
        if s:
            clauses.append('LOWER(TRIM(denominacion_programa)) LIKE :search')
            params['search'] = f'%{s}%'
    # solo_certificados: cualquier valor no vacio/"0"/"false" activa el filtro
    if solo_certificados and str(solo_certificados).strip().lower() not in {'0', 'false', 'no'}:
        clauses.append('(certificado IS NOT NULL AND certificado <> 0)')

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


@app.get('/programas/export')
def export_programas_excel(
    year: Optional[str] = None,
    municipio: Optional[str] = None,
    centro: Optional[str] = None,
    estrategia: Optional[str] = None,
    convenio: Optional[str] = None,
    vigencia: Optional[str] = None,
    numero_ficha: Optional[int] = None,
    search: Optional[str] = None,
    solo_certificados: Optional[str] = None,
):
    """Exporta Excel de programas_formacion respetando los filtros activos."""
    clauses = []
    params: dict = {}

    if year is not None:
        years = [y.strip() for y in str(year).split(',') if y.strip()]
        if years:
            if len(years) == 1:
                clauses.append('YEAR(fecha_corte) = :year_0')
            else:
                in_keys = []
                for i, val in enumerate(years):
                    key = f'year_{i}'
                    in_keys.append(f':{key}')
                    params[key] = int(val)
                clauses.append('YEAR(fecha_corte) IN (' + ','.join(in_keys) + ')')
            if 'year_0' not in params and years:
                params['year_0'] = int(years[0])
    if municipio:
        municipios = [m.strip().lower() for m in str(municipio).split(',') if m.strip()]
        if municipios:
            if len(municipios) == 1:
                clauses.append('LOWER(TRIM(ciudad_municipio)) = :municipio_0')
            else:
                in_keys = []
                for i, val in enumerate(municipios):
                    key = f'municipio_{i}'
                    in_keys.append(f':{key}')
                    params[key] = val
                clauses.append('LOWER(TRIM(ciudad_municipio)) IN (' + ','.join(in_keys) + ')')
            if 'municipio_0' not in params and municipios:
                params['municipio_0'] = municipios[0]
    if centro:
        centros = [c.strip().lower() for c in str(centro).split(',') if c.strip()]
        if centros:
            if len(centros) == 1:
                clauses.append('LOWER(TRIM(centro_formacion)) = :centro_0')
            else:
                in_keys = []
                for i, val in enumerate(centros):
                    key = f'centro_{i}'
                    in_keys.append(f':{key}')
                    params[key] = val
                clauses.append('LOWER(TRIM(centro_formacion)) IN (' + ','.join(in_keys) + ')')
            if 'centro_0' not in params and centros:
                params['centro_0'] = centros[0]
    if estrategia:
        estrategias = [e.strip().lower() for e in str(estrategia).split(',') if e.strip()]
        if estrategias:
            if len(estrategias) == 1:
                clauses.append('LOWER(TRIM(estrategia_programa)) = :estrategia_0')
            else:
                in_keys = []
                for i, val in enumerate(estrategias):
                    key = f'estrategia_{i}'
                    in_keys.append(f':{key}')
                    params[key] = val
                clauses.append('LOWER(TRIM(estrategia_programa)) IN (' + ','.join(in_keys) + ')')
            if 'estrategia_0' not in params and estrategias:
                params['estrategia_0'] = estrategias[0]
    if convenio:
        convenios = [c.strip().lower() for c in str(convenio).split(',') if c.strip()]
        if convenios:
            if len(convenios) == 1:
                clauses.append('LOWER(TRIM(convenio)) = :convenio_0')
            else:
                in_keys = []
                for i, val in enumerate(convenios):
                    key = f'convenio_{i}'
                    in_keys.append(f':{key}')
                    params[key] = val
                clauses.append('LOWER(TRIM(convenio)) IN (' + ','.join(in_keys) + ')')
            if 'convenio_0' not in params and convenios:
                params['convenio_0'] = convenios[0]
    if vigencia is not None:
        vigencias = [v.strip() for v in str(vigencia).split(',') if v.strip()]
        if vigencias:
            if len(vigencias) == 1:
                clauses.append('YEAR(fecha_inicio) = :vigencia_0')
            else:
                in_keys = []
                for i, val in enumerate(vigencias):
                    key = f'vigencia_{i}'
                    in_keys.append(f':{key}')
                    params[key] = int(val)
                clauses.append('YEAR(fecha_inicio) IN (' + ','.join(in_keys) + ')')
            if 'vigencia_0' not in params and vigencias:
                params['vigencia_0'] = int(vigencias[0])
    if numero_ficha is not None:
        clauses.append('numero_ficha = :numero_ficha')
        params['numero_ficha'] = int(numero_ficha)
    if search:
        s = str(search).strip().lower()
        if s:
            clauses.append('LOWER(TRIM(denominacion_programa)) LIKE :search')
            params['search'] = f'%{s}%'
    if solo_certificados and str(solo_certificados).strip().lower() not in {'0', 'false', 'no'}:
        clauses.append('(certificado IS NOT NULL AND certificado <> 0)')

    where_sql = ''
    if clauses:
        where_sql = ' WHERE ' + ' AND '.join(clauses)

    sql = (
        'SELECT * FROM programas_formacion'
        f'{where_sql} '
        'ORDER BY fecha_corte DESC, numero_ficha ASC, id ASC'
    )

    try:
        df = pd.read_sql(text(sql), con=engine, params=params)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'Error exportando programas: {e}')

    # Exportar todas las columnas de la tabla programas_formacion.
    df_export = df.copy()

    # En el frontend se oculta la columna "id"; aqui la mantenemos en el Excel
    # pero la marcamos como oculta para que no aparezca a simple vista.
    original_cols_programas = list(df_export.columns)
    hidden_programas_headers = set()
    if 'id' in original_cols_programas:
        hidden_programas_headers.add('id')

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name='programas')

        ws = writer.book['programas']
        max_row = ws.max_row
        max_col = ws.max_column

        wrap_alignment = Alignment(wrap_text=True, vertical='top')
        for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
            for cell in row:
                cell.alignment = wrap_alignment

        for cell in ws[1]:
            cell.font = Font(bold=True)

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

            # Ocultar columnas que no se muestran en la tabla del frontend.
            header_value = ws.cell(row=1, column=col_idx).value
            if header_value in hidden_programas_headers:
                ws.column_dimensions[col_letter].hidden = True

        if max_col >= 1 and max_row >= 1:
            last_col_letter = get_column_letter(max_col)
            table_ref = f'A1:{last_col_letter}{max_row}'
            table = Table(displayName='ProgramasExport', ref=table_ref)
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
    filename = f'programas_export_{ts}.xlsx'

    return StreamingResponse(
        output,
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={'Content-Disposition': f'attachment; filename="{filename}"'},
    )


@app.get('/programas/filters')
def get_programas_filters():
    """Devuelve valores distintos para los filtros de programas a nivel global.

    No aplica paginacion ni filtros previos: siempre consulta toda la tabla
    programas_formacion para construir los combos de la UI.
    """
    try:
        with engine.connect() as conn:
            years = [
                int(r[0])
                for r in conn.execute(
                    text('SELECT DISTINCT YEAR(fecha_corte) AS y FROM programas_formacion WHERE fecha_corte IS NOT NULL ORDER BY y DESC')
                ).fetchall()
                if r[0] is not None
            ]

            vigencias = [
                int(r[0])
                for r in conn.execute(
                    text('SELECT DISTINCT YEAR(fecha_inicio) AS y FROM programas_formacion WHERE fecha_inicio IS NOT NULL ORDER BY y DESC')
                ).fetchall()
                if r[0] is not None
            ]

            municipios = [
                str(r[0])
                for r in conn.execute(
                    text('SELECT DISTINCT ciudad_municipio FROM programas_formacion WHERE ciudad_municipio IS NOT NULL ORDER BY ciudad_municipio ASC')
                ).fetchall()
                if r[0] is not None
            ]

            centros = [
                str(r[0])
                for r in conn.execute(
                    text('SELECT DISTINCT centro_formacion FROM programas_formacion WHERE centro_formacion IS NOT NULL ORDER BY centro_formacion ASC')
                ).fetchall()
                if r[0] is not None
            ]

            niveles = [
                str(r[0])
                for r in conn.execute(
                    text('SELECT DISTINCT nivel_formacion FROM programas_formacion WHERE nivel_formacion IS NOT NULL ORDER BY nivel_formacion ASC')
                ).fetchall()
                if r[0] is not None
            ]

            estrategias = [
                str(r[0])
                for r in conn.execute(
                    text('SELECT DISTINCT estrategia_programa FROM programas_formacion WHERE estrategia_programa IS NOT NULL ORDER BY estrategia_programa ASC')
                ).fetchall()
                if r[0] is not None
            ]

            convenios = [
                str(r[0])
                for r in conn.execute(
                    text('SELECT DISTINCT convenio FROM programas_formacion WHERE convenio IS NOT NULL ORDER BY convenio ASC')
                ).fetchall()
                if r[0] is not None
            ]
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'Error obteniendo filtros de programas: {e}')

    return JSONResponse(
        content=jsonable_encoder(
            {
                'years': years,
                'vigencias': vigencias,
                'municipios': municipios,
                'centros': centros,
                'niveles': niveles,
                'estrategias': estrategias,
                'convenios': convenios,
            }
        )
    )


@app.delete('/programas/delete-all')
def delete_programas_all():
    """Elimina todos los registros de programas_formacion."""
    try:
        with engine.begin() as conn:
            result = conn.execute(text('DELETE FROM programas_formacion'))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'Error eliminando todos los programas: {e}')

    return JSONResponse({'deleted_rows': int(result.rowcount or 0)})


@app.delete('/programas/delete-by-vigencia')
def delete_programas_by_vigencia(vigencia: int):
    """Elimina registros por vigencia (anio de fecha_inicio)."""
    try:
        vig = int(vigencia)
    except Exception:
        raise HTTPException(status_code=400, detail='La vigencia es invalida')

    if vig < 1900 or vig > 2100:
        raise HTTPException(status_code=400, detail='La vigencia debe estar entre 1900 y 2100')

    try:
        with engine.begin() as conn:
            result = conn.execute(
                text('DELETE FROM programas_formacion WHERE YEAR(fecha_inicio) = :vigencia'),
                {'vigencia': vig},
            )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'Error eliminando programas por vigencia: {e}')

    return JSONResponse({'vigencia': vig, 'deleted_rows': int(result.rowcount or 0)})


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
