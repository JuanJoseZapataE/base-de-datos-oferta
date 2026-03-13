from sqlalchemy import create_engine, text
from dotenv import load_dotenv
import os, re

load_dotenv()
DATABASE_URL = os.getenv('DATABASE_URL')
print('Usando DATABASE_URL=', DATABASE_URL)
engine = create_engine(DATABASE_URL)

# Test simple
try:
    with engine.connect() as conn:
        r = conn.execute(text('SELECT 1'))
        print('DB ok, SELECT 1 ->', r.scalar())
except Exception as e:
    print('Error al conectar:', e)
    raise SystemExit(1)

# Comprobar tabla
with engine.connect() as conn:
    r = conn.execute(text("SELECT TABLE_NAME FROM information_schema.tables WHERE table_schema=DATABASE() AND table_name='fichas_formacion'"))
    found = r.fetchone() is not None
    if found:
        print('Tabla fichas_formacion encontrada.')
    else:
        print('Tabla fichas_formacion NO encontrada. Intentando crearla desde base_datos.sql...')
        sql_path = os.path.join(os.path.dirname(__file__), '..', 'base_datos.sql')
        sql_path = os.path.abspath(sql_path)
        if not os.path.exists(sql_path):
            print('No se encontró base_datos.sql en:', sql_path)
            raise SystemExit(1)
        sql_text = open(sql_path, 'r', encoding='utf-8').read()
        # Extraer CREATE TABLE fichas_formacion (...) block
        m = re.search(r"CREATE TABLE\s+fichas_formacion\s*\((.*?)\);", sql_text, re.S | re.I)
        if not m:
            print('No se pudo extraer CREATE TABLE fichas_formacion del SQL.')
            raise SystemExit(1)
        create_body = m.group(1)
        create_sql = 'CREATE TABLE IF NOT EXISTS fichas_formacion (' + create_body + ');'
        try:
            conn.execute(text(create_sql))
            print('CREATE TABLE ejecutado.')
        except Exception as e:
            print('Error al crear la tabla:', e)
            raise SystemExit(1)
print('Script finished successfully.')
