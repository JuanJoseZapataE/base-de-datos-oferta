import os
import sys
from dotenv import load_dotenv
from sqlalchemy import create_engine, text


def main():
    env_path = os.path.join(os.path.dirname(__file__), '.env')
    load_dotenv(env_path)
    DATABASE_URL = os.getenv('DATABASE_URL')
    print('Usando DATABASE_URL:', DATABASE_URL)
    if not DATABASE_URL:
        print('ERROR: DATABASE_URL no definida en', env_path)
        sys.exit(2)

    try:
        engine = create_engine(DATABASE_URL)
        with engine.connect() as conn:
            total = conn.execute(text('SELECT COUNT(*) FROM fichas_formacion')).scalar()
            print('COUNT fichas_formacion =', total)
            sample = conn.execute(text('SELECT * FROM fichas_formacion LIMIT 5')).fetchall()
            print('Muestra (hasta 5 filas):')
            for row in sample:
                print(dict(row))
    except Exception as e:
        print('ERROR al conectar/consultar la base de datos:')
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
