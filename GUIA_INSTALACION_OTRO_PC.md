# Guia de instalacion (carpeta ya disponible)

Esta guia asume que la otra persona ya tiene la carpeta del proyecto en su PC con Windows.

## 1. Requisitos previos

- Python recomendado: `3.11` (tambien suele funcionar `3.10` o `3.12`).
- MySQL Server + MySQL Workbench instalados.
- Carpeta del proyecto ya disponible localmente.

Verificar Python en terminal:

```powershell
python --version
```

## 2. Crear base de datos en MySQL Workbench

1. Abre MySQL Workbench y conecta al servidor local.
2. Abre el archivo `base_datos.sql`.
3. Ejecuta todo el script (icono de rayo):
   - Crea la BD `sena_oferta`.
   - Crea la tabla `fichas_formacion`.
   - Inserta datos iniciales.

Si prefieres por consola MySQL:

```sql
SOURCE C:/ruta/base de datos/base_datos.sql;
```

## 3. Configurar URL de base de datos (DATABASE_URL)

En `fastapi_app/.env` debe existir la variable `DATABASE_URL`.

Ejemplo:

```env
DATABASE_URL=mysql+pymysql://root:TU_PASSWORD@localhost/sena_oferta
```

Formato general:

```text
mysql+pymysql://USUARIO:PASSWORD@HOST/NOMBRE_BD
```

Notas:

- `USUARIO` usualmente: `root`
- `HOST` usualmente: `localhost` o `127.0.0.1`
- `NOMBRE_BD` para este proyecto: `sena_oferta`

## 4. Crear entorno virtual (mismo nombre)

En este proyecto el entorno se usa como `.venv` dentro de `fastapi_app`, asi que en el otro PC crea el mismo nombre.

Desde la raiz del proyecto:

```powershell
cd "C:\ruta\base de datos\fastapi_app"
python -m venv .venv
```

## 5. Activar entorno virtual

### PowerShell

```powershell
cd "C:\ruta\base de datos\fastapi_app"
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
```

### CMD

```bat
cd /d "C:\ruta\base de datos\fastapi_app"
.venv\Scripts\activate.bat
```

Para desactivar (ambos):

```bash
deactivate
```

## 6. Instalar dependencias

Con el entorno activado:

```powershell
python -m pip install --upgrade pip
pip install -r requirements.txt
```

Dependencias principales del backend:

- `fastapi`
- `uvicorn[standard]`
- `pandas`
- `openpyxl`
- `sqlalchemy`
- `pymysql`
- `python-dotenv`
- `python-multipart`

## 7. Ejecutar API (FastAPI + Uvicorn)

Desde `fastapi_app` con entorno activado:

```powershell
python run_uvicorn.py
```

O directo:

```powershell
uvicorn fastapi_app.main:app --reload --host 0.0.0.0 --port 8000
```

Comprobar que levanto:

- API: `http://127.0.0.1:8000`
- Docs: `http://127.0.0.1:8000/docs`

## 8. Ejecutar frontend

El frontend esta en `frontend/` y consume la API en `http://127.0.0.1:8000`.

Opcion recomendada (servidor simple):

```powershell
cd "C:\ruta\base de datos\frontend"
python -m http.server 5500
```

Abrir en navegador:

- `http://127.0.0.1:5500/index.html`

## 9. Checklist rapido de errores comunes

- Error de conexion MySQL:
  - Revisar `DATABASE_URL` en `fastapi_app/.env`.
  - Verificar usuario/password/host y que exista `sena_oferta`.
- Error al activar entorno en PowerShell:
  - Ejecutar `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass`.
- `ModuleNotFoundError`:
  - Confirmar entorno activado y correr `pip install -r requirements.txt`.
- Puerto 8000 ocupado:
  - Cambiar puerto en uvicorn, por ejemplo `--port 8001`.

## 10. Orden recomendado para levantar

1. Crear BD con `base_datos.sql`.
2. Crear `fastapi_app/.env` con `DATABASE_URL` correcta.
3. Crear entorno `.venv` y activarlo.
4. Instalar `requirements.txt`.
5. Levantar API con `python run_uvicorn.py`.
6. Levantar frontend con `python -m http.server 5500`.
7. Probar filtros, carga de Excel y exportacion.
