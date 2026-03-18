# Guía de instalación en otro PC (Windows)

Esta guía asume que ya copiaste toda la carpeta del proyecto a otro PC con Windows.

Ejemplo de ruta (ajústala a tu caso):

```text
C:\Users\TU_USUARIO\Documentos\PAGINA_OFERTA\
```

---

## 1. Requisitos previos

- Python 3.11 recomendado (también suele funcionar 3.10 o 3.12).
- MySQL Server y MySQL Workbench instalados.
- Carpeta del proyecto ya copiada en el nuevo PC.

Comprobar Python en PowerShell:

```powershell
python --version
pip --version
```

Si no responde o no es 3.x, instala Python desde la página oficial y marca la opción "Add Python to PATH".

---

## 2. Crear base de datos y tablas con MySQL Workbench

1. Abre **MySQL Workbench** y conéctate a tu servidor local (normalmente `localhost` o `127.0.0.1`).
2. En el menú, ve a **File → Open SQL Script…** y abre el archivo `base_datos.sql` que está en la carpeta del proyecto.
3. Verifica que en la pestaña del script se vea algo como:
   - `CREATE DATABASE sena_oferta;`
   - `USE sena_oferta;`
   - `CREATE TABLE fichas_formacion (...);`
   - `CREATE TABLE IF NOT EXISTS programas_formacion (...);`
   - `CREATE TABLE IF NOT EXISTS indicativa (...);`
4. Haz clic en el icono de **rayo (Execute)** para ejecutar TODO el script.

Esto crea:

- La base de datos: `sena_oferta`.
- Las tablas: `fichas_formacion`, `programas_formacion` e `indicativa`.

> Alternativa por consola MySQL:

```sql
SOURCE C:/RUTA/DEL/PROYECTO/base_datos.sql;
```

---

## 3. Configurar la URL de la base de datos (DATABASE_URL)

El backend usa la variable `DATABASE_URL` en el archivo `fastapi_app/.env`.

Abre `fastapi_app/.env` y ajusta la línea a tus datos reales:

```env
DATABASE_URL=mysql+pymysql://USUARIO:PASSWORD@HOST:PUERTO/sena_oferta
```

Ejemplo típico en local:

```env
DATABASE_URL=mysql+pymysql://root:MiPassword@localhost:3306/sena_oferta
```

Donde:

- `USUARIO`: normalmente `root` (o el usuario que uses en MySQL Workbench).
- `PASSWORD`: la contraseña de MySQL.
- `HOST`: `localhost` o `127.0.0.1`.
- `PUERTO`: normalmente `3306`.
- `sena_oferta`: debe coincidir con el nombre creado en `base_datos.sql`.

> Si cambias el nombre de la base en MySQL, también cámbialo aquí.

---

## 4. Crear el entorno virtual (carpeta `oferta`)

En este proyecto se usa un entorno virtual dentro de `fastapi_app` llamado **`oferta`**.

1. Abre **PowerShell**.
2. Ve a la carpeta del backend:

```powershell
cd "C:\RUTA\DEL\PROYECTO\fastapi_app"
```

3. Crea el entorno virtual `oferta` (solo la primera vez en cada PC):

```powershell
python -m venv oferta
```

Esto creará la carpeta `fastapi_app\oferta\` con `Scripts`, `Lib`, etc.

---

## 5. Activar el entorno virtual y ExecutionPolicy

### 5.1. Activar en PowerShell

Desde la carpeta `fastapi_app`:

```powershell
cd "C:\RUTA\DEL\PROYECTO\fastapi_app"
```

Intenta activar el entorno:

```powershell
.\n+oferta\Scripts\Activate.ps1
```

Si **funciona**, verás algo como `(oferta)` al inicio de la línea; ya puedes pasar al paso 6.

### 5.2. Si NO deja activar (scripts deshabilitados)

Si aparece un error tipo:

> "running scripts is disabled on this system"

entonces, en esa misma ventana de PowerShell ejecuta:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

Esto solo cambia la política **para esa ventana** (no toca todo el sistema).

Después vuelve a intentar activar el entorno:

```powershell
.
oferta\Scripts\Activate.ps1
```

Cuando veas `(oferta)` al inicio de la línea, el entorno está activo.

### 5.3. Activar en CMD (opcional)

Si prefieres usar CMD clásico:

```bat
cd /d C:\RUTA\DEL\PROYECTO\fastapi_app
oferta\Scripts\activate.bat
```

Para desactivar el entorno (PowerShell o CMD):

```bash

```

---

## 6. Instalar los requisitos del backend

Con el entorno **oferta** activado y estando en `fastapi_app`:

```powershell
python -m pip install --upgrade pip
pip install -r requirements.txt
```

Si falla por alguna librería, repite el comando o verifica la conexión a internet.

Dependencias principales:

- `fastapi`
- `uvicorn[standard]`
- `pandas`
- `openpyxl`
- `sqlalchemy`
- `pymysql`
- `python-dotenv`
- `python-multipart`

> El script `fastapi_app/start.ps1` también puede instalar requisitos automáticamente, pero es mejor dejarlos instalados con este paso.

---

## 7. Iniciar el servidor FastAPI (Uvicorn)

Con el entorno **oferta** activado y estando en `fastapi_app`, tienes dos opciones:

### Opción A: usando `run_uvicorn.py`

```powershell
python run_uvicorn.py
```

### Opción B: comando uvicorn directo

Desde la **raíz del proyecto** o desde `fastapi_app`:

```powershell
python -m uvicorn fastapi_app.main:app --reload --host 0.0.0.0 --port 8000
```

Si todo está bien, deberías ver en consola algo como:

```text
Uvicorn running on http://127.0.0.1:8000 (Press CTRL+C to quit)
```

### Probar que la API está arriba y conectada

1. Abre el navegador y entra a:
   - `http://127.0.0.1:8000`
   - `http://127.0.0.1:8000/docs`
2. Si `/docs` carga, FastAPI y Uvicorn están funcionando.
3. Si en la consola ves errores de conexión a MySQL, revisa:
   - Que MySQL Server esté iniciado.
   - Que la BD `sena_oferta` exista.
   - Que la variable `DATABASE_URL` en `fastapi_app/.env` tenga bien usuario, contraseña, host y puerto.

---

## 8. Levantar el frontend (páginas HTML)

El frontend está en la carpeta `frontend/` y consume la API en `http://127.0.0.1:8000`.

### Opción recomendada: servidor HTTP simple

En otra ventana de **PowerShell** (puede ser sin entorno):

```powershell
cd "C:\RUTA\DEL\PROYECTO\frontend"
python -m http.server 5500
```

Luego abre en el navegador:

- Módulo Inscripciones: `http://127.0.0.1:5500/index.html`
- Módulo Programas: `http://127.0.0.1:5500/programas.html`
- Módulo Indicativa: `http://127.0.0.1:5500/indicativa.html`

Asegúrate de que la API (`uvicorn`) siga corriendo en otra ventana.

---

## 9. Errores típicos y cómo resolverlos

**1. Error de conexión a MySQL (al iniciar uvicorn o usar la app)**

- Verifica `fastapi_app/.env` → `DATABASE_URL`.
- Asegúrate de que `sena_oferta` existe (refresca esquemas en MySQL Workbench).
- Revisa usuario, contraseña, host y puerto.

**2. No se puede activar el entorno en PowerShell (scripts deshabilitados)**

- Mensaje típico: _"running scripts is disabled on this system"_.
- Solución:

  ```powershell
  Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
  .\oferta\Scripts\Activate.ps1
  ```

  Esto solo afecta a esa ventana y se pierde al cerrarla.

**3. `ModuleNotFoundError` al levantar uvicorn**

- Asegúrate de que el entorno `oferta` está activo (ver `(oferta)` en la línea).
- Instala paquetes:

  ```powershell
  pip install -r requirements.txt
  ```

**4. Puerto 8000 ocupado**

- Cambia el puerto en el comando uvicorn, por ejemplo:

  ```powershell
  python -m uvicorn fastapi_app.main:app --reload --host 0.0.0.0 --port 8001
  ```

**5. El frontend abre pero no muestra datos**

- Revisa que la API siga corriendo en `http://127.0.0.1:8000`.
- Verifica la consola del navegador (F12 → Console) por errores de red (CORS, 500, etc.).

---

## 10. Resumen rápido (orden recomendado)

1. Instalar Python, MySQL Server y MySQL Workbench.
2. Ejecutar `base_datos.sql` en MySQL Workbench (crea `sena_oferta` y las tablas).
3. Ajustar `fastapi_app/.env` con la `DATABASE_URL` correcta.
4. Crear el entorno virtual `oferta` dentro de `fastapi_app`.
5. Activar el entorno (`oferta\Scripts\Activate.ps1`; si falla, usar `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass`).
6. Instalar dependencias con `pip install -r requirements.txt` dentro de `fastapi_app`.
7. Iniciar uvicorn (`python run_uvicorn.py` o `python -m uvicorn fastapi_app.main:app --reload`).
8. Levantar el servidor estático del frontend (`python -m http.server 5500` dentro de `frontend`).
9. Probar en el navegador: Inscripciones, Programas e Indicativa.

Con estos pasos, el proyecto debería funcionar en cualquier otro PC Windows con la misma estructura de carpetas.
