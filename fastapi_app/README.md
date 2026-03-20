📌 Proyecto FastAPI: Carga de Excel a MySQL
==========================================

Aplicación FastAPI para subir un archivo Excel y volcar su contenido en la tabla `fichas_formacion` de la base de datos `sena_oferta`.

Permite además:
- Consultar fichas con filtros.
- Actualizar en bloque los campos `periodo`, `oferta` y `tipo`.

---

🚀 Guía rápida (en 4 pasos)
---------------------------

1) Configurar la conexión a la base de datos
     - Definir la variable de entorno `DATABASE_URL` **o** editarla en el archivo `main.py`.
     - Ejemplo de `DATABASE_URL`:

         ```env
         DATABASE_URL=mysql+pymysql://root:password@localhost/sena_oferta
         ```

2) Instalar dependencias

     Desde la carpeta raíz del proyecto:

     ```bash
     pip install -r requirements.txt
     ```

3) Levantar el servidor

     ```bash
     uvicorn fastapi_app.main:app --reload --host 0.0.0.0 --port 8000
     ```

4) Subir un archivo Excel

     Ejemplo usando `curl`:

     ```bash
     curl -F "file=@/ruta/al/archivo.xlsx" http://localhost:8000/upload-excel
     ```

---

📂 Requisitos del archivo Excel
-------------------------------

- Debe contener las columnas que correspondan a la tabla `fichas_formacion`.
- El endpoint intenta mapear columnas:
    - Primero por **nombre normalizado** (minúsculas, espacios → guiones bajos).
    - Si los nombres no coinciden, intenta mapear por **posición**, siempre que el número de columnas coincida con el esperado.
- No hay autenticación ni validaciones avanzadas en esta versión.

---

📡 Endpoints principales
------------------------

### 1. Subir Excel

- `POST /upload-excel`
- Form-data: campo `file` con el archivo `.xlsx`.

Ejemplo:

```bash
curl -F "file=@/ruta/al/archivo.xlsx" http://localhost:8000/upload-excel
```

### 2. Listar fichas

- `GET /fichas`
- Filtros opcionales por query params:
    - `periodo` (ejemplo: `?periodo=2025`)
    - `oferta` (valores: `1`, `2`, `3`, `4`)
    - `tipo` (ejemplo: `presencial`, `a distancia`, `virtual`)

Ejemplo de consulta filtrada:

```bash
curl "http://localhost:8000/fichas?periodo=2025&oferta=1&tipo=presencial"
```

### 3. Actualizar fichas en bloque

- `POST /fichas/update`
- Permite actualizar `periodo`, `oferta` y/o `tipo` para un conjunto de códigos de ficha.

Cuerpo JSON de ejemplo:

```json
{
    "cod_fichas": [3140146, 3140121],
    "periodo": 2025,
    "oferta": "1",
    "tipo": "presencial"
}
```

Respuesta:

```json
{"updated_rows": 2}
```

📌 Detalles importantes
-----------------------

- `tipo` se normaliza internamente a mayúsculas:
    - `PRESENCIAL`
    - `A DISTANCIA`
    - `VIRTUAL`
    - `PRESENCIAL Y A DISTANCIA`
- `oferta` acepta valores `1`–`4` y se almacena como carácter.

---

🧪 Generar un Excel de prueba
-----------------------------

Puedes generar un archivo de prueba con registros de ejemplo ejecutando:

```bash
python fastapi_app/create_test_excel.py
```

Esto crea el archivo `test_fichas.xlsx` en la carpeta `fastapi_app`.

Para subirlo al servidor en ejecución:

```bash
curl -F "file=@fastapi_app/test_fichas.xlsx" http://localhost:8000/upload-excel
```

Luego puedes comprobar los registros insertados, por ejemplo:

```bash
curl "http://localhost:8000/fichas?periodo=2025"
```

---

ℹ️ Notas finales
----------------

- Actualmente **no** se incluye interfaz web estática; la aplicación expone únicamente la **API REST**.
- Se recomienda usar herramientas como `curl`, Postman o Thunder Client (VS Code) para probar los endpoints.
