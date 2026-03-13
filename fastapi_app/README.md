Proyecto FastAPI para subir un archivo Excel y volcarlo a la tabla `fichas_formacion` en la base de datos `sena_oferta`.

Instrucciones rápidas:

1) Editar la URL de conexión en la variable de entorno `DATABASE_URL` o en el archivo `main.py`.
   - Ejemplo de `DATABASE_URL`:

       mysql+pymysql://root:password@localhost/sena_oferta

2) Instalar dependencias:

```bash
pip install -r requirements.txt
```

3) Ejecutar el servidor:

```bash
uvicorn fastapi_app.main:app --reload --host 0.0.0.0 --port 8000
```

4) Subir Excel (ejemplo con `curl`):

```bash
curl -F "file=@/ruta/al/archivo.xlsx" http://localhost:8000/upload-excel
```

Notas:
- El Excel debe contener las columnas que correspondan a la tabla `fichas_formacion`.
- El endpoint intentará mapear columnas por nombre normalizado (minúsculas, espacios → guiones bajos). Si las columnas no coinciden, intentará mapear por posición si el número de columnas coincide con el esperado.
- No hay autenticación ni validación avanzada en esta versión.

Nuevos endpoints:

- GET `/fichas`: listar fichas con filtros opcionales por query params:
    - `periodo` (ej: `?periodo=2025`)
    - `oferta` (1,2,3,4)
    - `tipo` (ej: `presencial`, `a distancia`, `virtual`)
    Ejemplo:

```bash
curl "http://localhost:8000/fichas?periodo=2025&oferta=1&tipo=presencial"
```

- POST `/fichas/update`: actualizar campos `periodo`, `oferta` y/o `tipo` para un conjunto de `cod_fichas`.
    - Body JSON ejemplo:

```json
{
    "cod_fichas": [3140146, 3140121],
    "periodo": 2025,
    "oferta": "1",
    "tipo": "presencial"
}
```

Respuesta: `{"updated_rows": <número>}`

Observaciones:
- `tipo` se normaliza internamente a mayúsculas (`PRESENCIAL`, `A DISTANCIA`, `VIRTUAL`, `PRESENCIAL Y A DISTANCIA`).
- `oferta` acepta valores `1`-`4` y se guarda como carácter.

Generar un Excel de prueba
--------------------------

Hay un script que genera un Excel de prueba con los 3 registros de ejemplo:

```bash
python fastapi_app/create_test_excel.py
```

El script crea `test_fichas.xlsx` en la carpeta `fastapi_app`. Para subirlo al servidor en ejecución:

```bash
curl -F "file=@fastapi_app/test_fichas.xlsx" http://localhost:8000/upload-excel
```

Después puedes comprobar los registros con:

```bash
curl "http://localhost:8000/fichas?periodo=2025"
```

Nota: ya no se incluye interfaz web estática; la aplicación expone únicamente la API.
