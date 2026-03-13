import pandas as pd

EXPECTED_COLUMNS = [
    'cod_regional', 'regional', 'cod_municipio', 'municipio', 'cod_centro', 'centro_formacion',
    'cod_programa', 'denominacion_programa', 'cod_ficha', 'estado_ficha', 'jornada', 'nivel_formacion',
    'cupo', 'inscritos_primera_opcion', 'inscritos_segunda_opcion', 'oferta', 'tipo', 'perfil_ingreso', 'periodo'
]

rows = [
    [66, 'REGIONAL RISARALDA', 57066001, 'PEREIRA', 9308, 'CENTRO DE COMERCIO Y SERVICIOS', 121202, 'GESTIÓN BANCARIA Y DE ENTIDADES FINANCIERAS', 3140146, 'En Ejecución', 'DIURNA', 'TECNÓLOGO', 31, 131, 25, 'I', 'PRESENCIAL Y A DISTANCIA', None, 2025],
    [66, 'REGIONAL RISARALDA', 57066001, 'PEREIRA', 9308, 'CENTRO DE COMERCIO Y SERVICIOS', 635503, 'COCINA.', 3140121, 'En Ejecución', 'DIURNA', 'TÉCNICO', 33, 141, 10, 'I', 'PRESENCIAL Y A DISTANCIA', None, 2025],
    [66, 'REGIONAL RISARALDA', 57066001, 'PEREIRA', 9308, 'CENTRO DE COMERCIO Y SERVICIOS', 621600, 'COORDINACION DE SERVICIOS HOTELEROS', 3140156, 'En Ejecución', 'DIURNA', 'TECNÓLOGO', 30, 48, 10, 'I', 'PRESENCIAL Y A DISTANCIA', None, 2025],
]

df = pd.DataFrame(rows, columns=EXPECTED_COLUMNS)
# Guardar como Excel
output = 'test_fichas.xlsx'
df.to_excel(output, index=False)
print(f'Archivo creado: {output} (columnas: {len(df.columns)}, filas: {len(df)})')
