# Este archivo contiene los endpoints de exportación a Excel
# Agregar este código al final de fastapi_app/main.py

# ===== ENDPOINTS DE EXPORTACIÓN A EXCEL =====

def _export_table_to_excel(table_name: str, output_filename: str) -> StreamingResponse:
    """
    Función genérica para exportar cualquier tabla a Excel
    """
    try:
        sql = f'SELECT * FROM {table_name}'
        
        with engine.connect() as conn:
            df = pd.read_sql(text(sql), con=engine)
        
        if df.empty:
            df = pd.DataFrame()
        
        # Convertir columnas datetime a string para mejor manejo en Excel
        for col in df.columns:
            if df[col].dtype == 'object':
                try:
                    df[col] = pd.to_datetime(df[col], errors='coerce')
                    df[col] = df[col].dt.strftime('%Y-%m-%d %H:%M:%S')
                except:
                    pass
        
        # Crear archivo Excel en memoria
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Datos', index=False)
            
            # Dar formato a la hoja
            workbook = writer.book
            worksheet = writer.sheets['Datos']
            
            # Ajustar ancho de columnas
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
            
            # Formatear encabezados
            for cell in worksheet[1]:
                if cell.value:
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        output.seek(0)
        
        return StreamingResponse(
            iter([output.getvalue()]),
            media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            headers={'Content-Disposition': f'attachment; filename="{output_filename}"'}
        )
    
    except Exception as e:
        import traceback
        raise HTTPException(status_code=500, detail=f'Error exportando {table_name}: {str(e)}\n{traceback.format_exc()}')


@app.get('/catalogo/exportar-excel')
def export_catalogo():
    """Exportar tabla catalogo a Excel"""
    return _export_table_to_excel('catalogo', 'catalogo.xlsx')


@app.get('/registro-calificado/exportar-excel')
def export_registro_calificado():
    """Exportar tabla registro_calificado_presencial a Excel"""
    return _export_table_to_excel('registro_calificado_presencial', 'registro_calificado_presencial.xlsx')


@app.get('/oferta/exportar-excel')
def export_oferta():
    """Exportar tabla OFERTA_seguimiento_metas a Excel"""
    return _export_table_to_excel('OFERTA_seguimiento_metas', 'oferta_seguimiento_metas.xlsx')


@app.get('/consolidado-colegios/exportar-excel')
def export_consolidado_colegios():
    """Exportar tabla consolidado_colegios a Excel"""
    return _export_table_to_excel('consolidado_colegios', 'consolidado_colegios.xlsx')
