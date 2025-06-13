import pandas as pd

# Crear datos para Excel - formato costeo
data = {
    "Concepto": ["Madera", "Clavos", "Horas de trabajo", "Electricidad"],
    "Cantidad": [10, 100, 20, 1],
    "Costeo Unitario (MXN)": [200, 1, 150, 300]
}
df_costeo = pd.DataFrame(data)
df_costeo["Total"] = df_costeo["Cantidad"] * df_costeo["Costeo Unitario (MXN)"]

# Cálculos adicionales
subtotal = df_costeo["Total"].sum()
margen = 0.30 # 30% de margen de ganacia
precio_venta = subtotal * (1 + margen)

# Guardar archivo Excel con plantilla
ruta_excel = "C:/Users/Alex/Desktop/plantilla_costeo_carpinteria.xlsx"

# Crar Excel con formateo
with pd.ExcelWriter(ruta_excel, engine="xlsxwriter") as writer:
    df_costeo.to_excel(writer, sheet_name="Costeo", index=False, startrow=0)

    workbook = writer.book
    worksheet = writer.sheets["Costeo"]

    # Formato moneda para columnas de costos y totales
    formato_moneda = workbook.add_format({"num_format": "$#,##0.00"})
    formato_cabecera = workbook.add_format({"bold": True, "bg_color": "#F0F0F0"})
    formato_porcentaje = workbook.add_format({"num_format": "0.00%"})

    worksheet.set_column("A:A", 25)
    worksheet.set_column("B:B", 12)
    worksheet.set_column("C:D", 20, formato_moneda)

    # Resumen
    fila_inicio_resumen = len(df_costeo) + 2
    worksheet.write(fila_inicio_resumen, 2, "Subtotal", formato_cabecera)
    worksheet.write(fila_inicio_resumen, 3, subtotal, formato_moneda)

    worksheet.write(fila_inicio_resumen + 1, 2, "Margen de ganancia", formato_cabecera)
    worksheet.write(fila_inicio_resumen + 1, 3, margen, formato_porcentaje)

    worksheet.write(fila_inicio_resumen + 2, 2, "Precio sugerido", formato_cabecera)
    worksheet.write(fila_inicio_resumen + 2, 3, precio_venta, formato_moneda)

print("✅ Archivo Excel creado exitosamente.")