import openpyxl

# Carga el archivo de trabajo
workbook = openpyxl.load_workbook("BARRANQUERAS A CONCEPCION.xlsx")

# Selecciona la hoja activa
hoja = workbook["Hoja1"]

# Crea un archivo de texto para guardar los datos de la fila
with open("BARRANQUERAS.txt", "w") as archivo_txt:
    # Itera sobre las filas en la hoja
    for fila in hoja.iter_rows(values_only=True):
        # Imprime los valores en cada celda de la fila
        if "BARRANQUERAS" in fila:
            # Si se encuentra, agrega la fila al archivo de texto
            archivo_txt.write(",".join([str(x) for x in fila]))
            archivo_txt.write("\n")

# Cierra el libro original
workbook.close()
