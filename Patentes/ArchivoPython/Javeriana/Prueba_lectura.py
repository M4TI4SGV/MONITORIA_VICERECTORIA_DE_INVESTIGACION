import pandas as pd

# Cargar el archivo Excel sin leer una hoja específica
xls = pd.ExcelFile('pruebas.xlsx')

# Imprimir todas las hojas disponibles en el archivo Excel
print("Pestañas disponibles en el archivo Excel:", xls.sheet_names)
