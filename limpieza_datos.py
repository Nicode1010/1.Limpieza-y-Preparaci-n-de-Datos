import pandas as pd

# Cargar el archivo Excel
archivo = "estudiantes.xlsx"  # Asegúrate de que este archivo esté en la misma carpeta
df = pd.read_excel(archivo)

# Mostrar información básica del dataset
print("Primeras filas del archivo:")
print(df.head())

print("\nInformación general del dataset:")
print(df.info())

print("\nValores nulos por columna:")
print(df.isnull().sum())