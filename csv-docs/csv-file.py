import pandas as pd

encabezados = ['Nombre', 'Habitantes', 'Recomendable', 'Pedestal']
ciudades = ['Leon', 'Ciudad de Mexico', 'Guadalajara', 'Monterrey']
habitantes = [1250728, 575, 3125741, 745]
recomendables = ['Si', 'Si', 'No', 'No']
pedestales = ['Alto', 'Alto', 'Bajo', 'Bajo']

datos = {
    'Nombre': ciudades,
    'Habitantes': habitantes,
    'Recomendable': recomendables,
    'Pedestal': pedestales
}

nombre_archivo = 'ciudades.csv'

df = pd.DataFrame(datos)

df.to_csv(nombre_archivo, index=False)

print("Archivo CSV generado correctamente.")
