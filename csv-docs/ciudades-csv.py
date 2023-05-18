import csv
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

def crearDocumento():
    # Datos de ejemplo
    productos = ['Camisa', 'Blusa', 'Zapatos', 'Lentes', 'Pantalón', 'Otro', 'Gorra', 'si']
    precios = [145.75, 99.90, 742.8, 500, 125, 299]
    codigo_de_producto = [1, 2, 3, 4, 5, 6, 7, 8]
    estatus = ['Activo', 'Activo', 'Inactivo', 'Activo', 'Inactivo', 'Activo', 'Inactivo', 'no']

    # Obtener el tamaño máximo
    valor_maximo = max(len(productos), len(precios), len(codigo_de_producto), len(estatus))

    # Rellenar las listas con campos vacíos si es necesario
    productos.extend([''] * (valor_maximo - len(productos)))
    precios.extend([0] * (valor_maximo - len(precios)))
    codigo_de_producto.extend([0] * (valor_maximo - len(codigo_de_producto)))
    estatus.extend([''] * (valor_maximo - len(estatus)))

    # Nombre del archivo CSV
    nombre_archivo_csv = 'ciudades.csv'
    suma = 0.0

    # Crear y escribir en el archivo CSV
    with open(nombre_archivo_csv, 'w', newline='') as archivo_csv:
        escritor_csv = csv.writer(archivo_csv)

        # Escribir encabezados
        escritor_csv.writerow(['Nombre del producto', 'Precio', 'Código producto', 'Estatus'])

        # Escribir los datos
        for i, (producto, precio, codigo, est) in enumerate(zip(productos, precios, codigo_de_producto, estatus), start=4):
            escritor_csv.writerow([producto, precio, codigo, est])
            
        for i, precio in enumerate(precios):
            suma += precios[i]

        # Escribir la suma total
        escritor_csv.writerow(['Suma total', suma])

    print("Archivo CSV generado correctamente.")

crearDocumento()
