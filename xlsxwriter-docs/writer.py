# Iniciamos importando la librería
# Para instalarla se usa "pip install XlsxWriter" y aqui solo se importa
import xlsxwriter
# se necesitará una librería para nombrar los archivos
from datetime import datetime

def crearTablaDatos():
    # Obtener la fecha y hora actual
    fecha_hora_actual = datetime.now().strftime("%Y_%m_%d_%H_%M")

    #------------------- fecha actual para nombrar archivos ---------------
    # Crear un nombre de archivo único utilizando la fecha y hora actual
    #nombre_archivo = f"ventas_{fecha_hora_actual}.xlsx"
    # ---------------------------------------------------------------------

    # Creamos nuestro nuevo archivo de excel, aqui mismo se le asigna el nombre que tendrá.
    libro = xlsxwriter.Workbook("ventas.xlsx")
    #Aqui crearemos la hoja de excel y le asignamos su nombre
    hoja1 = libro.add_worksheet("productos")

    # Establecer el formato para el fondo blanco
    formatoBlanco = libro.add_format({'bg_color': '#FFFFFF'})
    # En este apartado nos dedicamos a crear los estilos necesarios para la tabla y archivo en general
    formatoTitulo = libro.add_format({
        'font_size': 24,
        'bold': True,
        'color': '#333333',
        'align': 'center',
        'bg_color': '#F2F2F2',
        'bottom': 1,
        'border_color': '#CCCCCC',
        'border': 1
    })
    formato1 = libro.add_format({'bg_color': "#6AA121", 'border': 1, 'align' : 'center', 'bold' : True})
    formato2 = libro.add_format({'bg_color': "#FBB917", 'border': 1, 'align' : 'center'})
    formato3 = libro.add_format({'color' : 'black', 'align' : 'center', 'bold' : True, 'bg_color': "#FBB917", 'border': 1})

    # Asignamos algunos productos y sus caracteristicas dentro de arreglos
    nombres_productos = ["Blusa azul", "Zapatos negros", "Sombrero claro", "Lentes oscuros"]
    precios_productos = [255, 378, 499.90, 768]
    codigos_productos = [1, 2, 3, 4]
    codigos_sku = ["ABC-123", "DEF-456", "GHI-789", "JKL-012"]

    #Asignamos el título que se mostrará en el excel y le añadimos su estilo.
    hoja1.write(1, 5, "Muestra de los productos", formatoTitulo)

    # Asignamos el formato blanco al área del documento
    hoja1.set_tab_color('#FFFFFF')
    hoja1.set_column(0, 70, None, formatoBlanco)
    hoja1.set_row(0, None, formatoBlanco)

    # Asignamos el formato 1 a los encabezados de la tabla
    hoja1.write(3, 2, "Nombre del producto", formato1)
    hoja1.write(3, 3, "Código del producto", formato1)
    hoja1.write(3, 4, "SKU del producto", formato1)
    hoja1.write(3, 5, "Precio del producto", formato1)

    #Se inicia en 4 la tabla pues yo elegí la fila en que quería que iniciara
    fila = 4  # Iniciar en la segunda fila (fila 1)
    suma_total = 0.0
    contador = 1

    for i in range(len(nombres_productos)):
        nombre = nombres_productos[i]
        codigo = codigos_productos[i]
        sku = codigos_sku[i]
        precio = precios_productos[i]
        suma_total += precios_productos[i]
        print(f"El producto {nombre} es el número {codigo} y tiene código SKU {sku}")
        
        hoja1.write(fila, 2, nombre, formato2)
        hoja1.write(fila, 3, codigo, formato2)
        hoja1.write(fila, 4, sku, formato2)
        hoja1.write(fila, 5, precio, formato2)
        
        fila += 1  # Aumentar el valor de la fila en cada iteración
        contador += 1
        if(contador > len(nombres_productos)):
            hoja1.write(fila, 4, "Total", formato3)
            hoja1.write(fila, 5, suma_total, formato3)

    # Esto solo es para ver los resultados en consola
    print("El precio total de los gastos es de: $" + str(suma_total))
    #Este método sirve para ajustar el tamaño automático de cada una de las celdas
    hoja1.autofit()
    libro.close()

crearTablaDatos()
# ---------------------------------------------------------------------------------------------------------
