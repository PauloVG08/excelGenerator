#Instanciamos la librería que vamos a ocupar
import xlsxwriter
from datetime import datetime
from itertools import zip_longest

#Usaremos un solo método llamado crearTablaDatos
def crearTablaDatos():
    #Obtenemos la fecha actual para nombrar de diferente manera todos los archivos que generemos
    fecha_hora_actual = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")

    #Esto se descomenta cuando queramos que genere un archivo cada que corramos el programa
    nombreArchivo = f"ventas{fecha_hora_actual}.xlsx"
    #Creamos el libro de excel que vamos a utilizar y le asignamos el nombre
    libro = xlsxwriter.Workbook("ventas.xlsx")

    #Asignamos una hoja a nuestro nuevo libro de excel
    hoja1 = libro.add_worksheet("productos")

    #---------------------------------- Diseño -----------------------------------------------------------------
    #Creamos distintos formatos de diseño según lo ocupamos
    formatoBlanco = libro.add_format({'bg_color': '#FFFFFF'})
    formatoTitulo = libro.add_format({
        'font_size': 24,
        'bold': True,
        'color': '#0000CD',
        'align': 'center',
        'bg_color': '#F2F2F2',
        'bottom': 1,
        'border_color': '#CCCCCC',
        'border': 1
    })
    formato1 = libro.add_format({'bg_color': "#6AA121", 'border': 1, 'align' : 'center', 'bold' : True})
    formato2 = libro.add_format({'bg_color': "#FBB917", 'border': 1, 'align' : 'center'})
    formato3 = libro.add_format({'color' : 'black', 'align' : 'center', 'bold' : True, 'bg_color': "#FBB917", 'border': 1})
    #hoja1.set_tab_color('#FFFFFF')
    hoja1.set_column(0, 70, None, formatoBlanco)
    hoja1.set_row(0, None, formatoBlanco)
    #-----------------------------------------------------------------------------------------------------------

    #------------------------------- Asignación de datos -------------------------------------------------------
    nombres_productos = ["Blusa azul", "Zapatos negros", "Sombrero claro", "Lentes oscuros", "Gorra"]
    precios_productos = [255, 378, 499.90, 768, 127.35]
    codigos_productos = [1, 2, 3, 4]
    codigos_sku = ["ABC-123", "DEF-456", "GHI-789", "JKL-012"]

    valorMaximo = max(len(nombres_productos), len(precios_productos), len(codigos_productos), len(codigos_sku))
    print(valorMaximo)

    #------------------------------------------------------------------------------------------------------------

    #-------------------- Le asignamos el título a la hoja ya usando los diseños previamente realizados ---------
    hoja1.write(1, 5, "Muestra de los productos", formatoTitulo)
    hoja1.write(1, 4, "", formatoTitulo)
    hoja1.write(1, 6, "", formatoTitulo)
    hoja1.write(1, 7, "", formatoTitulo)
    #------------------------------------------------------------------------------------------------------------

    #-------------Asignamos un diseño a los encabezados de la tabla ---------------------------------------------
    hoja1.write(3, 2, "Nombre del producto", formato1)
    hoja1.write(3, 3, "Código del producto", formato1)
    hoja1.write(3, 4, "SKU del producto", formato1)
    hoja1.write(3, 5, "Precio del producto", formato1)

    fecha_documento = datetime.now().strftime("%Y/%m/%d")
    hora_documento = datetime.now().strftime(" %H:%M")

    hoja1.write(1, 9, "Fecha en que se generó el reporte: ")
    hoja1.write(1, 10, fecha_documento)
    hoja1.write(2, 9, "Hora en que se generó el reporte: ")
    hoja1.write(2, 10, hora_documento)
    #------------------------------------------------------------------------------------------------------------

    #----------Creación de la tabla asignando los datos según se tengan -----------------------------------------
    fila = 4
    suma_total = 0.0
    contador = 1
  
    for i in range(valorMaximo):
        if len(nombres_productos) < valorMaximo:
            nombres_productos.extend([''] * (valorMaximo - len(nombres_productos)))
        
        if len(precios_productos) < valorMaximo:
            precios_productos.extend([0] * (valorMaximo - len(precios_productos)))

        if len(codigos_productos) < valorMaximo:
            codigos_productos.extend([''] * (valorMaximo - len(codigos_productos)))

        if len(codigos_sku) < valorMaximo:
            codigos_sku.extend([''] * (valorMaximo - len(codigos_sku)))

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
        
        fila += 1
        contador += 1
        if(contador > valorMaximo):
            hoja1.write(fila, 4, "Total",formato3)
            hoja1.write(fila, 5, suma_total, formato3)

    print("El precio total de los gastos es de: $" + str(suma_total))
    print("Valor de la fila: " + str(fila))
    #------------------------------------------------------------------------------------------------------------

    #-----------Con los datos de la tabla, creamos una gráfica de barras-----------------------------------------
    grafica = libro.add_chart({'type': 'column'})
    grafica.set_title({'name': 'Ventas por producto'})

    grafica.add_series({
    'values': f'=productos!$F$5:$F${fila}',
    'categories': f'=productos!$C$5:$C${fila}',
    'points': [
        {'fill': {'color': '#1E90FF'}},
        {'fill': {'color': '#DC143C'}},
        {'fill': {'color': '#006400'}},
        {'fill': {'color': '#FF8C00'}}
    ]
    })
    grafica.set_size({'width': 600, 'height': 400})
    grafica.set_x_axis({'name': 'Nombre del producto'})
    grafica.set_y_axis({'name': 'Precio', 'num_format': '$#,##0.00'})

    hoja1.insert_chart(f'C{fila + 4}', grafica)

    #----------Ahora se hace aquí una gráfica de pastel ------------------------------------
    graficaPastel = libro.add_chart({'type': 'pie'})
    graficaPastel.add_series({
        'name': 'Gráfica de pastel',
        'values': f'=productos!$F$5:$F${fila}',
        'categories': f'=productos!$C$5:$C${fila}',
        'points': [
            {'fill': {'color': '#1E90FF'}},
            {'fill': {'color': '#DC143C'}},
            {'fill': {'color': '#006400'}},
            {'fill': {'color': '#FF8C00'}}
    ]
    })
    graficaPastel.set_title({'name': 'Gráfica de pastel'})
    graficaPastel.set_size({'width': 600, 'height': 400})
    hoja1.insert_chart(f'I{ fila + 4 }', graficaPastel, {'x_offset': 25, 'y_offset': 10})

    #------------------------------------------------------------------------------------------------------------

    #Se le da un tamaño automático a cada celda con este método
    hoja1.autofit()
    #Cerramos el libro de excel
    libro.close()

#Llamamos la función
crearTablaDatos()

