# Iniciamos importando la librería
# Para instalarla se usa "pip install XlsxWriter" y aqui solo se importa
import xlsxwriter

# Creamos nuestro nuevo archivo de excel, aqui mismo se le asigna el nombre que tendrá.
libro = xlsxwriter.Workbook("ventas.xlsx")
#Aqui crearemos la hoja de excel y le asignamos su nombre
hoja1 = libro.add_worksheet("productos")

# En este apartado nos dedicamos a crear los estilos necesarios para la tabla
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
formato1 = libro.add_format({'bg_color': "#6AA121", 'border': 1})
formato2 = libro.add_format({'bg_color': "#FBB917", 'border': 1})

# Asignamos algunos productos y sus caracteristicas dentro de arreglos
nombres_productos = ["Blusa azul", "Zapatos negros", "Sombrero claro", "Lentes oscuros"]
codigos_productos = [1, 2, 3, 4]
codigos_sku = ["ABC-123", "DEF-456", "GHI-789", "JKL-012"]

#Asignamos el título que se mostrará en el excel y le añadimos su estilo.
hoja1.write(1, 5, "Muestra de los productos", formatoTitulo)

# Asignamos el formato 1 a los encabezados de la tabla
hoja1.write(3, 2, "Nombre del producto", formato1)
hoja1.write(3, 3, "Código del producto", formato1)
hoja1.write(3, 4, "SKU del producto", formato1)

#Se inicia en 4 la tabla pues yo elegí la fila en que quería que iniciara
fila = 4  # Iniciar en la segunda fila (fila 1)


for i in range(len(nombres_productos)):
    nombre = nombres_productos[i]
    codigo = codigos_productos[i]
    sku = codigos_sku[i]
    print(f"El producto {nombre} es el número {codigo} y tiene código SKU {sku}")
    
    hoja1.write(fila, 2, nombre, formato2)
    hoja1.write(fila, 3, codigo, formato2)
    hoja1.write(fila, 4, sku, formato2)
    
    fila += 1  # Aumentar el valor de la fila en cada iteración

libro.close()
