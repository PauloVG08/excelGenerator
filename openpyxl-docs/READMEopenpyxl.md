# excelGenerator
Creación de módulo capaz de generar reportes en un archivo tipo excel.

# Openpyxl
Es un paquete de python que permite acceder a un archivo excel usando el lenguaje de python. 
Esta herramienta nos va permitir crear un archivo, manipularlo, extraer datos, insertar datos, 
crear gráficos, etc. Esto nos facilitará en usar una DATA de excel desde python.

# Hojas 
Por defecto al crear un archivo se crea una hoja llamada "sheet", nosotros podemos agregar
más hojas al archivo de la siguiente manera: 
wsb = wb.create_sheet("hoja1")
wsb = wb.create_sheet("hoja2", 0) donde el primer parametro es el nombre de la hoja y
                                  el segundo parámetro es la posición.

Podemos cambiar también su titulo de esta forma: 
ws2.title = "hoja1nueva"

Para asignar un color a la hoja es: 
ws2.sheet_properties.tabColor = "1072BA"

