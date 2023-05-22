import xlsxwriter
from datetime import datetime
from estilos_documentoPM import ClaseEstilos

ce = ClaseEstilos()
libro = ce.crearLibro()
hoja = ce.crearHoja(libro)

#--------------------------------Funciones de procesamiento de datos----------------------------------------------
datos_encabezado = {
    "otorgante": [],
    "otorgante_anterior": [1, 3],
    "nombre_otorgante": ['Nombre otorgante 1'],
    "institucion": ['Institucion 1'],
    "formato": [1],
    "fecha": ['24/05/2023'],
    "periodo_reporta": [1],
    "version": [0]
}

valorMaximo = len(datos_encabezado["otorgante"])

# Ejemplo de función para almacenar los datos en las listas correspondientes
def almacenarDatosEncabezado(datos_encabezado, hoja):
    # Obtener los valores opcionales del diccionario
    otorgante = datos_encabezado.get("otorgante")
    otorgante_anterior = datos_encabezado.get("otorgante_anterior")
    nombre_otorgante = datos_encabezado.get("nombre_otorgante")
    institucion = datos_encabezado.get("institucion")
    formato = datos_encabezado.get("formato")
    fecha = datos_encabezado.get("fecha")
    periodo_reporta = datos_encabezado.get("periodo_reporta")
    version = datos_encabezado.get("version")

    # Verificar si los valores son None y asignar valores por defecto si es necesario
    if otorgante == []:
        otorgante = ['']  # Valor por defecto si no se proporciona

    if otorgante_anterior is None:
        otorgante_anterior = ['']

    if nombre_otorgante is None:
        nombre_otorgante = ['']

    if institucion is None:
        institucion = ['']

    if formato is None:
        formato = ['']

    if fecha is None:
        fecha = ['']

    if periodo_reporta is None:
        periodo_reporta = ['']

    if version is None:
        version = ['']

    fila = 2

    # Escribir los datos en las celdas correspondientes
    for valor in otorgante:
        hoja.write(fila, 0, valor, ce.agregarEstiloAmarilloInfo(libro))

    for valor in otorgante_anterior:
        hoja.write(fila, 1, valor, ce.agregarEstiloAzulClaroInfo(libro))

    for valor in nombre_otorgante:
        hoja.write(fila, 2, valor, ce.agregarEstiloAmarilloInfo(libro))

    for valor in institucion:
        hoja.write(fila, 3, valor, ce.agregarEstiloAzulFuerteInfo(libro))

    for valor in formato:
        hoja.write(fila, 4, valor, ce.agregarEstiloAzulFuerteInfo(libro))

    for valor in fecha:
        hoja.write(fila, 5, valor, ce.agregarEstiloAzulFuerteInfo(libro))

    for valor in periodo_reporta:
        hoja.write(fila, 6, valor, ce.agregarEstiloAzulFuerteInfo(libro))

    for valor in version:
        hoja.write(fila, 7, valor, ce.agregarEstiloAzulFuerteInfo(libro))

    fila += 1


# Ejemplo de uso
almacenarDatosEncabezado(datos_encabezado, hoja)
print(datos_encabezado)

#---------------------------------------LLAMADA DE FUNCIONES------------------------------------------------------
ce.agregarEstiloFondo(libro, hoja)
ce.agregarEncabezadoEncabezado(libro, hoja)
ce.agregarEncabezadoEmpresa(libro, hoja)
ce.agregarEncabezadoAccionista(libro, hoja)
ce.agregarEncabezadoCredito(libro, hoja)
ce.agregarEncabezadoDetalleCredito(libro, hoja)
ce.agregarEncabezadoAval(libro, hoja)
ce.agregarColumnasEncabezado(libro, hoja)

#Se le da un tamaño automático a cada celda con este método
hoja.autofit()
#Se cierra el libro
libro.close()
