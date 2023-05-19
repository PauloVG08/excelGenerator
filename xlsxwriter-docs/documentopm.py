import xlsxwriter
from datetime import datetime

#Fecha para nombrar cada archivo diferente y no sobre escriba datos, activar al final.
fechaActual = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
nombreArchivo = f"FormatoPM{fechaActual}.xlsx"

#Creamos el libro de excel y le añadimos una hoja
libro = xlsxwriter.Workbook("FormatoPM.xlsx")
hoja = libro.add_worksheet("Formato PM")

#---------------------------------------FUNCIONES DE ESTILOS---------------------------------------------------------------#
#Se agrega un fondo blanco a todo el documento de excel
def agregarEstiloFondo():
    formatoBlanco = libro.add_format({'bg_color': '#FFFFFF'})
    hoja.set_column(0, 105, None, formatoBlanco)
    hoja.set_row(0, None, formatoBlanco)

#Estilo para generar el encabezado llamado ENCABEZADO del archivo
def agregarEncabezadoEncabezado():
    merge_range = 'A1:H1'

    formato = libro.add_format({'bg_color': "#FAEBD7", 'border': 1, 'align' : 'center', 'bold' : True})
    formato.set_font_name('Arial')
    formato.set_font_size(14)
    hoja.merge_range(merge_range, 'ENCABEZADO', formato)

#Estilo para generar el encabezado llamado EMPRESA del archivo
def agregarEncabezadoEmpresa():
    merge_range = 'I1:AI1'

    formato = libro.add_format({'bg_color': "#DDA0DD", 'border': 1, 'align' : 'center', 'bold' : True})
    formato.set_font_name('Arial')
    formato.set_font_size(14)
    hoja.merge_range(merge_range, 'EMPRESA', formato)

#Estilo para generar el encabezado llamado ACCIONISTA(OPCIONAL) del archivo
def agregarEncabezadoAccionista():
    merge_range = 'AJ1:BD1'

    formato = libro.add_format({'bg_color': "#AFEEEE", 'border': 1, 'align' : 'center', 'bold' : True})
    formato.set_font_name('Arial')
    formato.set_font_size(14)
    hoja.merge_range(merge_range, 'ACCIONISTA(OPCIONAL)', formato)

#Estilo para generar el encabezado llamado CREDITO del archivo
def agregarEncabezadoCredito():
    merge_range = 'BE1:CC1'

    formato = libro.add_format({'bg_color': "#FA8072", 'border': 1, 'align' : 'center', 'bold' : True})
    formato.set_font_name('Arial')
    formato.set_font_size(14)
    hoja.merge_range(merge_range, 'CREDITO', formato)

#Estilo para generar el encabezado llamado DETALLE CREDITO del archivo
def agregarEncabezadoDetalleCredito():
    merge_range = 'CD1:CH1'

    formato = libro.add_format({'bg_color': "#1E90FF", 'border': 1, 'align' : 'center', 'bold' : True})
    formato.set_font_name('Arial')
    formato.set_font_size(14)
    hoja.merge_range(merge_range, 'DETALLE CREDITO', formato)

#Estilo para generar el encabezado llamado AVAL del archivo
def agregarEncabezadoAval():
    merge_range = 'CI1:DB1'

    formato = libro.add_format({'bg_color': "#90EE90", 'border': 1, 'align' : 'center', 'bold' : True})
    formato.set_font_name('Arial')
    formato.set_font_size(14)
    hoja.merge_range(merge_range, 'AVAL', formato)  

#Se agrega el color amarillo a la celda que lo necesite
def agregarEstiloAmarillo():
    formatoAmarillo = libro.add_format({'bg_color': "#FFFF00", 'border': 1, 'align' : 'center', 'bold' : True})
    formatoAmarillo.set_font_name('Arial')
    formatoAmarillo.set_font_size(10)
    return formatoAmarillo

#Se agrega el color azul cielo a la celda que lo necesite
def agregarEstiloAzulClaro():
    formatoAzulClaro = libro.add_format({'bg_color': "#A9F5F2", 'border': 1, 'align' : 'center', 'bold' : True})
    formatoAzulClaro.set_font_name('Arial')
    formatoAzulClaro.set_font_size(10)
    return formatoAzulClaro

#Se agrega el color azul fuerte a la celda que lo necesite
def agregarEstiloAzulFuerte():
    formatoAzulFuerte = libro.add_format({'bg_color': "#2E9AFE", 'border': 1, 'align' : 'center', 'bold' : True})
    formatoAzulFuerte.set_font_name('Arial')
    formatoAzulFuerte.set_font_size(10)
    return formatoAzulFuerte

#Se agrega el color verde a la celda que lo necesite
def agregarEstiloVerde():
    formatoVerde = libro.add_format({'bg_color': "#81F79F", 'border': 1, 'align' : 'center', 'bold' : True})
    formatoVerde.set_font_name('Arial')
    formatoVerde.set_font_size(10)
    return formatoVerde

def agregarEstiloGris():
    formatoGris = libro.add_format({'bg_color': "#A9D0F5", 'border': 1, 'align' : 'center', 'bold' : True})
    formatoGris.set_font_name('Arial')
    formatoGris.set_font_size(10)
    return formatoGris

#Se añaden los colores amarillos por default a las casillas que lo necesiten.
def agregarColumnasEncabezado():
    #Casillas amarillas que por default tienen ese color.
    hoja.write(1, 0, "Otorgante", agregarEstiloAmarillo())
    hoja.write(1, 2, "Nombre Otorgante", agregarEstiloAmarillo())

    #Casillas azul claro que por default tinene ese color
    hoja.write(1, 1, "Otorgante Anterior", agregarEstiloAzulClaro())
    hoja.write(1, 9, "CURP", agregarEstiloAzulClaro())
    #hoja.write(1, 9, "Apellido paterno", agregarEstiloAzulClaro())
    hoja.write(1, 13, "Apellido paterno", agregarEstiloAzulClaro())
    hoja.write(1, 14, "Apellido materno", agregarEstiloAzulClaro())
    hoja.write(1, 18, "Clave Banxico 2/SCIAN", agregarEstiloAzulClaro())
    hoja.write(1, 19, "Clave Banxico 3/SCIAN", agregarEstiloAzulClaro())
    hoja.write(1, 27, "Telefono", agregarEstiloAzulClaro())
    hoja.write(1, 28, "Extencion", agregarEstiloAzulClaro())
    hoja.write(1, 29, "Fax", agregarEstiloAzulClaro())
    hoja.write(1, 31, "Edo extranjero", agregarEstiloAzulClaro())
    hoja.write(1, 36, "CURP", agregarEstiloAzulClaro())
    hoja.write(1, 50, "Telefono", agregarEstiloAzulClaro())
    hoja.write(1, 51, "Extencion", agregarEstiloAzulClaro())
    hoja.write(1, 52, "Fax", agregarEstiloAzulClaro())
    hoja.write(1, 54, "Edo extranjero", agregarEstiloAzulClaro())
    hoja.write(1, 55, "Pais", agregarEstiloAzulClaro())
    hoja.write(1, 59, "Num contrato anterior", agregarEstiloAzulClaro())
    hoja.write(1, 69, "Fecha reestructura", agregarEstiloAzulClaro())
    hoja.write(1, 70, "Pago efectivo", agregarEstiloAzulClaro())
    hoja.write(1, 71, "Fecha liquidacion", agregarEstiloAzulClaro())
    hoja.write(1, 72, "Quita", agregarEstiloAzulClaro())
    hoja.write(1, 73, "Dacion", agregarEstiloAzulClaro())
    hoja.write(1, 74, "Quebranto o castigo", agregarEstiloAzulClaro())
    hoja.write(1, 75, "Clave observacion", agregarEstiloAzulClaro())
    hoja.write(1, 76, "Especiales", agregarEstiloAzulClaro())
    hoja.write(1, 87, "CURP", agregarEstiloAzulClaro())
    hoja.write(1, 100, "Telefono", agregarEstiloAzulClaro())
    hoja.write(1, 101, "Extension", agregarEstiloAzulClaro())
    hoja.write(1, 102, "Fax", agregarEstiloAzulClaro())
    hoja.write(1, 104, "Edo extranjero", agregarEstiloAzulClaro())
    hoja.write(1, 105, "Pais", agregarEstiloAzulClaro())

    #Agregar color a celdas con color azul fuerte
    hoja.write(1, 3, "Institución", agregarEstiloAzulFuerte())
    hoja.write(1, 4, "Formato", agregarEstiloAzulFuerte())
    hoja.write(1, 5, "Fecha generacion", agregarEstiloAzulFuerte())
    hoja.write(1, 6, "Periodo que reporta", agregarEstiloAzulFuerte())
    hoja.write(1, 7, "Version", agregarEstiloAzulFuerte())
    hoja.write(1, 8, "RFC", agregarEstiloAzulFuerte())
    hoja.write(1, 10, "Compañia", agregarEstiloAzulFuerte())
    hoja.write(1, 11, "Nombre 1", agregarEstiloAzulFuerte())
    hoja.write(1, 12, "Nombre 2", agregarEstiloAzulFuerte())
    hoja.write(1, 16, "Calificacion cartera", agregarEstiloAzulFuerte())
    hoja.write(1, 17, "Clave Banxico 1/SCIAN", agregarEstiloAzulFuerte())
    hoja.write(1, 20, "Direccion 1", agregarEstiloAzulFuerte())
    hoja.write(1, 21, "Direccion 2", agregarEstiloAzulFuerte())
    hoja.write(1, 22, "Colonia", agregarEstiloAzulFuerte())
    hoja.write(1, 23, "Deleg/Mun", agregarEstiloAzulFuerte())
    hoja.write(1, 24, "Ciudad", agregarEstiloAzulFuerte())
    hoja.write(1, 25, "Estado", agregarEstiloAzulFuerte())
    hoja.write(1, 26, "CP", agregarEstiloAzulFuerte())
    hoja.write(1, 30, "Tipo cliente", agregarEstiloAzulFuerte())
    hoja.write(1, 32, "Pais", agregarEstiloAzulFuerte())
    hoja.write(1, 35, "RFC", agregarEstiloAzulFuerte())
    hoja.write(1, 37, "Nombre compañia", agregarEstiloAzulFuerte())
    hoja.write(1, 38, "Nombre 1", agregarEstiloAzulFuerte())
    hoja.write(1, 39, "Nombre 2", agregarEstiloAzulFuerte())
    hoja.write(1, 40, "Apellido Paterno Accionista", agregarEstiloAzulFuerte())
    hoja.write(1, 41, "Apellido Materno Accionista", agregarEstiloAzulFuerte())
    hoja.write(1, 42, "Porcentaje", agregarEstiloAzulFuerte())
    hoja.write(1, 43, "Dirección 1", agregarEstiloAzulFuerte())
    hoja.write(1, 44, "Direccion 2", agregarEstiloAzulFuerte())
    hoja.write(1, 45, "Colonia", agregarEstiloAzulFuerte())
    hoja.write(1, 46, "Deleg/Mun", agregarEstiloAzulFuerte())
    hoja.write(1, 47, "Ciudad", agregarEstiloAzulFuerte())
    hoja.write(1, 48, "Estado", agregarEstiloAzulFuerte())
    hoja.write(1, 49, "CP", agregarEstiloAzulFuerte())
    hoja.write(1, 53, "Tipo Cliente", agregarEstiloAzulFuerte())
    hoja.write(1, 56, "RFC", agregarEstiloAzulFuerte())
    hoja.write(1, 57, "Experiencias crediticias", agregarEstiloAzulFuerte())
    hoja.write(1, 58, "Num contrato", agregarEstiloAzulFuerte())
    hoja.write(1, 60, "Fecha apertura", agregarEstiloAzulFuerte())
    hoja.write(1, 62, "Tipo credito", agregarEstiloAzulFuerte())
    hoja.write(1, 63, "Saldo inicial", agregarEstiloAzulFuerte())
    hoja.write(1, 64, "Moneda", agregarEstiloAzulFuerte())
    hoja.write(1, 65, "Num pagos", agregarEstiloAzulFuerte())
    hoja.write(1, 68, "Fecha ultimo pago", agregarEstiloAzulFuerte())
    hoja.write(1, 77, "Fecha 1er incumplimiento", agregarEstiloAzulFuerte())
    hoja.write(1, 78, "Saldo Insoluto", agregarEstiloAzulFuerte())
    hoja.write(1, 81, "RFC", agregarEstiloAzulFuerte())
    hoja.write(1, 82, "Num contrato", agregarEstiloAzulFuerte())
    hoja.write(1, 83, "Dias vencidos", agregarEstiloAzulFuerte())
    hoja.write(1, 84, "Cantidad", agregarEstiloAzulFuerte())
    hoja.write(1, 85, "Intereses", agregarEstiloAzulFuerte())
    hoja.write(1, 86, "RFC", agregarEstiloAzulFuerte())
    hoja.write(1, 88, "Nombre Compañía", agregarEstiloAzulFuerte())
    hoja.write(1, 89, "Nombre 1", agregarEstiloAzulFuerte())
    hoja.write(1, 90, "Nombre 2", agregarEstiloAzulFuerte())
    hoja.write(1, 91, "Apellido Paterno Aval", agregarEstiloAzulFuerte())
    hoja.write(1, 92, "Apellido Materno Aval", agregarEstiloAzulFuerte())
    hoja.write(1, 93, "Direccion 1", agregarEstiloAzulFuerte())
    hoja.write(1, 94, "Direccion 2", agregarEstiloAzulFuerte())
    hoja.write(1, 95, "Colonia", agregarEstiloAzulFuerte())
    hoja.write(1, 96, "Deleg/Mun", agregarEstiloAzulFuerte())
    hoja.write(1, 97, "Ciudad", agregarEstiloAzulFuerte())
    hoja.write(1, 98, "Estado", agregarEstiloAzulFuerte())
    hoja.write(1, 99, "CP", agregarEstiloAzulFuerte())
    hoja.write(1, 103, "Tipo cliente", agregarEstiloAzulFuerte())

    #Agregar color a celdas con color verde
    hoja.write(1, 15, "Nacionalidad", agregarEstiloVerde())
    hoja.write(1, 61, "Plazo en meses", agregarEstiloVerde())
    hoja.write(1, 66, "Frecuencia pagos", agregarEstiloVerde())
    hoja.write(1, 67, "Importe pagos", agregarEstiloVerde())
    hoja.write(1, 79, "Crédito máximo utilizado", agregarEstiloVerde())

    #Agregar color a celdas grises
    hoja.write(1, 33, "Telefono movil", agregarEstiloVerde())
    hoja.write(1, 34, "Correo electronico", agregarEstiloVerde())
    hoja.write(1, 80, "Fecha Ingreso Cartera Vencida", agregarEstiloVerde())


#---------------------------------------LLAMADA DE FUNCIONES-----------------------------------------------
agregarEstiloFondo()
agregarEncabezadoEncabezado()
agregarEncabezadoEmpresa()
agregarEncabezadoAccionista()
agregarEncabezadoCredito()
agregarEncabezadoDetalleCredito()
agregarEncabezadoAval()
agregarColumnasEncabezado()

#Se le da un tamaño automático a cada celda con este método
hoja.autofit()
#Se cierra el libro
libro.close()
