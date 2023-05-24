import xlsxwriter
from datetime import datetime

class ClaseEstilos:
    #---------------------------------------FUNCIONES PARA CREAR LIBRO Y HOJA--------------------------------------------------#
    def crearLibro(self):
        #Fecha para nombrar cada archivo diferente y no sobre escriba datos, activar al final.
        fechaActual = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
        nombreArchivo = f"FormatoPM{fechaActual}.xlsx"

        #Creamos el libro de excel y le añadimos una hoja
        libro = xlsxwriter.Workbook("FormatoPM.xlsx")
        return libro
    
    def crearHoja(self, libro):
        hoja = libro.add_worksheet("Formato PM")
        return hoja

    #---------------------------------------FUNCIONES DE ESTILOS---------------------------------------------------------------#
    #Se agrega un fondo blanco a todo el documento de excel
    def agregarEstiloFondo(self, libro, hoja):
        formatoBlanco = libro.add_format({'bg_color': '#FFFFFF'})
        hoja.set_column(0, 105, None, formatoBlanco)
        hoja.set_row(0, None, formatoBlanco)

    #Estilo para generar el encabezado llamado ENCABEZADO del archivo
    def agregarEncabezadoEncabezado(self, libro, hoja):
        merge_range = 'A1:H1'

        formato = libro.add_format({'bg_color': "#FAEBD7", 'border': 1, 'align' : 'center', 'bold' : True})
        formato.set_font_name('Arial')
        formato.set_font_size(14)
        hoja.merge_range(merge_range, 'ENCABEZADO', formato)

    #Estilo para generar el encabezado llamado EMPRESA del archivo
    def agregarEncabezadoEmpresa(self, libro, hoja):
        merge_range = 'I1:AI1'

        formato = libro.add_format({'bg_color': "#DDA0DD", 'border': 1, 'align' : 'center', 'bold' : True})
        formato.set_font_name('Arial')
        formato.set_font_size(14)
        hoja.merge_range(merge_range, 'EMPRESA', formato)

    #Estilo para generar el encabezado llamado ACCIONISTA(OPCIONAL) del archivo
    def agregarEncabezadoAccionista(self, libro, hoja):
        merge_range = 'AJ1:BD1'

        formato = libro.add_format({'bg_color': "#AFEEEE", 'border': 1, 'align' : 'center', 'bold' : True})
        formato.set_font_name('Arial')
        formato.set_font_size(14)
        hoja.merge_range(merge_range, 'ACCIONISTA(OPCIONAL)', formato)

    #Estilo para generar el encabezado llamado CREDITO del archivo
    def agregarEncabezadoCredito(self, libro, hoja):
        merge_range = 'BE1:CC1'

        formato = libro.add_format({'bg_color': "#FA8072", 'border': 1, 'align' : 'center', 'bold' : True})
        formato.set_font_name('Arial')
        formato.set_font_size(14)
        hoja.merge_range(merge_range, 'CREDITO', formato)

    #Estilo para generar el encabezado llamado DETALLE CREDITO del archivo
    def agregarEncabezadoDetalleCredito(self, libro, hoja):
        merge_range = 'CD1:CH1'

        formato = libro.add_format({'bg_color': "#1E90FF", 'border': 1, 'align' : 'center', 'bold' : True})
        formato.set_font_name('Arial')
        formato.set_font_size(14)
        hoja.merge_range(merge_range, 'DETALLE CREDITO', formato)

    #Estilo para generar el encabezado llamado AVAL del archivo
    def agregarEncabezadoAval(self, libro, hoja):
        merge_range = 'CI1:DB1'

        formato = libro.add_format({'bg_color': "#90EE90", 'border': 1, 'align' : 'center', 'bold' : True})
        formato.set_font_name('Arial')
        formato.set_font_size(14)
        hoja.merge_range(merge_range, 'AVAL', formato)  

    #Se agrega el color amarillo a la celda que cuente con este color por default
    def agregarEstiloAmarillo(self, libro):
        formatoAmarillo = libro.add_format({'bg_color': "#FFFF00", 'border': 1, 'align' : 'center', 'bold' : True})
        formatoAmarillo.set_font_name('Arial')
        formatoAmarillo.set_font_size(10)
        return formatoAmarillo
    
    #Se agrega el color amarillo a la celda que cuente con información
    def agregarEstiloAmarilloInfo(self, libro):
        formatoAmarillo = libro.add_format({'bg_color': "#FFFF00", 'border': 1, 'align' : 'center', 'bold' : False})
        formatoAmarillo.set_font_name('Arial')
        formatoAmarillo.set_font_size(10)
        return formatoAmarillo

    #Se agrega el color azul cielo a la celda que lo necesite
    def agregarEstiloAzulClaro(self, libro):
        formatoAzulClaro = libro.add_format({'bg_color': "#A9F5F2", 'border': 1, 'align' : 'center', 'bold' : True})
        formatoAzulClaro.set_font_name('Arial')
        formatoAzulClaro.set_font_size(10)
        return formatoAzulClaro
    
    #Se agrega el color amarillo a la celda que cuente con información
    def agregarEstiloAzulClaroInfo(self, libro):
        formatoAzulClaroInfo = libro.add_format({'bg_color': "#A9F5F2", 'border': 1, 'align' : 'center', 'bold' : False})
        formatoAzulClaroInfo.set_font_name('Arial')
        formatoAzulClaroInfo.set_font_size(10)
        return formatoAzulClaroInfo

    #Se agrega el color azul fuerte a la celda que lo necesite
    def agregarEstiloAzulFuerte(self, libro):
        formatoAzulFuerte = libro.add_format({'bg_color': "#2E9AFE", 'border': 1, 'align' : 'center', 'bold' : True})
        formatoAzulFuerte.set_font_name('Arial')
        formatoAzulFuerte.set_font_size(10)
        return formatoAzulFuerte
    
    #Se agrega el color amarillo a la celda que cuente con información
    def agregarEstiloAzulFuerteInfo(self, libro):
        formatoAzulFuerteInfo = libro.add_format({'bg_color': "#2E9AFE", 'border': 1, 'align' : 'center', 'bold' : False})
        formatoAzulFuerteInfo.set_font_name('Arial')
        formatoAzulFuerteInfo.set_font_size(10)
        return formatoAzulFuerteInfo

    #Se agrega el color verde a la celda que lo necesite
    def agregarEstiloVerde(self, libro):
        formatoVerde = libro.add_format({'bg_color': "#81F79F", 'border': 1, 'align' : 'center', 'bold' : True})
        formatoVerde.set_font_name('Arial')
        formatoVerde.set_font_size(10)
        return formatoVerde
    
    #Se agrega el color amarillo a la celda que cuente con información
    def agregarEstiloVerdeInfo(self, libro):
        formatoVerdeInfo = libro.add_format({'bg_color': "#81F79F", 'border': 1, 'align' : 'center', 'bold' : False})
        formatoVerdeInfo.set_font_name('Arial')
        formatoVerdeInfo.set_font_size(10)
        return formatoVerdeInfo

    def agregarEstiloGris(self, libro):
        formatoGris = libro.add_format({'bg_color': "#A9D0F5", 'border': 1, 'align' : 'center', 'bold' : True})
        formatoGris.set_font_name('Arial')
        formatoGris.set_font_size(10)
        return formatoGris
    
        #Se agrega el color amarillo a la celda que cuente con información
    def agregarEstiloGrisInfo(self, libro):
        formatoGrisInfo = libro.add_format({'bg_color': "#A9D0F5", 'border': 1, 'align' : 'center', 'bold' : False})
        formatoGrisInfo.set_font_name('Arial')
        formatoGrisInfo.set_font_size(10)
        return formatoGrisInfo

    #Se añaden los colores amarillos por default a las casillas que lo necesiten.
    def agregarColumnasEncabezado(self, libro, hoja):
        #Casillas amarillas que por default tienen ese color.
        hoja.write(1, 0, "Otorgante", self.agregarEstiloAmarillo(libro))
        hoja.write(1, 2, "Nombre Otorgante", self.agregarEstiloAmarillo(libro))

        #Casillas azul claro que por default tinene ese color
        hoja.write(1, 1, "Otorgante Anterior", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 9, "CURP", self.agregarEstiloAzulClaro(libro))
        #hoja.write(1, 9, "Apellido paterno", agregarEstiloAzulClaro())
        hoja.write(1, 13, "Apellido paterno", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 14, "Apellido materno", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 18, "Clave Banxico 2/SCIAN", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 19, "Clave Banxico 3/SCIAN", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 27, "Telefono", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 28, "Extension",self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 29, "Fax", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 31, "Edo extranjero", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 36, "CURP", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 50, "Telefono", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 51, "Extension", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 52, "Fax", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 54, "Edo extranjero", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 55, "Pais", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 59, "Num contrato anterior", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 69, "Fecha reestructura", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 70, "Pago efectivo", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 71, "Fecha liquidacion", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 72, "Quita", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 73, "Dacion", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 74, "Quebranto o castigo", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 75, "Clave observacion", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 76, "Especiales", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 87, "CURP", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 100, "Telefono", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 101, "Extension", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 102, "Fax", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 104, "Edo extranjero", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 105, "Pais", self.agregarEstiloAzulClaro(libro))

        #Agregar color a celdas con color azul fuerte
        hoja.write(1, 3, "Institución", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 4, "Formato", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 5, "Fecha generacion", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 6, "Periodo que reporta", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 7, "Version", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 8, "RFC", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 10, "Compañia", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 11, "Nombre 1", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 12, "Nombre 2", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 16, "Calificacion cartera", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 17, "Clave Banxico 1/SCIAN", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 20, "Direccion 1", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 21, "Direccion 2", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 22, "Colonia", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 23, "Deleg/Mun", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 24, "Ciudad", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 25, "Estado", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 26, "CP", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 30, "Tipo cliente", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 32, "Pais", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 35, "RFC", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 37, "Nombre compañia", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 38, "Nombre 1", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 39, "Nombre 2", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 40, "Apellido Paterno Accionista", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 41, "Apellido Materno Accionista", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 42, "Porcentaje", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 43, "Dirección 1", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 44, "Direccion 2", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 45, "Colonia", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 46, "Deleg/Mun", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 47, "Ciudad", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 48, "Estado", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 49, "CP", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 53, "Tipo Cliente", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 56, "RFC", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 57, "Experiencias crediticias", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 58, "Num contrato", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 60, "Fecha apertura", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 62, "Tipo credito", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 63, "Saldo inicial", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 64, "Moneda", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 65, "Num pagos", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 68, "Fecha ultimo pago", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 77, "Fecha 1er incumplimiento", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 78, "Saldo Insoluto", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 81, "RFC", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 82, "Num contrato", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 83, "Dias vencidos", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 84, "Cantidad", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 85, "Intereses", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 86, "RFC", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 88, "Nombre Compañía", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 89, "Nombre 1", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 90, "Nombre 2", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 91, "Apellido Paterno Aval", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 92, "Apellido Materno Aval", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 93, "Direccion 1", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 94, "Direccion 2", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 95, "Colonia", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 96, "Deleg/Mun", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 97, "Ciudad", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 98, "Estado", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 99, "CP", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 103, "Tipo cliente", self.agregarEstiloAzulFuerte(libro))

        #Agregar color a celdas con color verde
        hoja.write(1, 15, "Nacionalidad", self.agregarEstiloVerde(libro))
        hoja.write(1, 61, "Plazo en meses", self.agregarEstiloVerde(libro))
        hoja.write(1, 66, "Frecuencia pagos", self.agregarEstiloVerde(libro))
        hoja.write(1, 67, "Importe pagos", self.agregarEstiloVerde(libro))
        hoja.write(1, 79, "Crédito máximo utilizado", self.agregarEstiloVerde(libro))

        #Agregar color a celdas grises
        hoja.write(1, 33, "Telefono movil", self.agregarEstiloGris(libro))
        hoja.write(1, 34, "Correo electronico", self.agregarEstiloGris(libro))
        hoja.write(1, 80, "Fecha Ingreso Cartera Vencida", self.agregarEstiloGris(libro))

        return hoja


