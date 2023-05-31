import os
import xlsxwriter
from datetime import datetime

class ClaseEstilos:
    #---------------------------------------FUNCIONES PARA CREAR LIBRO Y HOJA--------------------------------------------------#
    def crearLibro(self):
        #Fecha para nombrar cada archivo diferente y no sobre escriba datos, activar al final.
        fechaActual = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
        #nombreArchivo = f"FormatoPF{fechaActual}.xlsx"

        #Creamos el libro de excel y le añadimos una hoja
        libro = xlsxwriter.Workbook("FormatoPF.xlsx")
        #libro = xlsxwriter.Workbook(nombreArchivo)
        #ruta_descargas = os.path.join(os.path.expanduser("~"), "Downloads", "FormatoPM.xlsx")
        #ruta_descargas = os.path.join(os.path.expanduser("~"), "Downloads", nombreArchivo)
        #libro = xlsxwriter.Workbook(ruta_descargas)
        #print(f"El archivo se ha guardado en: {ruta_descargas}")
        return libro
    
    # def crearLibro(self):
    #     fechaActual = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
    #     nombreArchivo = f"FormatoPF{fechaActual}.xlsx"

    #     ruta_guardado = input("Ingrese la ubicación donde desea guardar el archivo: ")

    #     if not os.path.exists(ruta_guardado):
    #         print("La ubicación especificada no existe. El archivo se guardará en la ubicación predeterminada.")
    #         ruta_guardado = os.path.join(os.path.expanduser("~"), "Downloads")

    #     ruta_archivo = os.path.join(ruta_guardado, nombreArchivo)

    #     libro = xlsxwriter.Workbook(ruta_archivo)

    #     print(f"El archivo se ha guardado en: {ruta_archivo}")

    #     return libro
    
    def crearHoja(self, libro):
        hoja = libro.add_worksheet("Formato PF")
        return hoja
    
    #---------------------------------------FUNCIONES DE ESTILOS---------------------------------------------------------------#
     #Se agrega un fondo blanco a todo el documento de excel
    def agregarEstiloFondo(self, libro, hoja):
        formatoBlanco = libro.add_format({'bg_color': '#FFFFFF'})
        hoja.set_column(0, 105, None, formatoBlanco)
        hoja.set_row(0, None, formatoBlanco)

    #Estilo para generar el encabezado llamado ENCABEZADO del archivo
    def agregarEncabezadoEncabezado(self, libro, hoja):
        merge_range = 'A1:F1'

        formato = libro.add_format({'bg_color': "#1E4EDB", 'border': 1, 'align' : 'center', 'bold' : True, 'color': '#FFFFFF'})
        formato.set_font_name('Arial')
        formato.set_font_size(14)
        hoja.merge_range(merge_range, 'E N C A B E Z A D O', formato)

    #Estilo para generar el encabezado llamado DATOS PERSONALES del archivo
    def agregarEncabezadoDatoPersonales(self, libro, hoja):
        merge_range = 'G1:X1'

        formato = libro.add_format({'bg_color': "#C0C1C7", 'border': 1, 'align' : 'center', 'bold' : True, 'color': '#FFFFFF'})
        formato.set_font_name('Arial')
        formato.set_font_size(14)
        hoja.merge_range(merge_range, 'D A T O S  P E R S O N A L E S', formato)

    #Estilo para generar el encabezado llamado DOMICILIO del archivo
    def agregarEncabezadoDomicilio(self, libro, hoja):
        merge_range = 'Y1:AI1'

        formato = libro.add_format({'bg_color': "#8490D3", 'border': 1, 'align' : 'center', 'bold' : True, 'color': '#FFFFFF'})
        formato.set_font_name('Arial')
        formato.set_font_size(14)
        hoja.merge_range(merge_range, 'D O M I C I L I O', formato)

    #Estilo para generar el encabezado llamado EMPLEO del archivo
    def agregarEncabezadoEmpleo(self, libro, hoja):
        merge_range = 'AJ1:AZ1'

        formato = libro.add_format({'bg_color': "#C0C1C7", 'border': 1, 'align' : 'center', 'bold' : True, 'color': '#FFFFFF'})
        formato.set_font_name('Arial')
        formato.set_font_size(14)
        hoja.merge_range(merge_range, 'E M P L E O', formato)

    #Estilo para generar el encabezado llamado DETALLE DE LA CUENTA del archivo
    def agregarEncabezadoDetalleCuenta(self, libro, hoja):
        merge_range = 'BA1:CM1'

        formato = libro.add_format({'bg_color': "#8490D3", 'border': 1, 'align' : 'center', 'bold' : True, 'color': '#FFFFFF'})
        formato.set_font_name('Arial')
        formato.set_font_size(14)
        hoja.merge_range(merge_range, 'D E T A L L E  D E  L A  C U E N T A', formato)

    #Estilo para generar el encabezado llamado CIFRAS DE CONTROL del archivo
    def agregarEncabezadoCifrasControl(self, libro, hoja):
        merge_range = 'CN1:CU1'

        formato = libro.add_format({'bg_color': "#C0C1C7", 'border': 1, 'align' : 'center', 'bold' : True, 'color': '#FFFFFF'})
        formato.set_font_name('Arial')
        formato.set_font_size(14)
        hoja.merge_range(merge_range, 'C I F R A S  D E  C O N T R O L', formato)

    #Se agrega el color naranja a la celda que cuente con este color por default
    def agregarEstiloNaranja(self, libro):
        formato = libro.add_format({'bg_color': "#EFBA89", 'border': 1, 'align' : 'center', 'bold' : True})
        formato.set_font_name('Arial')
        formato.set_font_size(10)
        return formato

    #Se agrega el color naranja a la celda que cuente con este color por default
    def agregarEstiloNaranjaInfo(self, libro):
        formato = libro.add_format({'bg_color': "#EFBA89", 'border': 1, 'align' : 'center'})
        formato.set_font_name('Arial')
        formato.set_font_size(10)
        return formato

    #Se agrega el color azul claro a la celda que cuente con este color por default
    def agregarEstiloAzulClaro(self, libro):
        formato = libro.add_format({'bg_color': "#96EFEF", 'border': 1, 'align' : 'center', 'bold' : True})
        formato.set_font_name('Arial')
        formato.set_font_size(10)
        return formato

    #Se agrega el color azul Claro a la celda con informacion que cuente con este color por default
    def agregarEstiloAzulClaroInfo(self, libro):
        formato = libro.add_format({'bg_color': "#96EFEF", 'border': 1, 'align' : 'center'})
        formato.set_font_name('Arial')
        formato.set_font_size(10)
        return formato

    #Se agrega el color gris a la celda que cuente con este color por default
    def agregarEstiloAzulFuerte(self, libro):
        formato = libro.add_format({'bg_color': "#80B7DB", 'border': 1, 'align' : 'center', 'bold' : True})
        formato.set_font_name('Arial')
        formato.set_font_size(10)
        return formato
    
    #Se agrega el color gris a la celda con informacion que cuente con este color por default
    def agregarEstiloAzulFuerteInfo(self, libro):
        formato = libro.add_format({'bg_color': "#80B7DB", 'border': 1, 'align' : 'center'})
        formato.set_font_name('Arial')
        formato.set_font_size(10)
        return formato
    
    #Se agrega el color amarillo a la celda que cuente con este color por default
    def agregarEstiloAmarillo(self, libro):
        formato = libro.add_format({'bg_color': "#FAF794", 'border': 1, 'align' : 'center', 'bold' : True})
        formato.set_font_name('Arial')
        formato.set_font_size(10)
        return formato
    
    #Se agrega el color amarillo con información a la celda que cuente con este color por default
    def agregarEstiloAmarilloInfo(self, libro):
        formato = libro.add_format({'bg_color': "#FAF794", 'border': 1, 'align' : 'center'})
        formato.set_font_name('Arial')
        formato.set_font_size(10)
        return formato
    
    def agregarEstiloBlancoCelda(self, libro):
        formato = libro.add_format({'bg_color': "#FFFFFF"})
        return formato

    #Se añaden los subtitulos
    def agregarColumnasEncabezado(self, libro, hoja):
        #Celdas con color naranja
        hoja.write(1, 0, "Clave otorgante", self.agregarEstiloNaranja(libro))
        hoja.write(1, 1, "Nombre otorgante", self.agregarEstiloNaranja(libro))

        #Celdas con color azul Claro
        hoja.write(1, 2, "Identificador de medio", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 4, "Nota otorgante", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 8, "Apellido adicional", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 12, "CURP", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 13, "Número Seguridad Social", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 15, "Residencia", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 16, "Número Licencia Conducir", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 17, "Estado civil", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 18, "Sexo", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 19, "Clave Electoral IFE", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 20, "Número dependientes", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 21, "Fecha de función", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 22, "Indicador de función", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 23, "Tipo persona", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 30, "Fecha residencia", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 31, "Número teléfono", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 32, "Tipo domicilio", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 33, "Tipo asentamiento", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 34, "Origen domicilio", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 42, "Número teléfono", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 43, "Extensión", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 44, "Fax", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 45, "Puesto", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 46, "Fecha contratación", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 47, "Clave moneda", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 48, "Salario mensual", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 49, "Fecha último día de empleo", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 50, "Fecha verificación empleo", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 51, "Origen Razón Social Domicilio", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 59, "Valor activo valuación", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 66, "Fecha cierre cuenta", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 68, "Garantía", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 73, "Número pagos vencidos", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 75, "Histórico pagos", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 76, "Clave prevención", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 77, "Total pagos reportados", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 78, "Clave anterior otorgante", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 79, "Nombre anterior otorgante", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 80, "Número cuenta anterior", self.agregarEstiloAzulClaro(libro))
        hoja.write(1, 90, "Correo electrónico consumidor", self.agregarEstiloAzulClaro(libro))
        
        #Celdas de color azúl fuerte
        hoja.write(1, 3, "Fecha extracción", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 5, "Versión", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 6, "Apelllido paterno", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 7, "Apelllido materno", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 9, "Nombres", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 24, "Dirección", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 25, "Colonia poblacional", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 26, "Delegación/municipio", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 27, "Ciudad", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 28, "Estado", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 29, "CP", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 54, "Cuenta actual", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 55, "Tipo responsabilidad", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 56, "Tipo cuenta", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 57, "Tipo contrato", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 58, "Clave única monetoria", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 60, "Número pagos", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 61, "Frecuencia pagos", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 62, "Monto pagar", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 63, "Fecha apertura cuenta", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 64, "Fecha último pago", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 65, "Fecha última compra", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 67, "Fecha corte", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 69, "Crédito máximo", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 70, "Saldo actual", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 71, "Límite crédito", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 72, "Saldo vencido", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 74, "Pago actual", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 82, "Saldo insoluto", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 88, "Plazo meses", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 89, "Monto crédito originación", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 91, "Total saldos actuales", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 92, "Total saldos vencidos", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 93, "Total elementos nombre reportados", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 94, "Total elementos dirección reportados", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 95, "Total elementos empleo reportados", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 96, "Total elementos cuenta reportados", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 97, "Nombre otorgante", self.agregarEstiloAzulFuerte(libro))
        hoja.write(1, 98, "Domicilio devolución", self.agregarEstiloAzulFuerte(libro))

        #Celdas de color amarillo
        hoja.write(1, 10, "Fecha nacimiento", self.agregarEstiloAmarillo(libro))
        hoja.write(1, 11, "RFC", self.agregarEstiloAmarillo(libro))
        hoja.write(1, 14, "Nacionalidad", self.agregarEstiloAmarillo(libro))
        hoja.write(1, 35, "Nombre empresa", self.agregarEstiloAmarillo(libro))
        hoja.write(1, 36, "Dirección", self.agregarEstiloAmarillo(libro))
        hoja.write(1, 37, "Colonia población", self.agregarEstiloAmarillo(libro))
        hoja.write(1, 38, "Delegación o municipio", self.agregarEstiloAmarillo(libro))
        hoja.write(1, 39, "Ciudad", self.agregarEstiloAmarillo(libro))
        hoja.write(1, 40, "Estado", self.agregarEstiloAmarillo(libro))
        hoja.write(1, 41, "CP", self.agregarEstiloAmarillo(libro))
        hoja.write(1, 52, "Clave actual otorgante", self.agregarEstiloAmarillo(libro))
        hoja.write(1, 53, "Nombre otorgante", self.agregarEstiloAmarillo(libro))
        hoja.write(1, 81, "Fecha primer incumplimiento", self.agregarEstiloAmarillo(libro))
        hoja.write(1, 83, "Monto último pago", self.agregarEstiloAmarillo(libro))
        hoja.write(1, 84, "Fecha ingreso cartera vencida", self.agregarEstiloAmarillo(libro))
        hoja.write(1, 85, "Monto correspondiente intereses", self.agregarEstiloAmarillo(libro))
        hoja.write(1, 86, "Forma pago actual intereses", self.agregarEstiloAmarillo(libro))
        hoja.write(1, 87, "Días vencimiento", self.agregarEstiloAmarillo(libro))

        return hoja
        

