import xlsxwriter
from estilos_documentoPM import ClaseEstilos

class ControllerDocumento:
    # Ejemplo de función para almacenar los datos en las listas correspondientes
    def almacenarDatosEncabezado(self, datos_encabezado, libro, hoja):
        ce = ClaseEstilos()
        valorMaximo = max([len(valor) for valor in datos_encabezado.values()])

        # Recorrer los valores y escribir en las celdas correspondientes
        fila = 2
        for i in range(valorMaximo):
            try:
                otorgante = int(datos_encabezado.get("otorgante", [''])[i])
            except IndexError:
                otorgante = ''

            try:
                otorgante_anterior = int(datos_encabezado.get("otorgante_anterior", [''])[i])
            except IndexError:
                otorgante_anterior = ''
            
            try:
                nombre_otorgante = str(datos_encabezado.get("nombre_otorgante", [''])[i])
            except IndexError:
                nombre_otorgante = ''
                
            try:
                institucion = str(datos_encabezado.get("institucion", [''])[i])
            except IndexError:
                institucion = ''

            try:
                formato = str(datos_encabezado.get("formato", [''])[i])
            except IndexError:
                formato = ''

            try:
                fecha = str(datos_encabezado.get("fecha", [''])[i])
            except IndexError:
                fecha = ''
            
            try:
                periodo_reporta = str(datos_encabezado.get("periodo_reporta", [''])[i])
            except IndexError:
                periodo_reporta = ''

            try:
                version = int(datos_encabezado.get("version", [''])[i])
            except IndexError:
                version = ''
            
            fecha_convertida = ''

            if len(str(datos_encabezado["otorgante"][i])) != 4:
                print("El otorgante en la posición", i + 1, " no tiene la cantidad de digitos necesarios (4).")

            if i < len(datos_encabezado["otorgante_anterior"]) and len(str(datos_encabezado["otorgante_anterior"][i])) > 4:
                print("El otorgante anterior en la posición", i + 1, "no tiene la cantidad de dígitos necesarios (4).")

            if i < len(datos_encabezado["nombre_otorgante"]) and len(str(datos_encabezado["nombre_otorgante"][i])) > 75:
                print("El nombre otorgante en la posición", i + 1, " debe tener maximo 75 caracteres.")

            if i < len(datos_encabezado["institucion"]) and len(str(datos_encabezado["institucion"][i])) != 3:
                print("La institucion en la posición", i + 1, " debe tener 3 caracteres.")

            if i < len(datos_encabezado["formato"]) and len(str(datos_encabezado["formato"][i])) != 1:
                print("El nombre otorgante en la posición", i + 1, " debe tener maximo 75 caracteres.")

            if "fecha" in datos_encabezado and i < len(datos_encabezado["fecha"]):
                fecha = str(datos_encabezado["fecha"][i])
                if len(fecha) == 10:
                    fecha_convertida = fecha[:2] + fecha[3:5] + fecha[6:]
                else:
                    print("Formato de fecha incorrecto.")
            else:
                periodo_reporta = ''
                periodo_reporta_convertido = ''

            if "periodo_reporta" in datos_encabezado and i < len(datos_encabezado["periodo_reporta"]):
                periodo_reporta = str(datos_encabezado["periodo_reporta"][i])
                if len(periodo_reporta) == 10:
                    periodo_reporta_convertido = periodo_reporta[3:5] + periodo_reporta[6:]
                else:
                    print("Formato de período incorrecto.")
            else:
                periodo_reporta = ''
                periodo_reporta_convertido = ''

            if i < len(datos_encabezado["version"]) and len(str(datos_encabezado["version"][i])) != 2:
                print("El nombre otorgante en la posición", i + 1, " debe tener maximo 75 caracteres.") 

            fecha = fecha_convertida
            periodo_reporta = periodo_reporta_convertido

            hoja.write(fila, 0, otorgante, ce.agregarEstiloAmarilloInfo(libro))
            hoja.write(fila, 1, otorgante_anterior, ce.agregarEstiloAzulClaroInfo(libro))
            hoja.write(fila, 2, nombre_otorgante, ce.agregarEstiloAmarilloInfo(libro))
            hoja.write(fila, 3, institucion, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 4, formato, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 5, fecha, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 6, periodo_reporta, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 7, version, ce.agregarEstiloAzulFuerteInfo(libro))
            fila += 1
        return hoja


