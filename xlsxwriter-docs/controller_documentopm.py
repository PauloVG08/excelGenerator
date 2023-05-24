import xlsxwriter
from estilos_documentoPM import ClaseEstilos

class ControllerDocumento:
#------------------------------------Funcion de Encabezado---------------------------------------------------
    def llenarDatosEncabezado(self, datos_encabezado, libro, hoja):
        ce = ClaseEstilos()
        valorMaximo = max([len(valor) for valor in datos_encabezado.values()])

        # Recorrer los valores y escribir en las celdas correspondientes
        fila = 2
        for i in range(valorMaximo):
            try:
                otorgante = int(datos_encabezado.get("otorgante", [''])[i])
            except IndexError:
                otorgante = ' '

            try:
                otorgante_anterior = int(datos_encabezado.get("otorgante_anterior", [''])[i])
            except IndexError:
                otorgante_anterior = ' '
            
            try:
                nombre_otorgante = str(datos_encabezado.get("nombre_otorgante", [''])[i])
            except IndexError:
                nombre_otorgante = ' '
                
            try:
                institucion = str(datos_encabezado.get("institucion", [''])[i])
            except IndexError:
                institucion = ' '

            try:
                formato = str(datos_encabezado.get("formato", [''])[i])
            except IndexError:
                formato = ' '

            try:
                fecha = str(datos_encabezado.get("fecha", [''])[i])
            except IndexError:
                fecha = ' '
            
            try:
                periodo_reporta = str(datos_encabezado.get("periodo_reporta", [''])[i])
            except IndexError:
                datos_encabezado[i] = ' '
                periodo_reporta = datos_encabezado[i]

            try:
                version = int(datos_encabezado.get("version", [''])[i])
            except IndexError:
                datos_encabezado[i] = ' '
                version = datos_encabezado[i]
            
            fecha_convertida = ' '

            # if i < len(datos_encabezado["otorgante"]) and  len(str(otorgante)) < 4:
            #     print("El otorgante en la posición", i + 1, " no tiene la cantidad de digitos necesarios (4).")
            
            if i < len(datos_encabezado["otorgante"]) and len(str(datos_encabezado["otorgante"][i])) > 4:
                otorgante = int(str(datos_encabezado["otorgante"][i])[:4])

            if i < len(datos_encabezado["otorgante_anterior"]) and len(str(otorgante_anterior)) > 4:
                otorgante_anterior = int(str(datos_encabezado["otorgante_anterior"][i])[:4])

            if i < len(datos_encabezado["nombre_otorgante"]) and len(str(nombre_otorgante)) > 75:
                nombre_otorgante = str(nombre_otorgante[:75])

            # if i < len(datos_encabezado["institucion"]) and len(str(institucion)) < 3:
            #     print("La institucion en la posición", i + 1, " debe tener 3 caracteres.")

            if i < len(datos_encabezado["institucion"]) and len(str(institucion)) > 3:
                institucion = str(institucion[:3])

            if i < len(datos_encabezado["formato"]) and len(str(formato)) > 1:
                formato = str(formato[:1])


            if i < len(datos_encabezado["fecha"]) and len(str(fecha)) > 1:
                 fecha = str(datos_encabezado["fecha"][i])
                 fecha_convertida = fecha[:2] + fecha[3:5] + fecha[6:10]
            else:
                fecha = ''
                fecha_convertida = ''


            if i < len(datos_encabezado["periodo_reporta"]) and len(str(periodo_reporta)) > 1:
                 periodo_reporta = str(datos_encabezado["periodo_reporta"][i])
                 periodo_reporta_convertido = periodo_reporta[3:5] + periodo_reporta[6:10]
            else:
                periodo_reporta = ''
                periodo_reporta_convertido = ''


            if i < len(datos_encabezado["version"]) and len(str(version)) > 2:
                version = int(str(datos_encabezado["version"][i])[:2]) 

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
    
#--------------------------Funcion Empresa--------------------------------------------------------------------------
    def llenarDatosEmpresa(self, datos_empresa, libro, hoja):
        ce = ClaseEstilos()
        valorMaximo = max([len(valor) for valor in datos_empresa.values()])

        # Recorrer los valores y escribir en las celdas correspondientes
        fila = 2

        for i in range(valorMaximo):
            try:
                rfc_empresa = str(datos_empresa.get("rfc_empresa", [''])[i])
            except IndexError:
                rfc_empresa = ''
    
            try:
                curp_empresa = str(datos_empresa.get("curp_empresa", [''])[i])
            except (IndexError, TypeError):
                curp_empresa = ' '

            try:
                compania_empresa = str(datos_empresa.get("compania_empresa", [''])[i])
            except IndexError:
                compania_empresa = ''

            try:
                nombre1_empresa = str(datos_empresa.get("nombre1_empresa", [''])[i])
            except IndexError:
                nombre1_empresa = ' '

            try:
                nombre2_empresa = str(datos_empresa.get("nombre2_empresa", [''])[i])
            except IndexError:
                nombre2_empresa = ' '

            try:
                apellido_paterno_empresa = str(datos_empresa.get("apellido_paterno_empresa", [''])[i])
            except IndexError:
                apellido_paterno_empresa = ' '

            try:
                apellido_materno_empresa = str(datos_empresa.get("apellido_materno_empresa", [''])[i])
            except IndexError:
                apellido_materno_empresa = ' '

            try:
                nacionalidad_empresa = str(datos_empresa.get("nacionalidad_empresa", [''])[i])
            except IndexError:
                nacionalidad_empresa = ' '

            try:
                cal_cartera_empresa = str(datos_empresa.get("cal_cartera_empresa", [''])[i])
            except IndexError:
                cal_cartera_empresa = ' '

            try:
                clave_banxico1 = int(datos_empresa.get("clave_banxico1", [''])[i])
            except IndexError:
                clave_banxico1 = ' '

            try:
                clave_banxico2 = int(datos_empresa.get("clave_banxico2", [''])[i])
            except IndexError:
                clave_banxico2 = ' '

            try:
                clave_banxico3 = int(datos_empresa.get("clave_banxico3", [''])[i])
            except IndexError:
                clave_banxico3 = ' '

            try:
                direccion1_empresa = str(datos_empresa.get("direccion1_empresa", [''])[i])
            except IndexError:
                direccion1_empresa = ' '

            try:
                direccion2_empresa = str(datos_empresa.get("direccion2_empresa", [''])[i])
            except IndexError:
                direccion2_empresa = ' ' 

            try:
                colonia_empresa = str(datos_empresa.get("colonia_empresa", [''])[i])
            except IndexError:
                colonia_empresa = ' '

            try:
                deleg_mun_empresa = str(datos_empresa.get("deleg_mun_empresa", [''])[i])
            except IndexError:
                deleg_mun_empresa = ' '

            try:
                ciudad_empresa = str(datos_empresa.get("ciudad_empresa", [''])[i])
            except IndexError:
                ciudad_empresa = ' '

            try:
                estado_empresa = str(datos_empresa.get("estado_empresa", [''])[i])
            except IndexError:
                estado_empresa = ' '

            try:
                cp_empresa = str(datos_empresa.get("cp_empresa", [''])[i])
            except IndexError:
                cp_empresa = ' '

            try:
                telefono_empresa = str(datos_empresa.get("telefono_empresa", [''])[i])
            except IndexError:
                telefono_empresa = ' '

            try:
                extension_empresa = str(datos_empresa.get("extension_empresa", [''])[i])
            except IndexError:
                extension_empresa = ' '

            try:
                fax_empresa = str(datos_empresa.get("fax_empresa", [''])[i])
            except IndexError:
                fax_empresa = ' '

            try:
                tipo_cliente_empresa = int(datos_empresa.get("tipo_cliente_empresa", [''])[i])
            except IndexError:
                tipo_cliente_empresa = ' '

            try:
                edo_extranjero_empresa = str(datos_empresa.get("edo_extranjero_empresa", [''])[i])
            except IndexError:
                edo_extranjero_empresa = ' '

            try:
                pais_empresa = str(datos_empresa.get("pais_empresa", [''])[i])
            except IndexError:
                pais_empresa = ' '

            try:
                tel_movil_empresa = str(datos_empresa.get("tel_movil_empresa", [''])[i])
            except IndexError:
                tel_movil_empresa = ' '

            try:
                correo_empresa = str(datos_empresa.get("correo_empresa", [''])[i])
            except IndexError:
                correo_empresa = ' '

            #Comparación de alertas--------------------------------------------------------------
            # if i < len(datos_empresa["rfc_empresa"]) and len(rfc_empresa) < 12:
            #     print("El RFC empresa en la posición", i + 1, " debe tener 12 o 13 caracteres.")

            if i < len(datos_empresa["rfc_empresa"]) and len(rfc_empresa) > 13:
                rfc_empresa = str(rfc_empresa[:13])
            
            # if i < len(datos_empresa["curp_empresa"]) and len(curp_empresa) < 18:
            #     print("La CURP empresa en la posición", i + 1, " debe tener 18 caracteres.")  

            if i < len(datos_empresa["curp_empresa"]) and len(curp_empresa) > 18:  
                curp_empresa = str(curp_empresa[:18])

            if i < len(datos_empresa["compania_empresa"]) and len(compania_empresa) > 150:
                compania_empresa = str(compania_empresa[:150])

            if i < len(datos_empresa["nombre1_empresa"]) and len(nombre1_empresa) > 30:
                nombre1_empresa = str(nombre1_empresa[:30])

            if i < len(datos_empresa["nombre2_empresa"]) and len(nombre2_empresa) > 30:
                nombre2_empresa = str(nombre2_empresa[:30])

            # if i < len(datos_empresa["apellido_paterno_empresa"]) and (len(apellido_paterno_empresa) < 2):
            #     print("El apellido paterno empresa debe tener minimo 2 caracteres y máximo 25")

            if i < len(datos_empresa["apellido_paterno_empresa"]) and len(apellido_paterno_empresa) > 25:
                apellido_paterno_empresa = str(apellido_paterno_empresa[:25])

            # if i < len(datos_empresa["apellido_materno_empresa"]) and (len(apellido_materno_empresa) < 2):
            #     print("El apellido materno empresa debe tener minimo 2 caracteres y máximo 25")

            if i < len(datos_empresa["apellido_materno_empresa"]) and len(apellido_materno_empresa) > 25:
                apellido_materno_empresa = str(apellido_materno_empresa[:25])

            if i < len(datos_empresa["nacionalidad_empresa"]) and len(nacionalidad_empresa) > 2:
                nacionalidad_empresa = str(nacionalidad_empresa[:2])

            if i < len(datos_empresa["cal_cartera_empresa"]) and len(cal_cartera_empresa) > 2:
                cal_cartera_empresa = str(cal_cartera_empresa[:2])

            if i < len(datos_empresa["clave_banxico1"]) and len(str(clave_banxico1)) > 11:
                clave_banxico1 = int(str(datos_empresa["clave_banxico1"][i])[:11])

            if i < len(datos_empresa["clave_banxico2"]) and len(str(clave_banxico2)) > 11:
                clave_banxico2 = int(str(datos_empresa["clave_banxico2"][i])[:11])

            if i < len(datos_empresa["clave_banxico3"]) and len(str(clave_banxico3)) > 11:
                clave_banxico3 = int(str(clave_banxico3)[:11])

            if i < len(datos_empresa["direccion1_empresa"]) and len(str(datos_empresa["direccion1_empresa"][i])) > 40:
                direccion1_empresa = str(direccion1_empresa[:40])
                
            if i < len(datos_empresa["colonia_empresa"]) and len(str(datos_empresa["colonia_empresa"][i])) > 60:
                colonia_empresa = str(colonia_empresa[:60])

            if i < len(datos_empresa["deleg_mun_empresa"]) and (len(deleg_mun_empresa) > 40):
                deleg_mun_empresa = str(deleg_mun_empresa[:40])

            if i < len(datos_empresa["ciudad_empresa"]) and (len(ciudad_empresa) > 40):
                ciudad_empresa = str(ciudad_empresa[:40])

            if i < len(datos_empresa["estado_empresa"]) and (len(estado_empresa) > 4):
                estado_empresa = str(estado_empresa[:4])

            if i < len(datos_empresa["cp_empresa"]) and (len(cp_empresa) > 10):
                cp_empresa = str(cp_empresa[:10])

            if i < len(datos_empresa["telefono_empresa"]) and (len(telefono_empresa) > 11):
                telefono_empresa = str(telefono_empresa[:11])

            if i < len(datos_empresa["extension_empresa"]) and (len(extension_empresa) > 8):
                extension_empresa = str(extension_empresa[:8])

            if i < len(datos_empresa["fax_empresa"]) and (len(fax_empresa) > 11):
                fax_empresa = str(fax_empresa[:11])

            if i < len(datos_empresa["tipo_cliente_empresa"]) and (len(str(tipo_cliente_empresa)) > 1):
                tipo_cliente_empresa = int(str(datos_empresa["tipo_cliente_empresa"][i])[:1])

            if i < len(datos_empresa["edo_extranjero_empresa"]) and (len(edo_extranjero_empresa) > 40):
                edo_extranjero_empresa = str(edo_extranjero_empresa[:40])

            if i < len(datos_empresa["pais_empresa"]) and (len(pais_empresa) > 2):
                pais_empresa = str(pais_empresa[:2])

            if i < len(datos_empresa["tel_movil_empresa"]) and (len(tel_movil_empresa) > 10):
                tel_movil_empresa = str(tel_movil_empresa[:10])

            if i < len(datos_empresa["correo_empresa"]) and (len(correo_empresa) > 100):
                correo_empresa = str(correo_empresa[:100])

            #Escribir los datos en la hoja-----------------------------
            hoja.write(fila, 8, rfc_empresa, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 9, curp_empresa, ce.agregarEstiloAzulClaroInfo(libro))
            hoja.write(fila, 10, compania_empresa, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 11, nombre1_empresa, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 12, nombre2_empresa, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 13, apellido_paterno_empresa, ce.agregarEstiloAzulClaroInfo(libro))
            hoja.write(fila, 14, apellido_materno_empresa, ce.agregarEstiloAzulClaroInfo(libro))
            hoja.write(fila, 15, nacionalidad_empresa, ce.agregarEstiloVerdeInfo(libro))
            hoja.write(fila, 16, cal_cartera_empresa, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 17, clave_banxico1, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 18, clave_banxico2, ce.agregarEstiloAzulClaroInfo(libro))
            hoja.write(fila, 19, clave_banxico3, ce.agregarEstiloAzulClaroInfo(libro))
            hoja.write(fila, 20, direccion1_empresa, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 21, direccion2_empresa, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 22, colonia_empresa, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 23, deleg_mun_empresa, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 24, ciudad_empresa, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 25, estado_empresa, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 26, cp_empresa, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 27, telefono_empresa, ce.agregarEstiloAzulClaroInfo(libro))
            hoja.write(fila, 28, extension_empresa, ce.agregarEstiloAzulClaroInfo(libro))
            hoja.write(fila, 29, fax_empresa, ce.agregarEstiloAzulClaroInfo(libro))
            hoja.write(fila, 30, tipo_cliente_empresa, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 31, edo_extranjero_empresa, ce.agregarEstiloAzulClaroInfo(libro))
            hoja.write(fila, 32, pais_empresa, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 33, tel_movil_empresa, ce.agregarEstiloGrisInfo(libro))
            hoja.write(fila, 34, correo_empresa, ce.agregarEstiloGrisInfo(libro))
            fila += 1
        return hoja

#--------------------------Funcion Accionista-------------------------------------------------------------------------
    def llenarDatosAccionista(self, datos_accionista, libro, hoja):
        ce = ClaseEstilos()
        valorMaximo = max([len(valor) for valor in datos_accionista.values()])

         # Recorrer los valores y escribir en las celdas correspondientes
        fila = 2

        for i in range(valorMaximo):
            try:
                rfc_accionista = str(datos_accionista.get("rfc_accionista", [''])[i])
            except IndexError:
                rfc_accionista = ''
    
            try:
                curp_accionista = str(datos_accionista.get("curp_accionista", [''])[i])
            except (IndexError, TypeError):
                curp_accionista = ' '

            try:
                compania_accionista = str(datos_accionista.get("compania_accionista", [''])[i])
            except IndexError:
                compania_accionista = ''

            try:
                nombre1_accionista = str(datos_accionista.get("nombre1_accionista", [''])[i])
            except IndexError:
                nombre1_accionista = ' '

            try:
                nombre2_accionista = str(datos_accionista.get("nombre2_accionista", [''])[i])
            except IndexError:
                nombre2_accionista = ' '

            try:
                apellido_paterno_accionista = str(datos_accionista.get("apellido_paterno_accionista", [''])[i])
            except IndexError:
                apellido_paterno_accionista = ' '

            try:
                apellido_materno_accionista = str(datos_accionista.get("apellido_materno_accionista", [''])[i])
            except IndexError:
                apellido_materno_accionista = ' '

            try:
                porcentaje_accionista = int(datos_accionista.get("porcentaje_accionista", [''])[i])
            except IndexError:
                porcentaje_accionista = ' '

            try:
                direccion1_accionista = str(datos_accionista.get("direccion1_accionista", [''])[i])
            except IndexError:
                direccion1_accionista = ' '

            try:
                direccion2_accionista = str(datos_accionista.get("direccion2_accionista", [''])[i])
            except IndexError:
                direccion2_accionista = ' ' 

            try:
                colonia_accionista = str(datos_accionista.get("colonia_accionista", [''])[i])
            except IndexError:
                colonia_accionista = ' '

            try:
                deleg_mun_accionista = str(datos_accionista.get("deleg_mun_accionista", [''])[i])
            except IndexError:
                deleg_mun_accionista = ' '

            try:
                ciudad_accionista = str(datos_accionista.get("ciudad_accionista", [''])[i])
            except IndexError:
                ciudad_accionista = ' '

            try:
                estado_accionista = str(datos_accionista.get("estado_accionista", [''])[i])
            except IndexError:
                estado_accionista = ' '

            try:
                cp_accionista = str(datos_accionista.get("cp_accionista", [''])[i])
            except IndexError:
                cp_accionista = ' '

            try:
                telefono_accionista = str(datos_accionista.get("telefono_accionista", [''])[i])
            except IndexError:
                telefono_accionista = ' '

            try:
                extension_accionista = str(datos_accionista.get("extension_accionista", [''])[i])
            except IndexError:
                extension_accionista = ' '

            try:
                fax_accionista = str(datos_accionista.get("fax_accionista", [''])[i])
            except IndexError:
                fax_accionista= ' '

            try:
                tipo_cliente_accionista = int(datos_accionista.get("tipo_cliente_accionista", [''])[i])
            except IndexError:
                tipo_cliente_accionista = ' '

            try:
                edo_extranjero_accionista = str(datos_accionista.get("edo_extranjero_accionista", [''])[i])
            except IndexError:
                edo_extranjero_accionista = ' '

            try:
                pais_accionista = str(datos_accionista.get("pais_accionista", [''])[i])
            except IndexError:
                pais_accionista = ' '

            #Comparación de datos---------------------------------------------------------------------------
            # if i < len(datos_accionista["rfc_accionista"]) and len(rfc_accionista) < 12:
            #     print("El RFC en la posición", i + 1, " debe tener 12 o 13 caracteres.")

            if i < len(datos_accionista["rfc_accionista"]) and len(rfc_accionista) > 13:
                rfc_accionista = str(rfc_accionista[:13])
            
            if i < len(datos_accionista["curp_accionista"]) and len(curp_accionista) > 18:  
                curp_accionista = str(curp_accionista[:18])    

            if i < len(datos_accionista["compania_accionista"]) and len(compania_accionista) > 150:
                compania_accionista = str(compania_accionista[:150])

            if i < len(datos_accionista["nombre1_accionista"]) and len(nombre1_accionista) > 30:
                nombre1_accionista = str(nombre1_accionista[:30])

            if i < len(datos_accionista["nombre2_accionista"]) and len(nombre2_accionista) > 30:
                nombre2_accionista = str(nombre2_accionista[:30])

            if i < len(datos_accionista["apellido_paterno_accionista"]) and (len(apellido_paterno_accionista) > 25):
                apellido_paterno_accionista = str(apellido_paterno_accionista[:25])

            if i < len(datos_accionista["apellido_materno_accionista"]) and (len(apellido_materno_accionista) > 25):
                apellido_materno_accionista = str(apellido_materno_accionista[:25])

            if i < len(datos_accionista["porcentaje_accionista"]) and len(str(porcentaje_accionista)) > 2:
                porcentaje_accionista = int(str(datos_accionista["porcentaje_accionista"][i])[:2])

            if i < len(datos_accionista["direccion1_accionista"]) and len(str(direccion1_accionista)) > 40:
                direccion1_accionista = direccion1_accionista[:40]
                
            if i < len(datos_accionista["colonia_accionista"]) and len(str(colonia_accionista)) > 60:
                colonia_accionista = str(colonia_accionista[:60])

            if i < len(datos_accionista["deleg_mun_accionista"]) and (len(deleg_mun_accionista) > 40):
                deleg_mun_accionista = str(deleg_mun_accionista[:40])

            if i < len(datos_accionista["ciudad_accionista"]) and (len(ciudad_accionista) > 40):
                ciudad_accionista = str(ciudad_accionista[:40])

            # if i < len(datos_accionista["estado_accionista"]) and (len(estado_accionista) != 4):
            #     print("El estado debe tener máximo 4 caracteres")

            if i < len(datos_accionista["estado_accionista"]) and (len(estado_accionista) > 4):
                estado_accionista = str(estado_accionista[:4])

            if i < len(datos_accionista["cp_accionista"]) and (len(cp_accionista) > 10):
                 cp_accionista = str(cp_accionista[:10])

            if i < len(datos_accionista["telefono_accionista"]) and (len(telefono_accionista) > 11):
                telefono_accionista = str(telefono_accionista[:11])

            if i < len(datos_accionista["extension_accionista"]) and (len(extension_accionista) > 8):
                extension_accionista = str(extension_accionista[:8])

            if i < len(datos_accionista["fax_accionista"]) and (len(fax_accionista) > 11):
                fax_accionista = str(fax_accionista[:11])

            if i < len(datos_accionista["tipo_cliente_accionista"]) and (len(str(tipo_cliente_accionista)) > 1):
                tipo_cliente_accionista = int(str(datos_accionista["tipo_cliente_accionista"][i])[:1])

            if i < len(datos_accionista["edo_extranjero_accionista"]) and (len(edo_extranjero_accionista) > 40):
                edo_extranjero_accionista = str(edo_extranjero_accionista[:40])

            if i < len(datos_accionista["pais_accionista"]) and (len(pais_accionista) > 2):
                pais_accionista = str(pais_accionista[:2])

            #Escribir los datos en la hoja-----------------------------
            hoja.write(fila, 35, rfc_accionista, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 36, curp_accionista, ce.agregarEstiloAzulClaroInfo(libro))
            hoja.write(fila, 37, compania_accionista, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 38, nombre1_accionista, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 39, nombre2_accionista, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 40, apellido_paterno_accionista, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 41, apellido_materno_accionista, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 42, porcentaje_accionista, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 43, direccion1_accionista, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 44, direccion2_accionista, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 45, colonia_accionista, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 46, deleg_mun_accionista, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 47, ciudad_accionista, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 48, estado_accionista, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 49, cp_accionista, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 50, telefono_accionista, ce.agregarEstiloAzulClaroInfo(libro))
            hoja.write(fila, 51, extension_accionista, ce.agregarEstiloAzulClaroInfo(libro))
            hoja.write(fila, 52, fax_accionista, ce.agregarEstiloAzulClaroInfo(libro))
            hoja.write(fila, 53, tipo_cliente_accionista, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 54, edo_extranjero_accionista, ce.agregarEstiloAzulClaroInfo(libro))
            hoja.write(fila, 55, pais_accionista, ce.agregarEstiloAzulClaroInfo(libro))
            fila += 1
        return hoja

#--------------------------Funcion Credito--------------------------------------------------------------------------
    def llenarDatosCredito(self, datos_credito, libro, hoja):
        ce = ClaseEstilos()
        valorMaximo = max([len(valor) for valor in datos_credito.values()])

         # Recorrer los valores y escribir en las celdas correspondientes
        fila = 2

        for i in range(valorMaximo):
            try:
                rfc_credito = str(datos_credito.get("rfc_credito", [''])[i])
            except IndexError:
                rfc_credito = ''

            try:
                experiencias_crediticias = int(datos_credito.get("experiencias_crediticias", [''])[i])
            except IndexError:
                experiencias_crediticias = ''
                
            try:
                num_contrato = str(datos_credito.get("num_contrato", [''])[i])
            except IndexError:
                num_contrato = ''
                
            try:
                num_contrato_anterior = str(datos_credito.get("num_contrato_anterior", [''])[i])
            except IndexError:
                num_contrato_anterior = ''

            try:
                fecha_apertura = str(datos_credito.get("fecha_apertura", [''])[i])
            except IndexError:
                fecha_apertura = ''

            try:
                plazo_meses = int(datos_credito.get("plazo_meses", [''])[i])
            except IndexError:
                plazo_meses = ''

            try:
                tipo_credito = int(datos_credito.get("tipo_credito", [''])[i])
            except IndexError:
                tipo_credito = ''

            try:
                saldo_inicial = int(datos_credito.get("saldo_inicial", [''])[i])
            except IndexError:
                saldo_inicial = ''

            try:
                moneda = str(datos_credito.get("moneda", [''])[i])
            except IndexError:
                moneda = ''

            try:
                num_pagos = int(datos_credito.get("num_pagos", [''])[i])
            except IndexError:
                num_pagos = ''

            try:
                frecuencia_pagos = int(datos_credito.get("frecuencia_pagos", [''])[i])
            except IndexError:
                frecuencia_pagos = ''

            try:
                importe_pagos = int(datos_credito.get("importe_pagos", [''])[i])
            except IndexError:
                importe_pagos = ''

            try:
                fecha_ultimo_pago = str(datos_credito.get("fecha_ultimo_pago", [''])[i])
            except IndexError:
                fecha_ultimo_pago = ''

            try:
                fecha_reestructura = str(datos_credito.get("fecha_reestructura", [''])[i])
            except IndexError:
                fecha_reestructura = ''

            try:
                pago_efectivo = int(datos_credito.get("pago_efectivo", [''])[i])
            except IndexError:
                pago_efectivo = ''

            try:
                fecha_liquidacion = str(datos_credito.get("fecha_liquidacion", [''])[i])
            except IndexError:
                fecha_liquidacion = ''

            try:
                quita = int(datos_credito.get("quita", [''])[i])
            except IndexError:
                quita = ''

            try:
                dacion = int(datos_credito.get("dacion", [''])[i])
            except IndexError:
                dacion = ''

            try:
                quebranto_castigo = int(datos_credito.get("quebranto_castigo", [''])[i])
            except IndexError:
                quebranto_castigo = ''

            try:
                clave_observacion = str(datos_credito.get("clave_observacion", [''])[i])
            except IndexError:
                clave_observacion = ''

            try:
                especiales = str(datos_credito.get("especiales", [''])[i])
            except IndexError:
                especiales = ''

            try:
                fecha_primer_cumplimiento = str(datos_credito.get("fecha_primer_cumplimiento", [''])[i])
            except IndexError:
                fecha_primer_cumplimiento = ''

            try:
                saldo_insoluto = int(datos_credito.get("saldo_insoluto", [''])[i])
            except IndexError:
                saldo_insoluto = ''

            try:
                credito_maximo_utilizado = int(datos_credito.get("credito_maximo_utilizado", [''])[i])
            except IndexError:
                credito_maximo_utilizado = ''

            try:
                fecha_ingreso_cv = str(datos_credito.get("fecha_ingreso_cv", [''])[i])
            except IndexError:
                fecha_ingreso_cv = ''

            #Comparación de datos---------------------------------------------------------------------------
            if i < len(datos_credito["rfc_credito"]) and len(rfc_credito) > 13:
                rfc_credito = str(rfc_credito[:13])

            if i < len(datos_credito["experiencias_crediticias"]) and len(str(experiencias_crediticias)) > 6:
                 experiencias_crediticias = int(str(datos_credito["experiencias_crediticias"][i])[:6])

            if i < len(datos_credito["num_contrato"]) and len(str(num_contrato)) > 25:
                num_contrato = str(num_contrato[:25])

            if i < len(datos_credito["num_contrato_anterior"]) and len(str(num_contrato_anterior)) > 25:
                num_contrato_anterior = str(num_contrato_anterior[:25])


            if i < len(datos_credito["fecha_apertura"]) and len(str(fecha_apertura)) > 1:
                fecha_apertura = str(datos_credito["fecha_apertura"][i])
                fecha_apertura_convertida = int(fecha_apertura[:2] + fecha_apertura[3:5] + fecha_apertura[6:10])
            else:
                fecha_apertura = ''
                fecha_apertura_convertida = ''
            fecha_apertura = fecha_apertura_convertida


            if i < len(datos_credito["plazo_meses"]) and len(str(plazo_meses)) > 5:
                plazo_meses = int(str(datos_credito["plazo_meses"][i])[:5])

            if i < len(datos_credito["tipo_credito"]) and len(str(tipo_credito)) > 4:
                tipo_credito = int(str(datos_credito["tipo_credito"][i])[:4])

            if i < len(datos_credito["saldo_inicial"]) and len(str(saldo_inicial)) > 20:
                 saldo_inicial = int(str(saldo_inicial["tipo_credito"][i])[:20])

            if i < len(datos_credito["moneda"]) and len(str(moneda)) > 3:
                moneda = int(str(moneda["moneda"][i])[:3])

            if i < len(datos_credito["num_pagos"]) and len(str(num_pagos)) > 4:
                num_pagos = int(str(num_pagos["num_pagos"][i])[:4])

            if i < len(datos_credito["frecuencia_pagos"]) and len(str(frecuencia_pagos)) > 5:
                frecuencia_pagos = int(str(frecuencia_pagos["frecuencia_pagos"][i])[:5])

            if i < len(datos_credito["importe_pagos"]) and len(str(importe_pagos)) > 20:
                importe_pagos = int(str(importe_pagos["importe_pagos"][i])[:20])


            if i < len(datos_credito["fecha_ultimo_pago"]) and len(str(fecha_ultimo_pago)) > 1:
                fecha_ultimo_pago = str(datos_credito["fecha_ultimo_pago"][i])
                fecha_ultimo_pago_convertida = int(fecha_ultimo_pago[:2] + fecha_ultimo_pago[3:5] + fecha_ultimo_pago[6:10])
            else:
                fecha_ultimo_pago = ''
                fecha_ultimo_pago_convertida = ''
            fecha_ultimo_pago = fecha_ultimo_pago_convertida


            if i < len(datos_credito["fecha_reestructura"]) and len(str(fecha_reestructura)) > 1:
                fecha_reestructura = str(datos_credito["fecha_reestructura"][i])
                fecha_reestructura_convertida = int(fecha_reestructura[:2] + fecha_reestructura[3:5] + fecha_reestructura[6:10])
            else:
                fecha_reestructura = ''
                fecha_reestructura_convertida = ''
            fecha_reestructura = fecha_reestructura_convertida


            if i < len(datos_credito["pago_efectivo"]) and len(str(pago_efectivo)) > 20:
                pago_efectivo = int(str(pago_efectivo["pago_efectivo"][i])[:20])


            if i < len(datos_credito["fecha_liquidacion"]) and len(str(fecha_liquidacion)) > 1:
                fecha_liquidacion = str(datos_credito["fecha_liquidacion"][i])
                fecha_liquidacion_convertida = int(fecha_liquidacion[:2] + fecha_liquidacion[3:5] + fecha_liquidacion[6:10])
            else:
                fecha_liquidacion = ''
                fecha_liquidacion_convertida = ''
            fecha_liquidacion = fecha_liquidacion_convertida


            if i < len(datos_credito["quita"]) and len(str(quita)) > 20:
                 quita = int(str(quita["quita"][i])[:20])

            if i < len(datos_credito["dacion"]) and len(str(dacion)) > 20:
                dacion = int(str(dacion["dacion"][i])[:20])

            if i < len(datos_credito["quebranto_castigo"]) and len(str(quebranto_castigo)) > 20:
                quebranto_castigo = int(str(quebranto_castigo["quebranto_castigo"][i])[:20])

            if i < len(datos_credito["clave_observacion"]) and len(str(clave_observacion)) > 4:
                clave_observacion = str(clave_observacion[:4])

            if i < len(datos_credito["especiales"]) and len(str(especiales)) > 1:
                especiales = str(especiales[:1])


            if i < len(datos_credito["fecha_primer_cumplimiento"]) and len(str(fecha_primer_cumplimiento)) > 1:
                fecha_primer_cumplimiento = str(datos_credito["fecha_primer_cumplimiento"][i])
                fecha_primer_cumplimiento_convertida = int(fecha_primer_cumplimiento[:2] + fecha_primer_cumplimiento[3:5] + fecha_primer_cumplimiento[6:10])
            else:
                fecha_primer_cumplimiento = ''
                fecha_primer_cumplimiento_convertida = ''
            fecha_primer_cumplimiento = fecha_primer_cumplimiento_convertida


            if i < len(datos_credito["saldo_insoluto"]) and len(str(saldo_insoluto)) > 8:
                saldo_insoluto = int(str(saldo_insoluto["saldo_insoluto"][i])[:8])

            if i < len(datos_credito["credito_maximo_utilizado"]) and len(str(credito_maximo_utilizado)) > 20:
                credito_maximo_utilizado = int(str(credito_maximo_utilizado["credito_maximo_utilizado"][i])[:20])


            if i < len(datos_credito["fecha_ingreso_cv"]) and len(str(fecha_ingreso_cv)) > 1:
                fecha_ingreso_cv = str(datos_credito["fecha_ingreso_cv"][i])
                fecha_ingreso_cv_convertida = int(fecha_ingreso_cv[:2] + fecha_ingreso_cv[3:5] + fecha_ingreso_cv[6:10])
            else:
                fecha_ingreso_cv = ''
                fecha_ingreso_cv_convertida = ''
            fecha_ingreso_cv = fecha_ingreso_cv_convertida


            #Agregar los datos a las celdas de excel
            hoja.write(fila, 56, rfc_credito, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 57, experiencias_crediticias, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 58, num_contrato, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 59, num_contrato_anterior, ce.agregarEstiloAzulClaroInfo(libro))
            hoja.write(fila, 60, fecha_apertura, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 61, plazo_meses, ce.agregarEstiloVerdeInfo(libro))
            hoja.write(fila, 62, tipo_credito, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 63, saldo_inicial, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 64, moneda, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 65, num_pagos, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 66, frecuencia_pagos, ce.agregarEstiloVerdeInfo(libro))
            hoja.write(fila, 67, importe_pagos, ce.agregarEstiloVerdeInfo(libro))
            hoja.write(fila, 68, fecha_ultimo_pago, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 69, fecha_reestructura, ce.agregarEstiloAzulClaroInfo(libro))
            hoja.write(fila, 70, pago_efectivo, ce.agregarEstiloAzulClaroInfo(libro))
            hoja.write(fila, 71, fecha_liquidacion, ce.agregarEstiloAzulClaroInfo(libro))
            hoja.write(fila, 72, quita, ce.agregarEstiloAzulClaroInfo(libro))
            hoja.write(fila, 73, dacion, ce.agregarEstiloAzulClaroInfo(libro))
            hoja.write(fila, 74, quebranto_castigo, ce.agregarEstiloAzulClaroInfo(libro))
            hoja.write(fila, 75, clave_observacion, ce.agregarEstiloAzulClaroInfo(libro))
            hoja.write(fila, 76, especiales, ce.agregarEstiloAzulClaroInfo(libro))
            hoja.write(fila, 77, fecha_primer_cumplimiento, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 78, saldo_insoluto, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 79, credito_maximo_utilizado, ce.agregarEstiloVerdeInfo(libro))
            hoja.write(fila, 80, fecha_ingreso_cv, ce.agregarEstiloGrisInfo(libro))
            fila += 1
        return hoja

#--------------------------Funcion Detalle de crédito----------------------------------------------------------------
    def llenarDatosDetalleCredito(self, datos_detalle_credito, libro, hoja):
        ce = ClaseEstilos()
        valorMaximo = max([len(valor) for valor in datos_detalle_credito.values()])

        # Recorrer los valores y escribir en las celdas correspondientes
        fila = 2

        for i in range(valorMaximo):
            try:
                rfc_detalle_credito = str(datos_detalle_credito.get("rfc_detalle_credito", [''])[i])
            except IndexError:
                rfc_detalle_credito = ' '
                
            try:
                num_contrato_detalle_credito = str(datos_detalle_credito.get("num_contrato_detalle_credito", [''])[i])
            except IndexError:
                num_contrato_detalle_credito = ' '
                
            try:
                dias_vencidos_detalle_credito = int(datos_detalle_credito.get("dias_vencidos_detalle_credito", [''])[i])
            except IndexError:
                dias_vencidos_detalle_credito = ' '
         
            try:
                cantidad_detalle_credito = int(datos_detalle_credito.get("cantidad_detalle_credito", [''])[i])
            except IndexError:
                cantidad_detalle_credito = ' '
         
            try:
                intereses_detalle_credito = int(datos_detalle_credito.get("intereses_detalle_credito", [''])[i])
            except IndexError:
                intereses_detalle_credito = ' '

            #Comparación de valores----------------------------------------------------------------------------------------------

            if i < len(datos_detalle_credito["rfc_detalle_credito"]) and len(str(rfc_detalle_credito)) > 13:
                rfc_detalle_credito = str(rfc_detalle_credito[:13])

            if i < len(datos_detalle_credito["num_contrato_detalle_credito"]) and len(str(num_contrato_detalle_credito)) > 25:
                num_contrato_detalle_credito = str(num_contrato_detalle_credito[:25])

            if i < len(datos_detalle_credito["dias_vencidos_detalle_credito"]) and len(str(dias_vencidos_detalle_credito)) > 3:
                dias_vencidos_detalle_credito = int(str(datos_detalle_credito["dias_vencidos_detalle_credito"][i])[:3])

            if i < len(datos_detalle_credito["cantidad_detalle_credito"]) and len(str(cantidad_detalle_credito)) > 20:
                cantidad_detalle_credito = int(str(datos_detalle_credito["cantidad_detalle_credito"][i])[:20])

            if i < len(datos_detalle_credito["intereses_detalle_credito"]) and len(str(intereses_detalle_credito)) > 10:
                intereses_detalle_credito = int(str(datos_detalle_credito["intereses_detalle_credito"][i])[:10])

            hoja.write(fila, 81, rfc_detalle_credito, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 82, num_contrato_detalle_credito, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 83, dias_vencidos_detalle_credito, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 84, cantidad_detalle_credito, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 85, intereses_detalle_credito, ce.agregarEstiloAzulFuerteInfo(libro))
            fila += 1
        return hoja

    def llenarDatosAval(self, datos_aval, libro, hoja):
        ce = ClaseEstilos()
        valorMaximo = max([len(valor) for valor in datos_aval.values()])

         # Recorrer los valores y escribir en las celdas correspondientes
        fila = 2

        for i in range(valorMaximo):
            try:
                rfc_aval = str(datos_aval.get("rfc_aval", [''])[i])
            except IndexError:
                rfc_aval = ''
    
            try:
                curp_aval = str(datos_aval.get("curp_aval", [''])[i])
            except (IndexError, TypeError):
                curp_aval = ' '

            try:
                compania_aval = str(datos_aval.get("compania_aval", [''])[i])
            except IndexError:
                compania_aval = ''

            try:
                nombre1_aval = str(datos_aval.get("nombre1_aval", [''])[i])
            except IndexError:
                nombre1_aval = ' '

            try:
                nombre2_aval = str(datos_aval.get("nombre2_aval", [''])[i])
            except IndexError:
                nombre2_aval = ' '

            try:
                apellido_paterno_aval = str(datos_aval.get("apellido_paterno_aval", [''])[i])
            except IndexError:
                apellido_paterno_aval = ' '

            try:
                apellido_materno_aval = str(datos_aval.get("apellido_materno_aval", [''])[i])
            except IndexError:
                apellido_materno_aval = ' '

            try:
                direccion1_aval = str(datos_aval.get("direccion1_aval", [''])[i])
            except IndexError:
                direccion1_aval = ' '

            try:
                direccion2_aval = str(datos_aval.get("direccion2_aval", [''])[i])
            except IndexError:
                direccion2_aval = ' ' 

            try:
                colonia_aval = str(datos_aval.get("colonia_aval", [''])[i])
            except IndexError:
                colonia_aval = ' '

            try:
                deleg_mun_aval = str(datos_aval.get("deleg_mun_aval", [''])[i])
            except IndexError:
                deleg_mun_aval = ' '

            try:
                ciudad_aval = str(datos_aval.get("ciudad_aval", [''])[i])
            except IndexError:
                ciudad_aval = ' '

            try:
                estado_aval = str(datos_aval.get("estado_aval", [''])[i])
            except IndexError:
                estado_aval = ' '

            try:
                cp_aval = str(datos_aval.get("cp_aval", [''])[i])
            except IndexError:
                cp_aval = ' '

            try:
                telefono_aval = str(datos_aval.get("telefono_aval", [''])[i])
            except IndexError:
                telefono_aval = ' '

            try:
                extension_aval = str(datos_aval.get("extension_aval", [''])[i])
            except IndexError:
                extension_aval = ' '

            try:
                fax_aval = str(datos_aval.get("fax_aval", [''])[i])
            except IndexError:
                fax_aval = ' '

            try:
                tipo_cliente_aval = int(datos_aval.get("tipo_cliente_aval", [''])[i])
            except IndexError:
                tipo_cliente_aval = ' '

            try:
                edo_extranjero_aval = str(datos_aval.get("edo_extranjero_aval", [''])[i])
            except IndexError:
                edo_extranjero_aval = ' '

            try:
                pais_aval = str(datos_aval.get("pais_aval", [''])[i])
            except IndexError:
                pais_aval = ' '

            if i < len(datos_aval["rfc_aval"]) and len(rfc_aval) > 13:
                rfc_aval = str(rfc_aval[:13])
            
            if i < len(datos_aval["curp_aval"]) and len(curp_aval) > 18:
                curp_aval = str(curp_aval[:18])    

            if i < len(datos_aval["compania_aval"]) and len(compania_aval) > 150:
                compania_aval = str(compania_aval[:150])

            if i < len(datos_aval["nombre1_aval"]) and len(nombre1_aval) > 30:
                nombre1_aval = str(nombre1_aval[:30])

            if i < len(datos_aval["nombre2_aval"]) and len(nombre2_aval) > 30:
                nombre2_aval = str(nombre2_aval[:30])

            if i < len(datos_aval["apellido_paterno_aval"]) and (len(apellido_paterno_aval) > 25):
                apellido_paterno_aval = str(apellido_paterno_aval[:25])

            if i < len(datos_aval["apellido_materno_aval"]) and (len(apellido_materno_aval) > 25):
                apellido_materno_aval = str(apellido_materno_aval[:25])

            if i < len(datos_aval["direccion1_aval"]) and len(str(datos_aval["direccion1_aval"][i])) > 40:
                direccion1_aval = direccion1_aval[:40]
                
            if i < len(datos_aval["colonia_aval"]) and len(str(datos_aval["colonia_aval"][i])) > 60:
                colonia_aval = colonia_aval[:60]

            if i < len(datos_aval["deleg_mun_aval"]) and (len(deleg_mun_aval) > 40):
                deleg_mun_aval = deleg_mun_aval[:40]

            if i < len(datos_aval["ciudad_aval"]) and (len(ciudad_aval) > 40):
                ciudad_aval = ciudad_aval[:40]

            if i < len(datos_aval["estado_aval"]) and (len(estado_aval) > 4):
                estado_aval = estado_aval[:4]

            if i < len(datos_aval["cp_aval"]) and (len(cp_aval) > 10):
                cp_aval = cp_aval[:10]

            if i < len(datos_aval["telefono_aval"]) and (len(telefono_aval) > 11):
                telefono_aval = telefono_aval[:11]

            if i < len(datos_aval["extension_aval"]) and (len(extension_aval) > 8):
                extension_aval = extension_aval[:8]

            if i < len(datos_aval["fax_aval"]) and (len(fax_aval) > 11):
                fax_aval = fax_aval[:11]

            if i < len(datos_aval["tipo_cliente_aval"]) and (len(str(tipo_cliente_aval)) > 1):
                tipo_cliente_aval = int(str(datos_aval["tipo_cliente_aval"][i])[:1])

            if i < len(datos_aval["edo_extranjero_aval"]) and (len(edo_extranjero_aval) > 40):
                edo_extranjero_aval = edo_extranjero_aval[:40]

            if i < len(datos_aval["pais_aval"]) and (len(pais_aval) > 2):
                pais_aval = pais_aval[:2]

            hoja.write(fila, 86, rfc_aval, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 87, curp_aval, ce.agregarEstiloAzulClaroInfo(libro))
            hoja.write(fila, 88, compania_aval, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 89, nombre1_aval, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 90, nombre2_aval, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 91, apellido_paterno_aval, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 92, apellido_materno_aval, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 93, direccion1_aval, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 94, direccion2_aval, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 95, colonia_aval, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 96, deleg_mun_aval, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 97, ciudad_aval, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 98, estado_aval, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 99, cp_aval, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 100, telefono_aval, ce.agregarEstiloAzulClaroInfo(libro))
            hoja.write(fila, 101, extension_aval, ce.agregarEstiloAzulClaroInfo(libro))
            hoja.write(fila, 102, fax_aval, ce.agregarEstiloAzulClaroInfo(libro))
            hoja.write(fila, 103, tipo_cliente_aval, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 104, edo_extranjero_aval, ce.agregarEstiloAzulClaroInfo(libro))
            hoja.write(fila, 105, pais_aval, ce.agregarEstiloAzulClaroInfo(libro))
            fila += 1
        return hoja
