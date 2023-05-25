from estilosPF import ClaseEstilos

class ControllerDocumento:
    def llenarCeldasEncabezado(self, datos_encabezado, libro, hoja):
        ce = ClaseEstilos()
        valorMaximo = max([len(valor) for valor in datos_encabezado.values()])

        # Recorrer los valores y escribir en las celdas correspondientes
        fila = 2
        for i in range(valorMaximo):
            try:
                clave_otorgante = int(datos_encabezado.get("clave_otorgante", [''])[i])
            except IndexError:
                clave_otorgante = ' '
            
            try:
                nombre_otorgante = str(datos_encabezado.get("nombre_otorgante", [''])[i])
            except IndexError:
                nombre_otorgante = ''

            try:
                identificador_medio = str(datos_encabezado.get("identificador_medio", [''])[i])
            except IndexError:
                identificador_medio = ''

            try:
                fecha_extraccion = str(datos_encabezado.get("fecha_extraccion", [''])[i])
            except IndexError:
                fecha_extraccion = ''

            try:
                nota_otorgante = str(datos_encabezado.get("nota_otorgante", [''])[i])
            except IndexError:
                nota_otorgante = ''

            try:
                version = str(datos_encabezado.get("version", [''])[i])
            except IndexError:
                version = ''

            if i < len(datos_encabezado["clave_otorgante"]) and len(str(datos_encabezado["clave_otorgante"][i])) > 10:
                clave_otorgante = int(str(datos_encabezado["clave_otorgante"][i])[:10])

            if i < len(datos_encabezado["nombre_otorgante"]) and len(str(nombre_otorgante)) > 75:
                nombre_otorgante = str(nombre_otorgante[:75])

            if i < len(datos_encabezado["identificador_medio"]) and len(str(identificador_medio)) > 10:
                identificador_medio = str(identificador_medio[:10])

            
            if i < len(datos_encabezado["fecha_extraccion"]) and len(str(fecha_extraccion)) > 1:
                 fecha_extraccion = str(datos_encabezado["fecha_extraccion"][i])
                 fecha_extraccion_convertida = int(fecha_extraccion[:2] + fecha_extraccion[3:5] + fecha_extraccion[6:10])
            else:
                fecha_extraccion = ''
                fecha_extraccion_convertida = ''
            fecha_extraccion = fecha_extraccion_convertida


            if i < len(datos_encabezado["nota_otorgante"]) and len(str(nota_otorgante)) > 100:
                nota_otorgante = str(nota_otorgante[:100])

            if i < len(datos_encabezado["version"]) and len(str(datos_encabezado["version"][i])) > 1:
                version = int(str(datos_encabezado["version"][i])[:1])

            #Escribir los datos recibidos y convertidos en las celdas
            hoja.write(fila, 0, clave_otorgante, ce.agregarEstiloNaranjaInfo(libro))
            hoja.write(fila, 1, nombre_otorgante, ce.agregarEstiloNaranjaInfo(libro))
            hoja.write(fila, 2, identificador_medio, ce.agregarEstiloAzulCieloInfo(libro))
            hoja.write(fila, 3, fecha_extraccion, ce.agregarEstiloAzulFuerteInfo(libro))
            hoja.write(fila, 4, nota_otorgante, ce.agregarEstiloAzulCieloInfo(libro))
            hoja.write(fila, 5, version, ce.agregarEstiloAzulFuerteInfo(libro))
            fila += 1
        return hoja

    #-------------------------------Crear función del encabezado Datos Personales-------------------------------------
    def llenarCeldasDatosPersonales(self, datos_dp, libro, hoja):
        ce = ClaseEstilos()
        valorMaximo = max([len(valor) for valor in datos_dp.values()])

        # Recorrer los valores y escribir en las celdas correspondientes
        fila = 2
        variables = ["apellido_paterno_dp", "apellido_materno_dp", "apellido_adicional_dp", "nombres_dp", "fecha_nacimiento_dp", "rfc_dp",
                     "curp_dp", "numero_seguridad_social_dp", "nacionalidad", "residencia", "numero_licencia_conducir_dp", "estado_civil_dp",
                     "sexo_dp", "clave_elector_ife_dp", "numero_dependientes_dp", "fecha_defuncion_dp", "indicador_defuncion_dp", "tipo_person_dp"]

        # Agregar el bucle externo para iterar sobre el rango de valorMaximo
        for i in range(valorMaximo): 
            for variable in variables:
                try:
                    valor = datos_dp.get(variable, [''])[i]

                    if isinstance(valor, int):
                        valor = int(valor)
                    elif isinstance(valor, str):
                        valor = str(valor)

                except IndexError:
                    valor = ''

                if variable == "apellido_paterno_dp":
                    if i < len(datos_dp["apellido_paterno_dp"]) and len(valor) > 30:
                        valor = valor[:30]
                    hoja.write(fila, 6, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "apellido_materno_dp":
                    if i < len(datos_dp["apellido_materno_dp"]) and len(valor) > 30:
                        valor = valor[:10]
                    hoja.write(fila, 7, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "apellido_adicional_dp":
                    if i < len(datos_dp["apellido_adicional_dp"]) and len(valor) > 30:
                        valor = valor[:10]
                    hoja.write(fila, 8, valor, ce.agregarEstiloAzulCieloInfo(libro))

                elif variable == "nombres_dp":
                    if i < len(datos_dp["nombres_dp"]) and len(valor) > 50:
                        valor = valor[:50]
                    hoja.write(fila, 9, valor, ce.agregarEstiloAzulFuerteInfo(libro))


                elif variable == "fecha_nacimiento_dp":
                    if i < len(datos_dp["fecha_nacimiento_dp"]) and len(str(valor)) > 1:
                        valor = str(datos_dp["fecha_nacimiento_dp"][i])
                        fecha_nacimiento_dp = int(valor[:2] + valor[3:5] + valor[6:10])
                    else:
                        valor = ''
                        fecha_nacimiento_dp = ''
                    hoja.write(fila, 10, fecha_nacimiento_dp, ce.agregarEstiloAmarilloInfo(libro))


                elif variable == "rfc_dp":
                    if i < len(datos_dp["rfc_dp"]) and len(valor) > 13:
                        valor = valor[:13]
                    hoja.write(fila, 11, valor, ce.agregarEstiloAmarilloInfo(libro))

                elif variable == "curp_dp":
                    if i < len(datos_dp["curp_dp"]) and len(valor) > 18:
                        valor = valor[:18]
                    hoja.write(fila, 12, valor, ce.agregarEstiloAzulCieloInfo(libro))

                elif variable == "numero_seguridad_social_dp":
                    if i < len(datos_dp["numero_seguridad_social_dp"]) and len(str(valor)) > 11:
                        valor = int(valor[:11])
                    hoja.write(fila, 13, valor, ce.agregarEstiloAzulCieloInfo(libro))

                elif variable == "nacionalidad":
                    if i < len(datos_dp["nacionalidad"]) and len(valor) > 2:
                        valor = valor[:2]
                    hoja.write(fila, 14, valor, ce.agregarEstiloAmarilloInfo(libro))

                elif variable == "residencia":
                    if i < len(datos_dp["residencia"]) and len(str(valor)) > 1:
                        valor = int(valor[:1])
                    hoja.write(fila, 15, valor, ce.agregarEstiloAzulCieloInfo(libro))

                elif variable == "numero_licencia_conducir_dp":
                    if i < len(datos_dp["numero_licencia_conducir_dp"]) and len(valor) > 20:
                        valor = valor[:20]
                    hoja.write(fila, 16, valor, ce.agregarEstiloAzulCieloInfo(libro))

                elif variable == "estado_civil_dp":
                    if i < len(datos_dp["estado_civil_dp"]) and len(valor) > 1:
                        valor = valor[:1]
                    hoja.write(fila, 17, valor, ce.agregarEstiloAzulCieloInfo(libro))

                elif variable == "sexo_dp":
                    if i < len(datos_dp["sexo_dp"]) and len(valor) > 1:
                        valor = valor[:1]
                    hoja.write(fila, 18, valor, ce.agregarEstiloAzulCieloInfo(libro))

                elif variable == "clave_elector_ife_dp":
                    if i < len(datos_dp["clave_elector_ife_dp"]) and len(valor) > 20:
                        valor = valor[:20]
                    hoja.write(fila, 19, valor, ce.agregarEstiloAzulCieloInfo(libro))

                elif variable == "numero_dependientes_dp":
                    if i < len(datos_dp["numero_dependientes_dp"]) and len(str(valor)) > 2:
                        valor = int(valor[:2])
                    hoja.write(fila, 20, valor, ce.agregarEstiloAzulCieloInfo(libro))


                elif variable == "fecha_defuncion_dp":
                    if i < len(datos_dp["fecha_defuncion_dp"]) and len(str(valor)) > 1:
                        valor = str(datos_dp["fecha_defuncion_dp"][i])
                        fecha_defuncion_dp = int(valor[:2] + valor[3:5] + valor[6:10])
                    else:
                        valor = ''
                        fecha_defuncion_dp = ''
                    hoja.write(fila, 21, fecha_defuncion_dp, ce.agregarEstiloAzulCieloInfo(libro))


                elif variable == "indicador_defuncion_dp":
                    if i < len(datos_dp["indicador_defuncion_dp"]) and len(valor) > 1:
                        valor = valor[:1]
                    hoja.write(fila, 22, valor, ce.agregarEstiloAzulCieloInfo(libro))

                elif variable == "tipo_person_dp":
                    if i < len(datos_dp["tipo_person_dp"]) and len(valor) > 2:
                        valor = valor[:1]
                    hoja.write(fila, 23, valor, ce.agregarEstiloAzulCieloInfo(libro))
            fila += 1

        return hoja
            