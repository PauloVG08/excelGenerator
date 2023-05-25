from estilosPF import ClaseEstilos

class ControllerDocumento:
    #-------------------------------Crear función del encabezado Encabezado-------------------------------------
    def llenarCeldasEncabezado(self, datos_encabezado, libro, hoja):
        ce = ClaseEstilos()
        valorMaximo = max([len(valor) for valor in datos_encabezado.values()])
        variables = ["clave_otorgante", "nombre_otorgante", "identificador_medio", "fecha_extraccion",
                        "nota_otorgante", "version"]

        # Recorrer los valores y escribir en las celdas correspondientes
        fila = 2
        for i in range(valorMaximo):
            for variable in variables:
                try:
                    valor = datos_encabezado.get(variable, [''])[i]

                    if isinstance(valor, int):
                        valor = int(valor)
                    elif isinstance(valor, str):
                        valor = str(valor)

                except IndexError:
                    valor = '' 

                if variable == "clave_otorgante":
                    if i < len(datos_encabezado["clave_otorgante"]) and len(str(valor)) > 10:
                        valor = int(valor[:10])
                    hoja.write(fila, 0, valor, ce.agregarEstiloNaranjaInfo(libro))

                elif variable == "nombre_otorgante":
                    if i < len(datos_encabezado["nombre_otorgante"]) and len(valor) > 40:
                        valor = valor[:40]
                    hoja.write(fila, 1, valor, ce.agregarEstiloNaranjaInfo(libro))

                elif variable == "identificador_medio":
                    if i < len(datos_encabezado["identificador_medio"]) and len(valor) > 10:
                        valor = valor[:10]
                    hoja.write(fila, 2, valor, ce.agregarEstiloAzulCieloInfo(libro))


                elif variable == "fecha_extraccion":
                    if i < len(datos_encabezado["fecha_extraccion"]) and len(str(valor)) > 1:
                        valor = str(datos_encabezado["fecha_extraccion"][i])
                        fecha_extraccion = int(valor[:2] + valor[3:5] + valor[6:10])
                    else:
                        valor = ''
                        fecha_extraccion = ''
                    hoja.write(fila, 3, fecha_extraccion, ce.agregarEstiloAzulFuerteInfo(libro))


                elif variable == "nota_otorgante":
                    if i < len(datos_encabezado["nota_otorgante"]) and len(valor) > 100:
                        valor = valor[:100]
                    hoja.write(fila, 4, valor, ce.agregarEstiloAzulCieloInfo(libro))

                if variable == "version":
                    if i < len(datos_encabezado["version"]) and len(str(valor)) >= 1:
                        valor_str = str(valor)
                        primer_digito = valor_str[0]
                        valor = int(primer_digito)
                    else:
                        valor = ''
                    hoja.write(fila, 5, valor, ce.agregarEstiloAzulFuerteInfo(libro))
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
            
    
    #-------------------------------Crear función del encabezado Domicilio--------------------------------------------
    def llenarCeldasDomicilio(self, datos_dom, libro, hoja):
        ce = ClaseEstilos()
        valorMaximo = max([len(valor) for valor in datos_dom.values()])
        variables = ["direccion_dom", "colonia_poblacion_dom", "deleg_muni_dom", "ciudad_dom",
                        "estado_dom", "cp_dom", "fecha_residencia_dom", "telefono_dom",
                        "tipo_dom", "tipo_asentamiento", "origen_dom"]

        # Recorrer los valores y escribir en las celdas correspondientes
        fila = 2
        for i in range(valorMaximo):
            for variable in variables:
                try:
                    valor = datos_dom.get(variable, [''])[i]

                    if isinstance(valor, int):
                        valor = int(valor)
                    elif isinstance(valor, str):
                        valor = str(valor)

                except IndexError:
                    valor = '' 

                if variable == "direccion_dom":
                    if i < len(datos_dom["direccion_dom"]) and len(str(valor)) >= 80:
                        valor = valor[:80]
                    hoja.write(fila, 24, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "colonia_poblacion_dom":
                    if i < len(datos_dom["colonia_poblacion_dom"]) and len(str(valor)) >= 65:
                        valor = valor[:65]
                    hoja.write(fila, 25, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "deleg_muni_dom":
                    if i < len(datos_dom["deleg_muni_dom"]) and len(str(valor)) >= 65:
                        valor = valor[:65]
                    hoja.write(fila, 26, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "ciudad_dom":
                    if i < len(datos_dom["ciudad_dom"]) and len(str(valor)) >= 65:
                        valor = valor[:65]
                    hoja.write(fila, 27, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "estado_dom":
                    if i < len(datos_dom["estado_dom"]) and len(str(valor)) >= 4:
                        valor = valor[:4]
                    hoja.write(fila, 28, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "cp_dom":
                    if i < len(datos_dom["cp_dom"]) and len(str(valor)) >= 5:
                        valor = str(valor)[:5]
                        if valor:
                            hoja.write(fila, 29, int(valor), ce.agregarEstiloAzulFuerteInfo(libro))
                        else:
                            hoja.write(fila, 29, valor, ce.agregarEstiloAzulFuerteInfo(libro))
                    else:
                        hoja.write(fila, 29, '', ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "fecha_residencia_dom":
                    if i < len(datos_dom["fecha_residencia_dom"]) and len(str(valor)) > 0:
                        valor = str(datos_dom["fecha_residencia_dom"][i])
                        fecha_residencia_dom = int(valor[:2] + valor[3:5] + valor[6:10])
                    else:
                        valor = ''
                        fecha_residencia_dom = ''
                    hoja.write(fila, 30, fecha_residencia_dom, ce.agregarEstiloAzulCieloInfo(libro))


                elif variable == "telefono_dom":
                    if i < len(datos_dom["telefono_dom"]) and len(str(valor)) >= 20:
                        valor = int(valor[:20])
                    hoja.write(fila, 31, valor, ce.agregarEstiloAzulCieloInfo(libro))

                elif variable == "tipo_dom":
                    if i < len(datos_dom["tipo_dom"]) and len(str(valor)) >= 1:
                        valor = valor[:1]
                    hoja.write(fila, 32, valor, ce.agregarEstiloAzulCieloInfo(libro))
                
                elif variable == "tipo_asentamiento":
                    if i < len(datos_dom["tipo_asentamiento"]) and len(str(valor)) > 0:
                        valor = str(valor)[:5]
                        if valor:
                            hoja.write(fila, 33, int(valor), ce.agregarEstiloAzulCieloInfo(libro))
                        else:
                            hoja.write(fila, 33, valor, ce.agregarEstiloAzulCieloInfo(libro))
                    else:
                        hoja.write(fila, 33, '', ce.agregarEstiloAzulCieloInfo(libro))

                elif variable == "origen_dom":
                    if i < len(datos_dom["origen_dom"]) and len(str(valor)) >= 2:
                        valor = valor[:2]
                    hoja.write(fila, 34, valor, ce.agregarEstiloAzulCieloInfo(libro))

            fila += 1
        return hoja
    
    #-------------------------------Crear función del encabezado Empleo--------------------------------------------
    def llenarCeldasEmpleo(self, datos_emp, libro, hoja):
        ce = ClaseEstilos()
        valorMaximo = max([len(valor) for valor in datos_emp.values()])
        variables = ["nombre_empresa_emp", "direccion_emp", "colonia_poblacional_emp", "deleg_mun",
                        "ciudad_emp", "estado_emp", "cp_emp", "num_telefono_emp", "extension_emp",
                        "fax_emp", "puesto_emp", "fecha_contratacion_emp", "clave_moneda_emp",
                        "salario_mensual_emp", "fecha_ultimo_dia_emp", "fecha_verificacion_emp", "origen_razon_social_emp"]

        # Recorrer los valores y escribir en las celdas correspondientes
        fila = 2
        for i in range(valorMaximo):
            for variable in variables:
                try:
                    valor = datos_emp.get(variable, [''])[i]

                    if isinstance(valor, int):
                        valor = int(valor)
                    elif isinstance(valor, str):
                        valor = str(valor)

                except IndexError:
                    valor = '' 


                if variable == "nombre_empresa_emp":
                    if i < len(datos_emp["nombre_empresa_emp"]) and len(str(valor)) >= 40:
                        valor = valor[:40]
                    hoja.write(fila, 35, valor, ce.agregarEstiloAmarilloInfo(libro))

                elif variable == "direccion_emp":
                    if i < len(datos_emp["direccion_emp"]) and len(str(valor)) >= 80:
                        valor = valor[:80]
                    hoja.write(fila, 36, valor, ce.agregarEstiloAmarilloInfo(libro))

                elif variable == "colonia_poblacional_emp":
                    if i < len(datos_emp["colonia_poblacional_emp"]) and len(str(valor)) >= 65:
                        valor = valor[:65]
                    hoja.write(fila, 37, valor, ce.agregarEstiloAmarilloInfo(libro))

                elif variable == "deleg_mun":
                    if i < len(datos_emp["deleg_mun"]) and len(str(valor)) >= 65:
                        valor = valor[:65]
                    hoja.write(fila, 38, valor, ce.agregarEstiloAmarilloInfo(libro))

                elif variable == "ciudad_emp":
                    if i < len(datos_emp["ciudad_emp"]) and len(str(valor)) >= 65:
                        valor = valor[:65]
                    hoja.write(fila, 39, valor, ce.agregarEstiloAmarilloInfo(libro))

                elif variable == "estado_emp":
                    if i < len(datos_emp["estado_emp"]) and len(str(valor)) >= 4:
                        valor = valor[:4]
                    hoja.write(fila, 40, valor, ce.agregarEstiloAmarilloInfo(libro))

                elif variable == "cp_emp":
                    if i < len(datos_emp["cp_emp"]) and len(str(valor)) > 0:
                        valor = str(valor)[:5]
                        if valor:
                            hoja.write(fila, 41, int(valor), ce.agregarEstiloAmarilloInfo(libro))
                        else:
                            hoja.write(fila, 41, valor, ce.agregarEstiloAmarilloInfo(libro))
                    else:
                        hoja.write(fila, 41, '', ce.agregarEstiloAmarilloInfo(libro))

                elif variable == "num_telefono_emp":
                    if i < len(datos_emp["num_telefono_emp"]) and len(str(valor)) > 20:
                        valor = int(valor[:20])
                    hoja.write(fila, 42, valor, ce.agregarEstiloAzulCieloInfo(libro))

                elif variable == "extension_emp":
                    if i < len(datos_emp["extension_emp"]) and len(str(valor)) > 8:
                        valor = int(valor[:8])
                    hoja.write(fila, 43, valor, ce.agregarEstiloAzulCieloInfo(libro))

                elif variable == "fax_emp":
                    if i < len(datos_emp["fax_emp"]) and len(str(valor)) > 20:
                        valor = int(valor[:20])
                    hoja.write(fila, 44, valor, ce.agregarEstiloAzulCieloInfo(libro))




                elif variable == "puesto_emp":
                    if i < len(datos_emp["puesto_emp"]) and len(str(valor)) > 30:
                        valor = valor[:30]
                    hoja.write(fila, 45, valor, ce.agregarEstiloAzulCieloInfo(libro))


                elif variable == "fecha_contratacion_emp":
                    if i < len(datos_emp["fecha_contratacion_emp"]) and len(str(valor)) > 1:
                        valor = str(datos_emp["fecha_contratacion_emp"][i])
                        fecha_contratacion_emp = int(valor[:2] + valor[3:5] + valor[6:10])
                    else:
                        valor = ''
                        fecha_contratacion_emp = ''
                    hoja.write(fila, 46, fecha_contratacion_emp, ce.agregarEstiloAzulCieloInfo(libro))


                elif variable == "clave_moneda_emp":
                    if i < len(datos_emp["clave_moneda_emp"]) and len(str(valor)) == 2:
                        valor = valor[:2]
                    hoja.write(fila, 47, valor, ce.agregarEstiloAzulCieloInfo(libro))

                elif variable == "salario_mensual_emp":
                    if i < len(datos_emp["salario_mensual_emp"]) and len(str(valor)) > 0:
                        valor = str(valor)[:5]
                        if valor:
                            hoja.write(fila, 48, int(valor), ce.agregarEstiloAzulCieloInfo(libro))
                        else:
                            hoja.write(fila, 48, valor, ce.agregarEstiloAzulCieloInfo(libro))
                    else:
                        hoja.write(fila, 48, '', ce.agregarEstiloAzulCieloInfo(libro))


                elif variable == "fecha_ultimo_dia_emp":
                    if i < len(datos_emp["fecha_ultimo_dia_emp"]) and len(str(valor)) > 1:
                        valor = str(datos_emp["fecha_ultimo_dia_emp"][i])
                        fecha_ultimo_dia_emp = int(valor[:2] + valor[3:5] + valor[6:10])
                    else:
                        valor = ''
                        fecha_ultimo_dia_emp = ''
                    hoja.write(fila, 49, fecha_ultimo_dia_emp, ce.agregarEstiloAzulCieloInfo(libro))


                elif variable == "fecha_verificacion_emp":
                    if i < len(datos_emp["fecha_verificacion_emp"]) and len(str(valor)) > 1:
                        valor = str(datos_emp["fecha_verificacion_emp"][i])
                        fecha_verificacion_emp = int(valor[:2] + valor[3:5] + valor[6:10])
                    else:
                        valor = ''
                        fecha_verificacion_emp = ''
                    hoja.write(fila, 50, fecha_verificacion_emp, ce.agregarEstiloAzulCieloInfo(libro))


                elif variable == "origen_razon_social_emp":
                    if i < len(datos_emp["origen_razon_social_emp"]) and len(str(valor)) == 2:
                        valor = valor[:2]
                    hoja.write(fila, 51, valor, ce.agregarEstiloAzulCieloInfo(libro))
            fila += 1
        return hoja
