from estilosPF import ClaseEstilos
import textwrap

class ControllerDocumento:
    #-------------------------------Crear función para obtener valor maximo--------------------------------------
    def obtenerValorMaximo(self, datos_encabezado, datos_dp, datos_dom, datos_emp, datos_dc, datos_cc):
        maximo_datos_encabezado = 0
        maximo_datos_dp = 0
        maximo_datos_dom = 0
        maximo_datos_emp = 0
        maximo_datos_dc = 0
        maximo_datos_cc = 0

        for valores in datos_encabezado.values():
            cantidad_actual = len(valores)
            if cantidad_actual > maximo_datos_encabezado:
                maximo_datos_encabezado = cantidad_actual

        for valores in datos_dp.values():
            cantidad_actual = len(valores)
            if cantidad_actual > maximo_datos_dp:
                maximo_datos_dp = cantidad_actual

        for valores in datos_dom.values():
            cantidad_actual = len(valores)
            if cantidad_actual > maximo_datos_dom:
                maximo_datos_dom = cantidad_actual

        for valores in datos_emp.values():
            cantidad_actual = len(valores)
            if cantidad_actual > maximo_datos_emp:
                maximo_datos_emp = cantidad_actual

        for valores in datos_dc.values():
            cantidad_actual = len(valores)
            if cantidad_actual > maximo_datos_dc:
                maximo_datos_dc = cantidad_actual

        for valores in datos_cc.values():
            cantidad_actual = len(valores)
            if cantidad_actual > maximo_datos_cc:
                maximo_datos_cc = cantidad_actual
        
        valorMaximo = max(maximo_datos_encabezado, maximo_datos_emp, maximo_datos_dp, maximo_datos_dc, maximo_datos_dom, maximo_datos_cc)
        return valorMaximo

    #-------------------------------Crear función del encabezado Encabezado-------------------------------------
    def llenarCeldasEncabezado(self, datos_encabezado, datos_dp, datos_dom, datos_emp, datos_dc, datos_cc, libro, hoja):
        ce = ClaseEstilos()
        variables = ["clave_otorgante", "nombre_otorgante", "identificador_medio", "fecha_extraccion",
                        "nota_otorgante", "version"]
        
        valorMaximo = self.obtenerValorMaximo(datos_encabezado, datos_dp, datos_dom, datos_emp, datos_dc, datos_cc)
        for variable, valores in datos_encabezado.items():
            cantidad_actual = len(valores)
            if cantidad_actual > valorMaximo:
                valorMaximo = cantidad_actual

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
                        fecha_extraccion = str(valor[:2] + valor[3:5] + valor[6:10])
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
    def llenarCeldasDatosPersonales(self, datos_encabezado, datos_dp, datos_dom, datos_emp, datos_dc, datos_cc, libro, hoja):

        ce = ClaseEstilos()

        # Recorrer los valores y escribir en las celdas correspondientes
        fila = 2
        variables = ["apellido_paterno_dp", "apellido_materno_dp", "apellido_adicional_dp", "nombres_dp", "fecha_nacimiento_dp", "rfc_dp",
                     "curp_dp", "numero_seguridad_social_dp", "nacionalidad", "residencia", "numero_licencia_conducir_dp", "estado_civil_dp",
                     "sexo_dp", "clave_elector_ife_dp", "numero_dependientes_dp", "fecha_defuncion_dp", "indicador_defuncion_dp", "tipo_person_dp"]


        valorMaximo = self.obtenerValorMaximo(datos_encabezado, datos_dp, datos_dom, datos_emp, datos_dc, datos_cc)
        for variable, valores in datos_dp.items():
            cantidad_actual = len(valores)
            if cantidad_actual > valorMaximo:
                valorMaximo = cantidad_actual


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
                        fecha_nacimiento_dp = str(valor[:2] + valor[3:5] + valor[6:10])
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
                        fecha_defuncion_dp = str(valor[:2] + valor[3:5] + valor[6:10])
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
    def llenarCeldasDomicilio(self, datos_encabezado, datos_dp, datos_dom, datos_emp, datos_dc, datos_cc, libro, hoja):
        ce = ClaseEstilos()
        variables = ["direccion_dom", "colonia_poblacion_dom", "deleg_muni_dom", "ciudad_dom",
                        "estado_dom", "cp_dom", "fecha_residencia_dom", "telefono_dom",
                        "tipo_dom", "tipo_asentamiento", "origen_dom"]
        

        valorMaximo = self.obtenerValorMaximo(datos_encabezado, datos_dp, datos_dom, datos_emp, datos_dc, datos_cc)
        for variable, valores in datos_dom.items():
            cantidad_actual = len(valores)
            if cantidad_actual > valorMaximo:
                valorMaximo = cantidad_actual


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
                        fecha_residencia_dom = str(valor[:2] + valor[3:5] + valor[6:10])
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
                        hoja.write(fila, 33, '', ce.agregarEstiloAzulCieloInfo(libro))

                elif variable == "origen_dom":
                    if i < len(datos_dom["origen_dom"]) and len(str(valor)) >= 2:
                        valor = valor[:2]
                    hoja.write(fila, 34, valor, ce.agregarEstiloAzulCieloInfo(libro))

            fila += 1
        return hoja
    
    #-------------------------------Crear función del encabezado Empleo--------------------------------------------
    def llenarCeldasEmpleo(self, datos_encabezado, datos_dp, datos_dom, datos_emp, datos_dc, datos_cc, libro, hoja): 
        ce = ClaseEstilos()
        variables = ["nombre_empresa_emp", "direccion_emp", "colonia_poblacional_emp", "deleg_mun",
                        "ciudad_emp", "estado_emp", "cp_emp", "num_telefono_emp", "extension_emp",
                        "fax_emp", "puesto_emp", "fecha_contratacion_emp", "clave_moneda_emp",
                        "salario_mensual_emp", "fecha_ultimo_dia_emp", "fecha_verificacion_emp", "origen_razon_social_emp"]
        
        valorMaximo = self.obtenerValorMaximo(datos_encabezado, datos_dp, datos_dom, datos_emp, datos_dc, datos_cc)
        for variable, valores in datos_emp.items():
            cantidad_actual = len(valores)
            if cantidad_actual > valorMaximo:
                valorMaximo = cantidad_actual

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
                        fecha_contratacion_emp = str(valor[:2] + valor[3:5] + valor[6:10])
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
                        hoja.write(fila, 48, '', ce.agregarEstiloAzulCieloInfo(libro))


                elif variable == "fecha_ultimo_dia_emp":
                    if i < len(datos_emp["fecha_ultimo_dia_emp"]) and len(str(valor)) > 1:
                        valor = str(datos_emp["fecha_ultimo_dia_emp"][i])
                        fecha_ultimo_dia_emp = str(valor[:2] + valor[3:5] + valor[6:10])
                    else:
                        valor = ''
                        fecha_ultimo_dia_emp = ''
                    hoja.write(fila, 49, fecha_ultimo_dia_emp, ce.agregarEstiloAzulCieloInfo(libro))


                elif variable == "fecha_verificacion_emp":
                    if i < len(datos_emp["fecha_verificacion_emp"]) and len(str(valor)) > 1:
                        valor = str(datos_emp["fecha_verificacion_emp"][i])
                        fecha_verificacion_emp = str(valor[:2] + valor[3:5] + valor[6:10])
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

    #-------------------------------Crear función del encabezado Detalles crédito--------------------------------------------
    def llenarCeldasDC(self, datos_encabezado, datos_dp, datos_dom, datos_emp, datos_dc, datos_cc, libro, hoja):
        ce = ClaseEstilos()
        variables = ["clave_actual_dc", "nombre_otorgante_dc", "cuenta_actual_dc", "tipo_responsabilidad_dc",
                     "tipo_cuenta_dc", "tipo_contrato_dc", "clave_uni_monetaria_dc", "valor_activo_dc", 
                     "num_pagos_dc", "frecuencia_pagos_dc", "monto_pagar_dc", "fecha_apertura_dc",
                     "fecha_ultimo_pago_dc", "fecha_utlima_compra_dc", "fecha_cierre_cuenta_dc", "fecha_corte_dc",
                     "garantia_dc", "credito_maximo_dc", "saldo_actual_dc", "limite_credito_dc",
                     "saldo_vencido_dc", "num_pagos_vencidos_dc", "pago_actual_dc", "historico_pagos_dc",
                     "clave_prevencion_dc", "total_pagos_rep_dc", "clave_anterior_otor_dc", "nombre_anterior_otor_dc",
                     "num_cuenta_anterior_dc", "fecha_primer_incump_dc", "saldo_insoluto_dc", "monto_ultimo_pago",
                     "fecha_ingreso_carterav_dc", "monto_intereses_dc", "forma_pago_ac_int_dc", "dias_vencimiento_dc",
                     "plazo_meses_dc", "monto_creditoOri_dc", "correo_consumidor_dc"]
        
        valorMaximo = self.obtenerValorMaximo(datos_encabezado, datos_dp, datos_dom, datos_emp, datos_dc, datos_cc)
        for variable, valores in datos_dc.items():
            cantidad_actual = len(valores)
            if cantidad_actual > valorMaximo:
                valorMaximo = cantidad_actual
                    
        # Recorrer los valores y escribir en las celdas correspondientes
        fila = 2
        for i in range(valorMaximo):
            for variable in variables:
                try:
                    valor = datos_dc.get(variable, [''])[i]

                    if isinstance(valor, int):
                        valor = int(valor)
                    elif isinstance(valor, str):
                        valor = str(valor)
                except IndexError:
                    valor = '' 
            

                if variable == "clave_actual_dc":
                    if i < len(datos_dc["clave_actual_dc"]) and len(str(valor)) > 0:
                        valor = str(valor)[:10]
                        if valor:
                            hoja.write(fila, 52, int(valor), ce.agregarEstiloAmarilloInfo(libro))
                    else:
                        hoja.write(fila, 52, '', ce.agregarEstiloAmarilloInfo(libro))
                    
                elif variable == "nombre_otorgante_dc":
                    if i < len(datos_dc["nombre_otorgante_dc"]) and len(str(valor)) > 40:
                        valor = valor[:40]
                    hoja.write(fila, 53, valor, ce.agregarEstiloAmarilloInfo(libro))
                
                elif variable == "cuenta_actual_dc":
                    if i < len(datos_dc["cuenta_actual_dc"]) and len(str(valor)) > 25:
                        valor = valor[:25]
                    hoja.write(fila, 54, valor, ce.agregarEstiloAzulFuerteInfo(libro))
                
                elif variable == "tipo_responsabilidad_dc":
                    if i < len(datos_dc["tipo_responsabilidad_dc"]) and len(str(valor)) > 1:
                        valor = valor[:1]
                    hoja.write(fila, 55, valor, ce.agregarEstiloAzulFuerteInfo(libro))
                
                elif variable == "tipo_cuenta_dc":
                    if i < len(datos_dc["tipo_cuenta_dc"]) and len(str(valor)) > 1:
                        valor = valor[:1]
                    hoja.write(fila, 56, valor, ce.agregarEstiloAzulFuerteInfo(libro))
                
                elif variable == "tipo_contrato_dc":
                    if i < len(datos_dc["tipo_contrato_dc"]) and len(str(valor)) > 2:
                        valor = valor[:2]
                    hoja.write(fila, 57, valor, ce.agregarEstiloAzulFuerteInfo(libro))
                
                elif variable == "clave_uni_monetaria_dc":
                    if i < len(datos_dc["clave_uni_monetaria_dc"]) and len(str(valor)) > 2:
                        valor = valor[:2]
                    hoja.write(fila, 58, valor, ce.agregarEstiloAzulFuerteInfo(libro))
                
                if variable == "valor_activo_dc":
                    if i < len(datos_dc["valor_activo_dc"]) and len(str(valor)) > 0:
                        valor = str(valor)[:9]
                        if valor:
                            hoja.write(fila, 59, int(valor), ce.agregarEstiloAzulCieloInfo(libro))
                    else:
                        hoja.write(fila, 59, '', ce.agregarEstiloAzulCieloInfo(libro))

                if variable == "num_pagos_dc":
                    if i < len(datos_dc["num_pagos_dc"]) and len(str(valor)) > 0:
                        valor = str(valor)[:4]
                        if valor:
                            hoja.write(fila, 60, int(valor), ce.agregarEstiloAzulFuerteInfo(libro))
                    else:
                        hoja.write(fila, 60, '', ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "frecuencia_pagos_dc":
                    if i < len(datos_dc["frecuencia_pagos_dc"]) and len(str(valor)) > 1:
                        valor = valor[:1]
                    hoja.write(fila, 61, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                if variable == "monto_pagar_dc":
                    if i < len(datos_dc["monto_pagar_dc"]) and len(str(valor)) > 0:
                        valor = str(valor)[:9]
                        if valor:
                            hoja.write(fila, 62, int(valor), ce.agregarEstiloAzulFuerteInfo(libro))
                    else:
                        hoja.write(fila, 62, '', ce.agregarEstiloAzulFuerteInfo(libro))

                
                elif variable == "fecha_apertura_dc":
                    if i < len(datos_dc["fecha_apertura_dc"]) and len(str(valor)) > 0:
                        valor = str(datos_dc["fecha_apertura_dc"][i])
                        fecha_apertura_dc = str(valor[:2] + valor[3:5] + valor[6:10])
                    else:
                        valor = ''
                        fecha_apertura_dc = ''
                    hoja.write(fila, 63, fecha_apertura_dc, ce.agregarEstiloAzulFuerteInfo(libro))

                
                elif variable == "fecha_ultimo_pago_dc":
                    if i < len(datos_dc["fecha_ultimo_pago_dc"]) and len(str(valor)) > 0:
                        valor = str(datos_dc["fecha_ultimo_pago_dc"][i])
                        fecha_ultimo_pago_dc = str(valor[:2] + valor[3:5] + valor[6:10])
                    else:
                        valor = ''
                        fecha_ultimo_pago_dc = ''
                    hoja.write(fila, 64, fecha_ultimo_pago_dc, ce.agregarEstiloAzulFuerteInfo(libro))


                elif variable == "fecha_utlima_compra_dc":
                    if i < len(datos_dc["fecha_utlima_compra_dc"]) and len(str(valor)) > 0:
                        valor = str(datos_dc["fecha_utlima_compra_dc"][i])
                        fecha_utlima_compra_dc = str(valor[:2] + valor[3:5] + valor[6:10])
                    else:
                        valor = ''
                        fecha_utlima_compra_dc = ''
                    hoja.write(fila, 65, fecha_utlima_compra_dc, ce.agregarEstiloAzulFuerteInfo(libro))


                elif variable == "fecha_cierre_cuenta_dc":
                    if i < len(datos_dc["fecha_cierre_cuenta_dc"]) and len(str(valor)) > 0:
                        valor = str(datos_dc["fecha_cierre_cuenta_dc"][i])
                        fecha_cierre_cuenta = str(valor[:2] + valor[3:5] + valor[6:10])
                    else:
                        valor = ''
                        fecha_cierre_cuenta = ''
                    hoja.write(fila, 66, fecha_cierre_cuenta, ce.agregarEstiloAzulCieloInfo(libro))


                elif variable == "fecha_corte_dc":
                    if i < len(datos_dc["fecha_corte_dc"]) and len(str(valor)) > 0:
                        valor = str(datos_dc["fecha_corte_dc"][i])
                        fecha_corte_dc = str(valor[:2] + valor[3:5] + valor[6:10])
                    else:
                        valor = ''
                        fecha_corte_dc = ''
                    hoja.write(fila, 67, fecha_corte_dc, ce.agregarEstiloAzulFuerteInfo(libro))


                elif variable == "garantia_dc":
                    if i < len(datos_dc["garantia_dc"]) and len(str(valor)) > 200:
                        valor = valor[:200]
                    # Dividir el texto en líneas de máximo 30 caracteres
                    lineas = textwrap.wrap(valor, width=30)
                    # Escribir las líneas en la misma celda con saltos de línea
                    hoja.write(fila, 68, "\n".join(lineas), ce.agregarEstiloAzulCieloInfo(libro))

                if variable == "credito_maximo_dc":
                    if i < len(datos_dc["credito_maximo_dc"]) and len(str(valor)) > 0:
                        valor = str(valor)[:9]
                        if valor:
                            hoja.write(fila, 69, int(valor), ce.agregarEstiloAzulFuerteInfo(libro))
                    else:
                        hoja.write(fila, 69, '', ce.agregarEstiloAzulFuerteInfo(libro))

                if variable == "saldo_actual_dc":
                    if i < len(datos_dc["saldo_actual_dc"]) and len(str(valor)) > 10:
                        valor = str(valor)[:10]
                        if valor:
                            hoja.write(fila, 70, int(valor), ce.agregarEstiloAzulFuerteInfo(libro))
                        else:
                            hoja.write(fila, 70, valor, ce.agregarEstiloAzulFuerteInfo(libro))
                    else:
                        hoja.write(fila, 70, '', ce.agregarEstiloAzulFuerteInfo(libro))

                if variable == "limite_credito_dc":
                    if i < len(datos_dc["limite_credito_dc"]) and len(str(valor)) > 0:
                        valor = str(valor)[:9]
                        if valor:
                            hoja.write(fila, 71, int(valor), ce.agregarEstiloAzulFuerteInfo(libro))
                    else:
                        hoja.write(fila, 71, '', ce.agregarEstiloAzulFuerteInfo(libro))
                        
                if variable == "saldo_vencido_dc":
                    if i < len(datos_dc["saldo_vencido_dc"]) and len(str(valor)) > 0:
                        valor = str(valor)[:9]
                        if valor:
                            hoja.write(fila, 72, int(valor), ce.agregarEstiloAzulFuerteInfo(libro))
                    else:
                        hoja.write(fila, 72, '', ce.agregarEstiloAzulFuerteInfo(libro))

                if variable == "num_pagos_vencidos_dc":
                    if i < len(datos_dc["num_pagos_vencidos_dc"]) and len(str(valor)) > 0:
                        valor = str(valor)[:4]
                        if valor:
                            hoja.write(fila, 73, int(valor), ce.agregarEstiloAzulCieloInfo(libro))
                    else:
                        hoja.write(fila, 73, '', ce.agregarEstiloAzulCieloInfo(libro))

                elif variable == "pago_actual_dc":
                    if i < len(datos_dc["pago_actual_dc"]) and len(str(valor)) > 2:
                        valor = valor[:2]
                    hoja.write(fila, 74, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "historico_pagos_dc":
                    if i < len(datos_dc["historico_pagos_dc"]) and len(str(valor)) > 168:
                        valor = valor[:168]
                    # Dividir el texto en líneas de máximo 30 caracteres
                    lineas = textwrap.wrap(valor, width=30)
                    # Escribir las líneas en la misma celda con saltos de línea
                    hoja.write(fila, 75, "\n".join(lineas), ce.agregarEstiloAzulCieloInfo(libro))

                elif variable == "clave_prevencion_dc":
                    if i < len(datos_dc["clave_prevencion_dc"]) and len(str(valor)) > 2:
                        valor = valor[:2]
                    hoja.write(fila, 76, valor, ce.agregarEstiloAzulCieloInfo(libro))

                if variable == "total_pagos_rep_dc":
                    if i < len(datos_dc["total_pagos_rep_dc"]) and len(str(valor)) > 0:
                        valor = str(valor)[:3]
                        if valor:
                            hoja.write(fila, 77, int(valor), ce.agregarEstiloAzulCieloInfo(libro))
                    else:
                        hoja.write(fila, 77, '', ce.agregarEstiloAzulCieloInfo(libro))

                if variable == "clave_anterior_otor_dc":
                    if i < len(datos_dc["clave_anterior_otor_dc"]) and len(str(valor)) > 0:
                        valor = str(valor)[:10]
                        if valor:
                            hoja.write(fila, 78, int(valor), ce.agregarEstiloAzulCieloInfo(libro))
                    else:
                        hoja.write(fila, 78, '', ce.agregarEstiloAzulCieloInfo(libro))
                #----------------ALERTA DUDA-------------------------------------------
                elif variable == "nombre_anterior_otor_dc":
                    if i < len(datos_dc["nombre_anterior_otor_dc"]) and len(str(valor)) > 25:
                        valor = valor[:25]
                    hoja.write(fila, 79, valor, ce.agregarEstiloAzulCieloInfo(libro))  

                #----------------ALERTA DUDA-------------------------------------------
                elif variable == "num_cuenta_anterior_dc":
                    if i < len(datos_dc["num_cuenta_anterior_dc"]) and len(str(valor)) > 0:
                        valor = str(valor)[:10]
                        if valor:
                            hoja.write(fila, 80, int(valor), ce.agregarEstiloAzulCieloInfo(libro))
                    else:
                        hoja.write(fila, 80, '', ce.agregarEstiloAzulCieloInfo(libro)) 


                elif variable == "fecha_primer_incump_dc":
                    if i < len(datos_dc["fecha_primer_incump_dc"]) and len(str(valor)) > 0:
                        valor = str(datos_dc["fecha_primer_incump_dc"][i])
                        fecha_primer_incump_dc = str(valor[:2] + valor[3:5] + valor[6:10])
                    else:
                        valor = ''
                        fecha_primer_incump_dc = ''
                    hoja.write(fila, 81, fecha_primer_incump_dc, ce.agregarEstiloAmarilloInfo(libro))


                elif variable == "saldo_insoluto_dc":
                    if i < len(datos_dc["saldo_insoluto_dc"]) and len(str(valor)) > 0:
                        valor = str(valor)[:10]
                        if valor:
                            hoja.write(fila, 82, int(valor), ce.agregarEstiloAzulFuerteInfo(libro))
                        else:
                            hoja.write(fila, 82, valor, ce.agregarEstiloAzulFuerteInfo(libro))
                    else:
                        hoja.write(fila, 82, '', ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "monto_ultimo_pago":
                    if i < len(datos_dc["monto_ultimo_pago"]) and len(str(valor)) > 0:
                        valor = str(valor)[:10]
                        if valor:
                            hoja.write(fila, 83, int(valor), ce.agregarEstiloAmarilloInfo(libro))
                    else:
                        hoja.write(fila, 83, '', ce.agregarEstiloAmarilloInfo(libro))


                elif variable == "fecha_ingreso_carterav_dc":
                    if i < len(datos_dc["fecha_ingreso_carterav_dc"]) and len(str(valor)) > 0:
                        valor = str(datos_dc["fecha_ingreso_carterav_dc"][i])
                        fecha_ingreso_carterav_dc = str(valor[:2] + valor[3:5] + valor[6:10])
                    else:
                        valor = ''
                        fecha_ingreso_carterav_dc = ''
                    hoja.write(fila, 84, fecha_ingreso_carterav_dc, ce.agregarEstiloAmarilloInfo(libro))


                elif variable == "monto_intereses_dc":
                    if i < len(datos_dc["monto_intereses_dc"]) and len(str(valor)) > 0:
                        valor = str(valor)[:9]
                        if valor:
                            hoja.write(fila, 85, int(valor), ce.agregarEstiloAmarilloInfo(libro))
                    else:
                        hoja.write(fila, 85, '', ce.agregarEstiloAmarilloInfo(libro))

                elif variable == "forma_pago_ac_int_dc":
                    if i < len(datos_dc["forma_pago_ac_int_dc"]) and len(str(valor)) > 0:
                        valor = str(valor)[:2]
                        if valor:
                            hoja.write(fila, 86, int(valor), ce.agregarEstiloAmarilloInfo(libro))
                    else:
                        hoja.write(fila, 86, '', ce.agregarEstiloAmarilloInfo(libro))

                elif variable == "dias_vencimiento_dc":
                    if i < len(datos_dc["dias_vencimiento_dc"]) and len(str(valor)) > 0:
                        valor = str(valor)[:3]
                        if valor:
                            hoja.write(fila, 87, int(valor), ce.agregarEstiloAmarilloInfo(libro))
                    else:
                        hoja.write(fila, 87, '', ce.agregarEstiloAmarilloInfo(libro))

                elif variable == "plazo_meses_dc":
                    if i < len(datos_dc["plazo_meses_dc"]) and len(str(valor)) > 0:
                        valor = str(valor)[:3]
                        if valor:
                            hoja.write(fila, 88, int(valor), ce.agregarEstiloAzulFuerteInfo(libro))
                    else:
                        hoja.write(fila, 88, '', ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "monto_creditoOri_dc":
                    if i < len(datos_dc["monto_creditoOri_dc"]) and len(str(valor)) > 0:
                        valor = str(valor)[:10]
                        if valor:
                            hoja.write(fila, 89, int(valor), ce.agregarEstiloAzulFuerteInfo(libro))
                    else:
                        hoja.write(fila, 89, '', ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "correo_consumidor_dc":
                    if i < len(datos_dc["correo_consumidor_dc"]) and len(str(valor)) > 40:
                        valor = valor[:40]
                    hoja.write(fila, 90, valor, ce.agregarEstiloAzulCieloInfo(libro)) 

            fila += 1
        return hoja
    
    #-------------------------------Crear función del encabezado Empleo--------------------------------------------
    def llenarCeldasCifrasControl(self, datos_encabezado, datos_dp, datos_dom, datos_emp, datos_dc, datos_cc, libro, hoja):
        ce = ClaseEstilos()
        variables = ["total_saldos_act_cc", "total_saldos_venc_cc", "total_elementosNR_cc", "total_elementosDR_cc",
                     "total_elementosER_cc", "total_elementosCR_cc", "nombre_otorgante_cc", "domicilio_devolucion"]
        

        valorMaximo = self.obtenerValorMaximo(datos_encabezado, datos_dp, datos_dom, datos_emp, datos_dc, datos_cc)
        for variable, valores in datos_cc.items():
            cantidad_actual = len(valores)
            if cantidad_actual > valorMaximo:
                valorMaximo = cantidad_actual

                    
        # Recorrer los valores y escribir en las celdas correspondientes
        fila = 2
        for i in range(valorMaximo):
            for variable in variables:
                try:
                    valor = datos_cc.get(variable, [''])[i]

                    if isinstance(valor, int):
                        valor = int(valor)
                    elif isinstance(valor, str):
                        valor = str(valor)
                except IndexError:
                    valor = ''

                if variable == "total_saldos_act_cc":
                    if i < len(datos_cc["total_saldos_act_cc"]) and len(str(valor)) > 0:
                        valor = str(valor)[:14]
                        if valor:
                            hoja.write(fila, 91, int(valor), ce.agregarEstiloAzulFuerteInfo(libro))
                    else:
                        hoja.write(fila, 91, '', ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "total_saldos_venc_cc":
                    if i < len(datos_cc["total_saldos_venc_cc"]) and len(str(valor)) > 0:
                        valor = str(valor)[:14]
                        if valor:
                            hoja.write(fila, 92, int(valor), ce.agregarEstiloAzulFuerteInfo(libro))
                    else:
                        hoja.write(fila, 92, '', ce.agregarEstiloAzulFuerteInfo(libro))
               
                elif variable == "total_elementosNR_cc":
                    if i < len(datos_cc["total_elementosNR_cc"]) and len(str(valor)) > 0:
                        valor = str(valor)[:9]
                        if valor:
                            hoja.write(fila, 93, int(valor), ce.agregarEstiloAzulFuerteInfo(libro))
                    else:
                        hoja.write(fila, 93, '', ce.agregarEstiloAzulFuerteInfo(libro))
               
                elif variable == "total_elementosDR_cc":
                    if i < len(datos_cc["total_elementosDR_cc"]) and len(str(valor)) > 0:
                        valor = str(valor)[:9]
                        if valor:
                            hoja.write(fila, 94, int(valor), ce.agregarEstiloAzulFuerteInfo(libro))
                    else:
                        hoja.write(fila, 94, '', ce.agregarEstiloAzulFuerteInfo(libro))
               
                elif variable == "total_elementosER_cc":
                    if i < len(datos_cc["total_elementosER_cc"]) and len(str(valor)) > 0:
                        valor = str(valor)[:9]
                        if valor:
                            hoja.write(fila, 95, int(valor), ce.agregarEstiloAzulFuerteInfo(libro))
                    else:
                        hoja.write(fila, 95, '', ce.agregarEstiloAzulFuerteInfo(libro))
               
                elif variable == "total_elementosCR_cc":
                    if i < len(datos_cc["total_elementosCR_cc"]) and len(str(valor)) > 0:
                        valor = str(valor)[:9]
                        if valor:
                            hoja.write(fila, 96, int(valor), ce.agregarEstiloAzulFuerteInfo(libro))
                    else:
                        hoja.write(fila, 96, '', ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "nombre_otorgante_cc":
                    if i < len(datos_cc["nombre_otorgante_cc"]) and len(str(valor)) > 40:
                        valor = valor[:40]
                    hoja.write(fila, 97, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "domicilio_devolucion":
                    if i < len(datos_cc["domicilio_devolucion"]) and len(str(valor)) > 160:
                        valor = valor[:160]
                    # Dividir el texto en líneas de máximo 30 caracteres
                    lineas = textwrap.wrap(valor, width=20)
                    # Escribir las líneas en la misma celda con saltos de línea
                    hoja.write(fila, 98, "\n".join(lineas), ce.agregarEstiloAzulFuerteInfo(libro))    

            fila += 1
        return hoja

    