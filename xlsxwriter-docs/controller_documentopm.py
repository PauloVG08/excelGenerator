import textwrap
from estilos_documentoPM import ClaseEstilos

class ControllerDocumento:
#-----------------------------------Función  para obtener el valor maximo--------------------------------------
    def obtenerValorMaximo(self, datos_encabezado, datos_empresa, datos_accionista, datos_credito, datos_detalle_credito, datos_aval):
            maximo_datos_encabezado = 0
            maximo_datos_empresa = 0
            maximo_datos_accionista = 0
            maximo_datos_credito = 0
            maximo_datos_detalle_credito = 0
            maximo_datos_aval = 0

            for valores in datos_encabezado.values():
                cantidad_actual = len(valores)
                if cantidad_actual > maximo_datos_encabezado:
                    maximo_datos_encabezado = cantidad_actual

            for valores in datos_empresa.values():
                cantidad_actual = len(valores)
                if cantidad_actual > maximo_datos_empresa:
                    maximo_datos_empresa = cantidad_actual

            for valores in datos_accionista.values():
                cantidad_actual = len(valores)
                if cantidad_actual > maximo_datos_accionista:
                    maximo_datos_accionista = cantidad_actual

            for valores in datos_credito.values():
                cantidad_actual = len(valores)
                if cantidad_actual > maximo_datos_credito:
                    maximo_datos_credito = cantidad_actual

            for valores in datos_detalle_credito.values():
                cantidad_actual = len(valores)
                if cantidad_actual > maximo_datos_detalle_credito:
                    maximo_datos_detalle_credito = cantidad_actual

            for valores in datos_aval.values():
                cantidad_actual = len(valores)
                if cantidad_actual > maximo_datos_aval:
                    maximo_datos_aval = cantidad_actual

            valorMaximo = max(maximo_datos_encabezado, maximo_datos_empresa, maximo_datos_accionista, maximo_datos_credito, 
                              maximo_datos_detalle_credito, maximo_datos_aval)
            return valorMaximo

#------------------------------------Funcion de Encabezado---------------------------------------------------
    def llenarDatosEncabezado(self, datos_encabezado, datos_empresa, datos_accionista, datos_credito, datos_detalle_credito, datos_aval, libro, hoja):
        ce = ClaseEstilos()
        variables = ["otorgante", "otorgante_anterior", "nombre_otorgante", "institucion",
                     "formato", "fecha", "periodo_reporta", "version"]

        valorMaximo = self.obtenerValorMaximo(datos_encabezado, datos_empresa, datos_accionista, datos_credito, datos_detalle_credito, datos_aval)
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
            
            #Comparación de valores----------------------------------------------------------------------------------------------
                if variable == "otorgante":
                        if i < len(datos_encabezado["otorgante"]) and len(str(valor)) > 0:
                            valor = str(valor)[:4]
                            if valor:
                                hoja.write(fila, 0, int(valor), ce.agregarEstiloAmarilloInfo(libro))
                        else:
                            hoja.write(fila, 0, '', ce.agregarEstiloAmarilloInfo(libro))

                elif variable == "otorgante_anterior":
                        if i < len(datos_encabezado["otorgante_anterior"]) and len(str(valor)) > 0:
                            valor = str(valor)[:4]
                            if valor:
                                hoja.write(fila, 1, int(valor), ce.agregarEstiloAzulClaroInfo(libro))
                        else:
                            hoja.write(fila, 1, '', ce.agregarEstiloAzulClaroInfo(libro))

                elif variable == "nombre_otorgante":
                        if i < len(datos_encabezado["nombre_otorgante"]) and len(str(valor)) > 75:
                            valor = valor[:75]
                        # Dividir el texto en líneas de máximo 20 caracteres
                        lineas = textwrap.wrap(valor, width = 20)
                        # Escribir las líneas en la misma celda con saltos de línea
                        hoja.write(fila, 2, "\n".join(lineas), ce.agregarEstiloAmarilloInfo(libro))

                elif variable == "institucion":
                        if i < len(datos_encabezado["institucion"]) and len(str(valor)) > 3:
                            valor = valor[:3]
                        hoja.write(fila, 3, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "formato":
                        if i < len(datos_encabezado["formato"]) and len(str(valor)) > 1:
                            valor = valor[:1]
                        hoja.write(fila, 4, valor, ce.agregarEstiloAzulFuerteInfo(libro))


                elif variable == "fecha":
                        if i < len(datos_encabezado["fecha"]) and len(str(valor)) > 0:
                            valor = str(datos_encabezado["fecha"][i])
                            fecha = str(valor[:2] + valor[3:5] + valor[6:10])
                        else:
                            valor = ''
                            fecha = ''
                        hoja.write(fila, 5, fecha, ce.agregarEstiloAzulFuerteInfo(libro))


                elif variable == "periodo_reporta":
                        if i < len(datos_encabezado["periodo_reporta"]) and len(str(valor)) > 0:
                            valor = str(datos_encabezado["periodo_reporta"][i])
                            # periodo_reporta = str(valor[3:5] + valor[6:10])
                            periodo_reporta = str(valor[3:10])
                            #print(periodo_reporta)
                        else:
                            valor = ''
                            periodo_reporta = ''
                        hoja.write(fila, 6, periodo_reporta, ce.agregarEstiloAzulFuerteInfo(libro))


                elif variable == "version":
                        if i < len(datos_encabezado["version"]) and len(str(valor)) > 0:
                            valor = str(valor)[:2]
                            if valor:
                                hoja.write(fila, 7, int(valor), ce.agregarEstiloAzulFuerteInfo(libro))
                        else:
                            hoja.write(fila, 7, '', ce.agregarEstiloAzulFuerteInfo(libro)) 
            fila += 1
        return hoja
    
#--------------------------Funcion Empresa--------------------------------------------------------------------------
    def llenarDatosEmpresa(self, datos_encabezado, datos_empresa, datos_accionista, datos_credito, datos_detalle_credito, datos_aval, libro, hoja):
        ce = ClaseEstilos()
        
        variables = ["rfc_empresa", "curp_empresa", "compania_empresa", "nombre1_empresa",
                     "nombre2_empresa", "apellido_paterno_empresa", "apellido_materno_empresa", "nacionalidad_empresa",
                     "cal_cartera_empresa", "clave_banxico1", "clave_banxico2", "clave_banxico3",
                     "direccion1_empresa", "direccion2_empresa", "colonia_empresa", "deleg_mun_empresa",
                     "ciudad_empresa", "estado_empresa", "cp_empresa", "telefono_empresa",
                     "extension_empresa", "fax_empresa", "tipo_cliente_empresa", "edo_extranjero_empresa",
                     "pais_empresa", "tel_movil_empresa", "correo_empresa"]

        valorMaximo = self.obtenerValorMaximo(datos_encabezado, datos_empresa, datos_accionista, datos_credito, datos_detalle_credito, datos_aval)
        for variable, valores in datos_empresa.items():
            cantidad_actual = len(valores)
            if cantidad_actual > valorMaximo:
                valorMaximo = cantidad_actual

        # Recorrer los valores y escribir en las celdas correspondientes
        fila = 2
        for i in range(valorMaximo):
            for variable in variables:
                try:
                    valor = datos_empresa.get(variable, [''])[i]

                    if isinstance(valor, int):
                        valor = int(valor)
                    elif isinstance(valor, str):
                        valor = str(valor)

                except IndexError:
                    valor = ''

                #Comparación de valores----------------------------------------------------------------------------------------------
                if variable == "rfc_empresa":
                            if i < len(datos_empresa["rfc_empresa"]) and len(str(valor)) > 13:
                                valor = valor[:13]
                            hoja.write(fila, 8, valor, ce.agregarEstiloAzulFuerteInfo(libro))
                
                elif variable == "curp_empresa":
                            if i < len(datos_empresa["curp_empresa"]) and len(str(valor)) > 18:
                                valor = valor[:18]
                            hoja.write(fila, 9, valor, ce.agregarEstiloAzulClaroInfo(libro))

                elif variable == "compania_empresa":
                        if i < len(datos_empresa["compania_empresa"]) and len(str(valor)) > 150:
                            valor = valor[:150]
                        lineas = textwrap.wrap(valor, width = 30)
                        hoja.write(fila, 10, "\n".join(lineas), ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "nombre1_empresa":
                            if i < len(datos_empresa["nombre1_empresa"]) and len(str(valor)) > 30:
                                valor = valor[:30]
                            hoja.write(fila, 11, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "nombre2_empresa":
                            if i < len(datos_empresa["nombre2_empresa"]) and len(str(valor)) > 30:
                                valor = valor[:30]
                            hoja.write(fila, 12, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "apellido_paterno_empresa":
                            if i < len(datos_empresa["apellido_paterno_empresa"]) and len(str(valor)) > 25:
                                valor = valor[:25]
                            hoja.write(fila, 13, valor, ce.agregarEstiloAzulClaroInfo(libro))

                elif variable == "apellido_materno_empresa":
                            if i < len(datos_empresa["apellido_materno_empresa"]) and len(str(valor)) > 25:
                                valor = valor[:25]
                            hoja.write(fila, 14, valor, ce.agregarEstiloAzulClaroInfo(libro))

                elif variable == "nacionalidad_empresa":
                            if i < len(datos_empresa["nacionalidad_empresa"]) and len(str(valor)) > 2:
                                valor = valor[:2]
                            hoja.write(fila, 15, valor, ce.agregarEstiloVerdeInfo(libro))

                elif variable == "cal_cartera_empresa":
                            if i < len(datos_empresa["cal_cartera_empresa"]) and len(str(valor)) > 2:
                                valor = valor[:2]
                            hoja.write(fila, 16, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "clave_banxico1":
                        if i < len(datos_empresa["clave_banxico1"]) and len(str(valor)) > 0:
                            valor = str(valor)[:11]
                            if valor:
                                hoja.write(fila, 17, int(valor), ce.agregarEstiloAzulFuerteInfo(libro))
                        else:
                            hoja.write(fila, 17, '', ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "clave_banxico2":
                        if i < len(datos_empresa["clave_banxico2"]) and len(str(valor)) > 0:
                            valor = str(valor)[:11]
                            if valor:
                                hoja.write(fila, 18, int(valor), ce.agregarEstiloAzulClaroInfo(libro))
                        else:
                            hoja.write(fila, 18, '', ce.agregarEstiloAzulClaroInfo(libro))

                elif variable == "clave_banxico3":
                        if i < len(datos_empresa["clave_banxico3"]) and len(str(valor)) > 0:
                            valor = str(valor)[:11]
                            if valor:
                                hoja.write(fila, 19, int(valor), ce.agregarEstiloAzulClaroInfo(libro))
                        else:
                            hoja.write(fila, 19, '', ce.agregarEstiloAzulClaroInfo(libro))

                elif variable == "direccion1_empresa":
                        if i < len(datos_empresa["direccion1_empresa"]) and len(str(valor)) > 40:
                            valor = valor[:40]
                        lineas = textwrap.wrap(valor, width = 20)
                        hoja.write(fila, 20, "\n".join(lineas), ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "direccion2_empresa":
                        if i < len(datos_empresa["direccion2_empresa"]) and len(str(valor)) > 40:
                            valor = valor[:40]
                        lineas = textwrap.wrap(valor, width = 20)
                        hoja.write(fila, 21, "\n".join(lineas), ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "colonia_empresa":
                        if i < len(datos_empresa["colonia_empresa"]) and len(str(valor)) > 60:
                            valor = valor[:60]
                        lineas = textwrap.wrap(valor, width = 30)
                        hoja.write(fila, 22, "\n".join(lineas), ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "deleg_mun_empresa":
                        if i < len(datos_empresa["deleg_mun_empresa"]) and len(str(valor)) > 40:
                            valor = valor[:40]
                        lineas = textwrap.wrap(valor, width = 20)
                        hoja.write(fila, 23, "\n".join(lineas), ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "ciudad_empresa":
                        if i < len(datos_empresa["ciudad_empresa"]) and len(str(valor)) > 40:
                            valor = valor[:40]
                        lineas = textwrap.wrap(valor, width = 20)
                        hoja.write(fila, 24, "\n".join(lineas), ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "estado_empresa":
                            if i < len(datos_empresa["estado_empresa"]) and len(str(valor)) > 4:
                                valor = valor[:4]
                            hoja.write(fila, 25, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "cp_empresa":
                            if i < len(datos_empresa["cp_empresa"]) and len(str(valor)) > 10:
                                valor = valor[:10]
                            hoja.write(fila, 26, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "telefono_empresa":
                            if i < len(datos_empresa["telefono_empresa"]) and len(str(valor)) > 11:
                                valor = valor[:11]
                            hoja.write(fila, 27, valor, ce.agregarEstiloAzulClaroInfo(libro))

                elif variable == "extension_empresa":
                            if i < len(datos_empresa["extension_empresa"]) and len(str(valor)) > 8:
                                valor = valor[:8]
                            hoja.write(fila, 28, valor, ce.agregarEstiloAzulClaroInfo(libro))

                elif variable == "fax_empresa":
                            if i < len(datos_empresa["fax_empresa"]) and len(str(valor)) > 11:
                                valor = valor[:11]
                            hoja.write(fila, 29, valor, ce.agregarEstiloAzulClaroInfo(libro))

                elif variable == "tipo_cliente_empresa":
                        if i < len(datos_empresa["tipo_cliente_empresa"]) and len(str(valor)) > 0:
                            valor = str(valor)[:1]
                            if valor:
                                hoja.write(fila, 30, int(valor), ce.agregarEstiloAzulFuerteInfo(libro))
                        else:
                            hoja.write(fila, 30, '', ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "edo_extranjero_empresa":
                        if i < len(datos_empresa["edo_extranjero_empresa"]) and len(str(valor)) > 40:
                            valor = valor[:40]
                        lineas = textwrap.wrap(valor, width = 20)
                        hoja.write(fila, 31, "\n".join(lineas), ce.agregarEstiloAzulClaroInfo(libro))

                elif variable == "pais_empresa":
                            if i < len(datos_empresa["pais_empresa"]) and len(str(valor)) > 2:
                                valor = valor[:2]
                            hoja.write(fila, 32, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "tel_movil_empresa":
                            if i < len(datos_empresa["tel_movil_empresa"]) and len(str(valor)) > 10:
                                valor = valor[:10]
                            hoja.write(fila, 33, valor, ce.agregarEstiloGrisInfo(libro))

                elif variable == "correo_empresa":
                            if i < len(datos_empresa["correo_empresa"]) and len(str(valor)) > 100:
                                valor = valor[:100]
                            hoja.write(fila, 34, valor, ce.agregarEstiloGrisInfo(libro))
            fila += 1
        return hoja

#--------------------------Funcion Accionista-------------------------------------------------------------------------
    def llenarDatosAccionista(self, datos_encabezado, datos_empresa, datos_accionista, datos_credito, datos_detalle_credito, datos_aval, libro, hoja):
        ce = ClaseEstilos()
        variables = ["rfc_accionista", "curp_accionista", "compania_accionista", "nombre1_accionista",
                     "nombre2_accionista", "apellido_paterno_accionista", "apellido_materno_accionista", "porcentaje_accionista",
                     "direccion1_accionista", "clave_banxico1", "clave_banxico2", "clave_banxico3",
                     "direccion1_empresa", "direccion2_accionista", "colonia_accionista", "deleg_mun_accionista",
                     "ciudad_accionista", "estado_accionista", "cp_accionista", "telefono_accionista",
                     "extension_accionista", "fax_accionista", "tipo_cliente_accionista", "edo_extranjero_accionista",
                     "pais_accionista"]

        valorMaximo = self.obtenerValorMaximo(datos_encabezado, datos_empresa, datos_accionista, datos_credito, datos_detalle_credito, datos_aval)
        for variable, valores in datos_accionista.items():
            cantidad_actual = len(valores)
            if cantidad_actual > valorMaximo:
                valorMaximo = cantidad_actual

         # Recorrer los valores y escribir en las celdas correspondientes
        fila = 2
        for i in range(valorMaximo):
            for variable in variables:
                try:
                    valor = datos_accionista.get(variable, [''])[i]

                    if isinstance(valor, int):
                        valor = int(valor)
                    elif isinstance(valor, str):
                        valor = str(valor)

                except IndexError:
                    valor = ''

                #Comparación de datos---------------------------------------------------------------------------
                if variable == "rfc_accionista":
                                if i < len(datos_accionista["rfc_accionista"]) and len(str(valor)) > 13:
                                    valor = valor[:13]
                                hoja.write(fila, 35, valor, ce.agregarEstiloAzulFuerteInfo(libro))
                   
                elif variable == "curp_accionista":
                            if i < len(datos_accionista["curp_accionista"]) and len(str(valor)) > 18:
                                valor = valor[:18]
                            hoja.write(fila, 36, valor, ce.agregarEstiloAzulClaroInfo(libro))

                elif variable == "compania_accionista":
                        if i < len(datos_accionista["compania_accionista"]) and len(str(valor)) > 150:
                            valor = valor[:150]
                        lineas = textwrap.wrap(valor, width = 30)
                        hoja.write(fila, 37, "\n".join(lineas), ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "nombre1_accionista":
                            if i < len(datos_accionista["nombre1_accionista"]) and len(str(valor)) > 30:
                                valor = valor[:30]
                            hoja.write(fila, 38, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "nombre2_accionista":
                            if i < len(datos_accionista["nombre2_accionista"]) and len(str(valor)) > 30:
                                valor = valor[:30]
                            hoja.write(fila, 39, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "apellido_paterno_accionista":
                            if i < len(datos_accionista["apellido_paterno_accionista"]) and len(str(valor)) > 25:
                                valor = valor[:25]
                            hoja.write(fila, 40, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "apellido_materno_accionista":
                            if i < len(datos_accionista["apellido_materno_accionista"]) and len(str(valor)) > 25:
                                valor = valor[:25]
                            hoja.write(fila, 41, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "porcentaje_accionista":
                        if i < len(datos_accionista["porcentaje_accionista"]) and len(str(valor)) > 0:
                            valor = str(valor)[:2]
                            if valor:
                                hoja.write(fila, 42, int(valor), ce.agregarEstiloAzulFuerteInfo(libro))
                        else:
                            hoja.write(fila, 42, '', ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "direccion1_accionista":
                        if i < len(datos_accionista["direccion1_accionista"]) and len(str(valor)) > 40:
                            valor = valor[:40]
                        lineas = textwrap.wrap(valor, width = 20)
                        hoja.write(fila, 43, "\n".join(lineas), ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "direccion2_accionista":
                        if i < len(datos_accionista["direccion2_accionista"]) and len(str(valor)) > 40:
                            valor = valor[:40]
                        lineas = textwrap.wrap(valor, width = 20)
                        hoja.write(fila, 44, "\n".join(lineas), ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "colonia_accionista":
                        if i < len(datos_accionista["colonia_accionista"]) and len(str(valor)) > 60:
                            valor = valor[:60]
                        lineas = textwrap.wrap(valor, width = 30)
                        hoja.write(fila, 45, "\n".join(lineas), ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "deleg_mun_accionista":
                    if i < len(datos_accionista["deleg_mun_accionista"]) and len(str(valor)) > 40:
                        valor = valor[:40]
                    lineas = textwrap.wrap(valor, width = 20)
                    hoja.write(fila, 46, "\n".join(lineas), ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "ciudad_accionista":
                        if i < len(datos_accionista["ciudad_accionista"]) and len(str(valor)) > 40:
                            valor = valor[:40]
                        lineas = textwrap.wrap(valor, width = 20)
                        hoja.write(fila, 47, "\n".join(lineas), ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "estado_accionista":
                            if i < len(datos_accionista["estado_accionista"]) and len(str(valor)) > 4:
                                valor = valor[:4]
                            hoja.write(fila, 48, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "cp_accionista":
                            if i < len(datos_accionista["cp_accionista"]) and len(str(valor)) > 10:
                                valor = valor[:10]
                            hoja.write(fila, 49, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "telefono_accionista":
                            if i < len(datos_accionista["telefono_accionista"]) and len(str(valor)) > 11:
                                valor = valor[:11]
                            hoja.write(fila, 50, valor, ce.agregarEstiloAzulClaroInfo(libro))

                elif variable == "extension_accionista":
                            if i < len(datos_accionista["extension_accionista"]) and len(str(valor)) > 8:
                                valor = valor[:8]
                            hoja.write(fila, 51, valor, ce.agregarEstiloAzulClaroInfo(libro))

                elif variable == "fax_accionista":
                            if i < len(datos_accionista["fax_accionista"]) and len(str(valor)) > 11:
                                valor = valor[:11]
                            hoja.write(fila, 52, valor, ce.agregarEstiloAzulClaroInfo(libro))

                elif variable == "tipo_cliente_accionista":
                        if i < len(datos_accionista["tipo_cliente_accionista"]) and len(str(valor)) > 0:
                            valor = str(valor)[:1]
                            if valor:
                                hoja.write(fila, 53, int(valor), ce.agregarEstiloAzulFuerteInfo(libro))
                        else:
                            hoja.write(fila, 53, '', ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "edo_extranjero_accionista":
                        if i < len(datos_accionista["edo_extranjero_accionista"]) and len(str(valor)) > 40:
                            valor = valor[:40]
                        lineas = textwrap.wrap(valor, width = 20)
                        hoja.write(fila, 54, "\n".join(lineas), ce.agregarEstiloAzulClaroInfo(libro))

                elif variable == "pais_accionista":
                            if i < len(datos_accionista["pais_accionista"]) and len(str(valor)) > 2:
                                valor = valor[:2]
                            hoja.write(fila, 55, valor, ce.agregarEstiloAzulClaroInfo(libro))
            fila += 1
        return hoja

#--------------------------Funcion Credito--------------------------------------------------------------------------
    def llenarDatosCredito(self, datos_encabezado, datos_empresa, datos_accionista, datos_credito, datos_detalle_credito, datos_aval, libro, hoja):
        ce = ClaseEstilos()
        
        variables = ["rfc_credito", "experiencias_crediticias", "num_contrato", "num_contrato_anterior",
                     "fecha_apertura", "plazo_meses", "tipo_credito", "saldo_inicial",
                     "moneda", "num_pagos", "frecuencia_pagos", "importe_pagos",
                     "fecha_ultimo_pago", "fecha_reestructura", "pago_efectivo", "fecha_liquidacion",
                     "quita", "dacion", "quebranto_castigo", "clave_observacion",
                     "especiales", "fecha_primer_cumplimiento", "saldo_insoluto", "credito_maximo_utilizado",
                     "fecha_ingreso_cv"]

        valorMaximo = self.obtenerValorMaximo(datos_encabezado, datos_empresa, datos_accionista, datos_credito, datos_detalle_credito, datos_aval)
        for variable, valores in datos_credito.items():
            cantidad_actual = len(valores)
            if cantidad_actual > valorMaximo:
                valorMaximo = cantidad_actual

        fila = 2
        for i in range(valorMaximo):
            for variable in variables:
                try:
                    valor = datos_credito.get(variable, [''])[i]

                    if isinstance(valor, int):
                        valor = int(valor)
                    elif isinstance(valor, str):
                        valor = str(valor)

                except IndexError:
                    valor = ''

                #Comparación de datos---------------------------------------------------------------------------
                if variable == "rfc_credito":
                                if i < len(datos_credito["rfc_credito"]) and len(str(valor)) > 13:
                                    valor = valor[:13]
                                hoja.write(fila, 56, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "experiencias_crediticias":
                            if i < len(datos_credito["experiencias_crediticias"]) and len(str(valor)) > 0:
                                valor = str(valor)[:6]
                                if valor:
                                    hoja.write(fila, 57, int(valor), ce.agregarEstiloAzulFuerteInfo(libro))
                            else:
                                hoja.write(fila, 57, '', ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "num_contrato":
                                if i < len(datos_credito["num_contrato"]) and len(str(valor)) > 25:
                                    valor = valor[:25]
                                hoja.write(fila, 58, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "num_contrato_anterior":
                                if i < len(datos_credito["num_contrato_anterior"]) and len(str(valor)) > 25:
                                    valor = valor[:25]
                                hoja.write(fila, 59, valor, ce.agregarEstiloAzulClaroInfo(libro))

                elif variable == "fecha_apertura":
                            if i < len(datos_credito["fecha_apertura"]) and len(str(valor)) > 0:
                                valor = str(datos_credito["fecha_apertura"][i])
                                fecha_apertura = int(valor[:2] + valor[3:5] + valor[6:10])
                            else:
                                valor = ''
                                fecha_apertura = ''
                            hoja.write(fila, 60, fecha_apertura, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "plazo_meses":
                            if i < len(datos_credito["plazo_meses"]) and len(str(valor)) > 0:
                                valor = str(valor)[:5]
                                if valor:
                                    hoja.write(fila, 61, int(valor), ce.agregarEstiloVerdeInfo(libro))
                            else:
                                hoja.write(fila, 61, '', ce.agregarEstiloVerdeInfo(libro))

                elif variable == "tipo_credito":
                            if i < len(datos_credito["tipo_credito"]) and len(str(valor)) > 0:
                                valor = str(valor)[:4]
                                if valor:
                                    hoja.write(fila, 62, int(valor), ce.agregarEstiloAzulFuerteInfo(libro))
                            else:
                                hoja.write(fila, 62, '', ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "saldo_inicial":
                            if i < len(datos_credito["saldo_inicial"]) and len(str(valor)) > 0:
                                valor = str(valor)[:20]
                                if valor:
                                    hoja.write(fila, 63, int(valor), ce.agregarEstiloAzulFuerteInfo(libro))
                            else:
                                hoja.write(fila, 63, '', ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "moneda":
                                if i < len(datos_credito["moneda"]) and len(str(valor)) > 3:
                                    valor = valor[:3]
                                hoja.write(fila, 64, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "num_pagos":
                            if i < len(datos_credito["num_pagos"]) and len(str(valor)) > 0:
                                valor = str(valor)[:4]
                                if valor:
                                    hoja.write(fila, 65, int(valor), ce.agregarEstiloAzulFuerteInfo(libro))
                            else:
                                hoja.write(fila, 65, '', ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "frecuencia_pagos":
                            if i < len(datos_credito["frecuencia_pagos"]) and len(str(valor)) > 0:
                                valor = str(valor)[:5]
                                if valor:
                                    hoja.write(fila, 66, int(valor), ce.agregarEstiloVerdeInfo(libro))
                            else:
                                hoja.write(fila, 66, '', ce.agregarEstiloVerdeInfo(libro))

                elif variable == "importe_pagos":
                            if i < len(datos_credito["importe_pagos"]) and len(str(valor)) > 0:
                                valor = str(valor)[:20]
                                if valor:
                                    hoja.write(fila, 67, int(valor), ce.agregarEstiloVerdeInfo(libro))
                            else:
                                hoja.write(fila, 67, '', ce.agregarEstiloVerdeInfo(libro))

                elif variable == "fecha_ultimo_pago":
                            if i < len(datos_credito["fecha_ultimo_pago"]) and len(str(valor)) > 0:
                                valor = str(datos_credito["fecha_ultimo_pago"][i])
                                fecha_ultimo_pago = str(valor[:2] + valor[3:5] + valor[6:10])
                            else:
                                valor = ''
                                fecha_ultimo_pago = ''
                            hoja.write(fila, 68, fecha_ultimo_pago, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "fecha_reestructura":
                            if i < len(datos_credito["fecha_reestructura"]) and len(str(valor)) > 0:
                                valor = str(datos_credito["fecha_reestructura"][i])
                                fecha_reestructura = str(valor[:2] + valor[3:5] + valor[6:10])
                            else:
                                valor = ''
                                fecha_reestructura = ''
                            hoja.write(fila, 69, fecha_reestructura, ce.agregarEstiloAzulClaroInfo(libro))

                elif variable == "pago_efectivo":
                            if i < len(datos_credito["pago_efectivo"]) and len(str(valor)) > 0:
                                valor = str(valor)[:20]
                                if valor:
                                    hoja.write(fila, 70, int(valor), ce.agregarEstiloAzulClaroInfo(libro))
                            else:
                                hoja.write(fila, 70, '', ce.agregarEstiloAzulClaroInfo(libro))

                elif variable == "fecha_liquidacion":
                            if i < len(datos_credito["fecha_liquidacion"]) and len(str(valor)) > 0:
                                valor = str(datos_credito["fecha_liquidacion"][i])
                                fecha_liquidacion = str(valor[:2] + valor[3:5] + valor[6:10])
                            else:
                                valor = ''
                                fecha_liquidacion = ''
                            hoja.write(fila, 71, fecha_liquidacion, ce.agregarEstiloAzulClaroInfo(libro))

                elif variable == "quita":
                            if i < len(datos_credito["quita"]) and len(str(valor)) > 0:
                                valor = str(valor)[:20]
                                if valor:
                                    hoja.write(fila, 72, int(valor), ce.agregarEstiloAzulClaroInfo(libro))
                            else:
                                hoja.write(fila, 72, '', ce.agregarEstiloAzulClaroInfo(libro))

                elif variable == "dacion":
                            if i < len(datos_credito["dacion"]) and len(str(valor)) > 0:
                                valor = str(valor)[:20]
                                if valor:
                                    hoja.write(fila, 73, int(valor), ce.agregarEstiloAzulClaroInfo(libro))
                            else:
                                hoja.write(fila, 73, '', ce.agregarEstiloAzulClaroInfo(libro))

                elif variable == "quebranto_castigo":
                            if i < len(datos_credito["quebranto_castigo"]) and len(str(valor)) > 0:
                                valor = str(valor)[:20]
                                if valor:
                                    hoja.write(fila, 74, int(valor), ce.agregarEstiloAzulClaroInfo(libro))
                            else:
                                hoja.write(fila, 74, '', ce.agregarEstiloAzulClaroInfo(libro))

                elif variable == "clave_observacion":
                                if i < len(datos_credito["clave_observacion"]) and len(str(valor)) > 4:
                                    valor = valor[:4]
                                hoja.write(fila, 75, valor, ce.agregarEstiloAzulClaroInfo(libro))

                elif variable == "especiales":
                                if i < len(datos_credito["especiales"]) and len(str(valor)) > 1:
                                    valor = valor[:1]
                                hoja.write(fila, 76, valor, ce.agregarEstiloAzulClaroInfo(libro))

                elif variable == "fecha_primer_cumplimiento":
                            if i < len(datos_credito["fecha_primer_cumplimiento"]) and len(str(valor)) > 0:
                                valor = str(datos_credito["fecha_primer_cumplimiento"][i])
                                fecha_primer_cumplimiento = str(valor[:2] + valor[3:5] + valor[6:10])
                            else:
                                valor = ''
                                fecha_primer_cumplimiento = ''
                            hoja.write(fila, 77, fecha_primer_cumplimiento, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "saldo_insoluto":
                            if i < len(datos_credito["saldo_insoluto"]) and len(str(valor)) > 0:
                                valor = str(valor)[:8]
                                if valor:
                                    hoja.write(fila, 78, int(valor), ce.agregarEstiloAzulFuerteInfo(libro))
                            else:
                                hoja.write(fila, 78, '', ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "credito_maximo_utilizado":
                            if i < len(datos_credito["credito_maximo_utilizado"]) and len(str(valor)) > 0:
                                valor = str(valor)[:20]
                                if valor:
                                    hoja.write(fila, 79, int(valor), ce.agregarEstiloVerdeInfo(libro))
                            else:
                                hoja.write(fila, 79, '', ce.agregarEstiloVerdeInfo(libro))

                elif variable == "fecha_ingreso_cv":
                            if i < len(datos_credito["fecha_ingreso_cv"]) and len(str(valor)) > 0:
                                valor = str(datos_credito["fecha_ingreso_cv"][i])
                                fecha_ingreso_cv = str(valor[:2] + valor[3:5] + valor[6:10])
                            else:
                                valor = ''
                                fecha_ingreso_cv = ''
                            hoja.write(fila, 80, fecha_ingreso_cv, ce.agregarEstiloGrisInfo(libro))
            fila += 1
        return hoja

#--------------------------Funcion Detalle de crédito----------------------------------------------------------------
    def llenarDatosDetalleCredito(self, datos_encabezado, datos_empresa, datos_accionista, datos_credito, datos_detalle_credito, datos_aval, libro, hoja):
        ce = ClaseEstilos()
        
        variables = ["rfc_detalle_credito", "num_contrato_detalle_credito", "dias_vencidos_detalle_credito",
                     "cantidad_detalle_credito", "intereses_detalle_credito"]

        valorMaximo = self.obtenerValorMaximo(datos_encabezado, datos_empresa, datos_accionista, datos_credito, datos_detalle_credito, datos_aval)
        for variable, valores in datos_detalle_credito.items():
            cantidad_actual = len(valores)
            if cantidad_actual > valorMaximo:
                valorMaximo = cantidad_actual

        # Recorrer los valores y escribir en las celdas correspondientes
        fila = 2
        for i in range(valorMaximo):
            for variable in variables:
                try:
                    valor = datos_detalle_credito.get(variable, [''])[i]

                    if isinstance(valor, int):
                        valor = int(valor)
                    elif isinstance(valor, str):
                        valor = str(valor)

                except IndexError:
                    valor = ''

            #Comparación de valores----------------------------------------------------------------------------------------------

                if variable == "rfc_detalle_credito":
                    if i < len(datos_detalle_credito["rfc_detalle_credito"]) and len(str(valor)) > 13:
                        valor = valor[:13]
                    hoja.write(fila, 81, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "num_contrato_detalle_credito":
                    if i < len(datos_detalle_credito["num_contrato_detalle_credito"]) and len(str(valor)) > 25:
                        valor = valor[:25]
                    hoja.write(fila, 82, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "dias_vencidos_detalle_credito":
                        if i < len(datos_detalle_credito["dias_vencidos_detalle_credito"]) and len(str(valor)) > 0:
                            valor = str(valor)[:2]
                            if valor:
                                hoja.write(fila, 83, int(valor), ce.agregarEstiloAzulFuerteInfo(libro))
                        else:
                            hoja.write(fila, 83, '', ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "cantidad_detalle_credito":
                        if i < len(datos_detalle_credito["cantidad_detalle_credito"]) and len(str(valor)) > 0:
                            valor = str(valor)[:20]
                            if valor:
                                hoja.write(fila, 84, int(valor), ce.agregarEstiloAzulFuerteInfo(libro))
                        else:
                            hoja.write(fila, 84, '', ce.agregarEstiloAzulFuerteInfo(libro))
                
                elif variable == "intereses_detalle_credito":
                        if i < len(datos_detalle_credito["intereses_detalle_credito"]) and len(str(valor)) > 0:
                            valor = str(valor)[:20]
                            if valor:
                                hoja.write(fila, 85, int(valor), ce.agregarEstiloAzulFuerteInfo(libro))
                        else:
                            hoja.write(fila, 85, '', ce.agregarEstiloAzulFuerteInfo(libro))
            fila += 1
        return hoja

#--------------------------Funcion Aval----------------------------------------------------------------
    def llenarDatosAval(self, datos_encabezado, datos_empresa, datos_accionista, datos_credito, datos_detalle_credito, datos_aval, libro, hoja):
        ce = ClaseEstilos()
        variables = ["rfc_aval", "curp_aval", "compania_aval", "nombre1_aval",
                     "nombre2_aval", "apellido_paterno_aval", "apellido_materno_aval", "direccion1_aval", 
                     "direccion2_aval", "colonia_aval", "deleg_mun_aval", "ciudad_aval", 
                     "estado_aval", "cp_aval", "telefono_aval", "extension_aval",
                     "fax_aval", "tipo_cliente_aval", "edo_extranjero_aval", "pais_aval"]

        valorMaximo = self.obtenerValorMaximo(datos_encabezado, datos_empresa, datos_accionista, datos_credito, datos_detalle_credito, datos_aval)
        for variable, valores in datos_aval.items():
            cantidad_actual = len(valores)
            if cantidad_actual > valorMaximo:
                valorMaximo = cantidad_actual

        # Recorrer los valores y escribir en las celdas correspondientes
        fila = 2
        for i in range(valorMaximo):
            for variable in variables:
                try:
                    valor = datos_aval.get(variable, [''])[i]

                    if isinstance(valor, int):
                        valor = int(valor)
                    elif isinstance(valor, str):
                        valor = str(valor)

                except IndexError:
                    valor = ''

                if variable == "rfc_aval":
                            if i < len(datos_aval["rfc_aval"]) and len(str(valor)) > 13:
                                valor = valor[:13]
                            hoja.write(fila, 86, valor, ce.agregarEstiloAzulFuerteInfo(libro))
    
                elif variable == "curp_aval":
                            if i < len(datos_aval["curp_aval"]) and len(str(valor)) > 18:
                                valor = valor[:18]
                            hoja.write(fila, 87, valor, ce.agregarEstiloAzulClaroInfo(libro))

                elif variable == "compania_aval":
                    if i < len(datos_aval["compania_aval"]) and len(str(valor)) > 150:
                        valor = valor[:150]
                    lineas = textwrap.wrap(valor, width = 30)
                    hoja.write(fila, 88, "\n".join(lineas), ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "nombre1_aval":
                        if i < len(datos_aval["nombre1_aval"]) and len(str(valor)) > 30:
                            valor = valor[:30]
                        hoja.write(fila, 89, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "nombre2_aval":
                            if i < len(datos_aval["nombre2_aval"]) and len(str(valor)) > 30:
                                valor = valor[:30]
                            hoja.write(fila, 90, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "apellido_paterno_aval":
                            if i < len(datos_aval["apellido_paterno_aval"]) and len(str(valor)) > 25:
                                valor = valor[:25]
                            hoja.write(fila, 91, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "apellido_materno_aval":
                            if i < len(datos_aval["apellido_materno_aval"]) and len(str(valor)) > 25:
                                valor = valor[:25]
                            hoja.write(fila, 92, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "direccion1_aval":
                    if i < len(datos_aval["direccion1_aval"]) and len(str(valor)) > 40:
                        valor = valor[:40]
                    lineas = textwrap.wrap(valor, width = 20)
                    hoja.write(fila, 93, "\n".join(lineas), ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "direccion2_aval":
                        if i < len(datos_aval["direccion2_aval"]) and len(str(valor)) > 40:
                            valor = valor[:40]
                        lineas = textwrap.wrap(valor, width = 20)
                        hoja.write(fila, 94, "\n".join(lineas), ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "colonia_aval":
                        if i < len(datos_aval["colonia_aval"]) and len(str(valor)) > 60:
                            valor = valor[:60]
                        lineas = textwrap.wrap(valor, width = 30)
                        hoja.write(fila, 95, "\n".join(lineas), ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "deleg_mun_aval":
                    if i < len(datos_aval["deleg_mun_aval"]) and len(str(valor)) > 40:
                        valor = valor[:40]
                    lineas = textwrap.wrap(valor, width = 20)
                    hoja.write(fila, 96, "\n".join(lineas), ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "ciudad_aval":
                        if i < len(datos_aval["ciudad_aval"]) and len(str(valor)) > 40:
                            valor = valor[:40]
                        lineas = textwrap.wrap(valor, width = 20)
                        hoja.write(fila, 97, "\n".join(lineas), ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "estado_aval":
                            if i < len(datos_aval["estado_aval"]) and len(str(valor)) > 4:
                                valor = valor[:4]
                            hoja.write(fila, 98, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "cp_aval":
                            if i < len(datos_aval["cp_aval"]) and len(str(valor)) > 10:
                                valor = valor[:10]
                            hoja.write(fila, 99, valor, ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "telefono_aval":
                        if i < len(datos_aval["telefono_aval"]) and len(str(valor)) > 11:
                            valor = valor[:11]
                        hoja.write(fila, 100, valor, ce.agregarEstiloAzulClaroInfo(libro))

                elif variable == "extension_aval":
                            if i < len(datos_aval["extension_aval"]) and len(str(valor)) > 8:
                                valor = valor[:8]
                            hoja.write(fila, 101, valor, ce.agregarEstiloAzulClaroInfo(libro))

                elif variable == "fax_aval":
                            if i < len(datos_aval["fax_aval"]) and len(str(valor)) > 11:
                                valor = valor[:11]
                            hoja.write(fila, 102, valor, ce.agregarEstiloAzulClaroInfo(libro))

                elif variable == "tipo_cliente_aval":
                        if i < len(datos_aval["tipo_cliente_aval"]) and len(str(valor)) > 0:
                            valor = str(valor)[:1]
                            if valor:
                                hoja.write(fila, 103, int(valor), ce.agregarEstiloAzulFuerteInfo(libro))
                        else:
                            hoja.write(fila, 103, '', ce.agregarEstiloAzulFuerteInfo(libro))

                elif variable == "edo_extranjero_aval":
                        if i < len(datos_aval["edo_extranjero_aval"]) and len(str(valor)) > 40:
                            valor = valor[:40]
                        lineas = textwrap.wrap(valor, width = 20)
                        hoja.write(fila, 104, "\n".join(lineas), ce.agregarEstiloAzulClaroInfo(libro))

                elif variable == "pais_aval":
                            if i < len(datos_aval["pais_aval"]) and len(str(valor)) > 2:
                                valor = valor[:2]
                            hoja.write(fila, 105, valor, ce.agregarEstiloAzulClaroInfo(libro))
            fila += 1
        return hoja

