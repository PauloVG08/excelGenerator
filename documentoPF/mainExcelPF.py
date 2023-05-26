from estilosPF import ClaseEstilos
from controllerPF import ControllerDocumento

ce = ClaseEstilos()
libro = ce.crearLibro()
hoja = ce.crearHoja(libro)
cd = ControllerDocumento()

#---------------------------------------LLAMADA DE FUNCIONES DE ESTILO------------------------------------------------------
ce.agregarEstiloFondo(libro, hoja)
ce.agregarEncabezadoEncabezado(libro, hoja)
ce.agregarEncabezadoDatoPersonales(libro, hoja)
ce.agregarEncabezadoDomicilio(libro, hoja)
ce.agregarEncabezadoEmpleo(libro, hoja)
ce.agregarEncabezadoDetalleCuenta(libro, hoja)
ce.agregarEncabezadoCifrasControl(libro, hoja)
ce.agregarColumnasEncabezado(libro, hoja)

#----------------------------Recibiendo datos y llamando funciones del controlador-------------------------------------------
#----------------------------Recibiendo datos y llamando funcion Encabezado--------------------------------------------------S
datos_encabezado = {
    "clave_otorgante": [13,2,3],
    "nombre_otorgante": ['hola', 'si'],
    "identificador_medio": ['1qds'],
    "fecha_extraccion": ['28/04/200366'],
    "nota_otorgante": [],
    "version": [13, 4]
}

cd.llenarCeldasEncabezado(datos_encabezado, libro, hoja)
#-------------------------Recibiendo datos y llamando función Datos Personales------------------------------------------------
datos_dp = {
    "apellido_paterno_dp": ['1234567890111', 'si', 'a veces'],
    "apellido_materno_dp": ['niña', 'corona'],
    "apellido_adicional_dp": ['adicional'],
    "nombres_dp": ['Luz', 'Juan'],
    "fecha_nacimiento_dp": ['01/01/2003555'],
    "rfc_dp": ['dsa'],
    "curp_dp": ['asdj', 'asdas'],
    "numero_seguridad_social_dp": [1233, 456],
    "nacionalidad": ['MX'],
    "residencia": [5],
    "numero_licencia_conducir_dp": ['Hola'],
    "estado_civil_dp": ['D'],
    "sexo_dp": ['M'],
    "clave_elector_ife_dp": ['clave1', 'clave2'],
    "numero_dependientes_dp": [12],
    "fecha_defuncion_dp": ['30/06/1976'],
    "indicador_defuncion_dp": ['F'],
    "tipo_person_dp": ['ME', 'MX', 'LA'],
}

cd.llenarCeldasDatosPersonales(datos_dp, libro, hoja)

#-------------------------Recibiendo datos y llamando función Domicilio------------------------------------------------
datos_dom = {
    "direccion_dom": ['direccion1', 'direccion2'],
    "colonia_poblacion_dom": [],
    "deleg_muni_dom": [],
    "ciudad_dom": ['Leon', 'CDMX', 'Celaya'],
    "estado_dom": ['Guanajuanto', 'Edo Mex', 'Guanajuanto'],
    "cp_dom": [12341234444,4567],
    "fecha_residencia_dom": ['13/12/2017'],
    "telefono_dom": [1234, 45678, 4776009669],
    "tipo_dom": ['f'],
    "tipo_asentamiento": [123456666, 21],
    "origen_dom": ['em', 'QL'],
}

cd.llenarCeldasDomicilio(datos_dom, libro, hoja)

#-------------------------Recibiendo datos y llamando función Empleo------------------------------------------------
datos_emp = {
    "nombre_empresa_emp": ['name 1', 'name2', 'hola'],
    "direccion_emp": ['dic1'],
    "colonia_poblacional_emp": ['Echeveste'],
    "deleg_mun": ['Neza', 'Silao'],
    "ciudad_emp": ['Cancún', 'Zamora'],
    "estado_emp": ['BCS', 'EDM'],
    "cp_emp": [123, 123456],
    "num_telefono_emp": [4776009669, 123],
    "extension_emp": [123, 12345678],
    "fax_emp": [12345678901, 123],
    "puesto_emp": ['Si', 'Guardia'],
    "fecha_contratacion_emp": ['25/12/2018', '11/06/2017'],
    "clave_moneda_emp": ["MX", "Brasil"],
    "salario_mensual_emp": [12345, 1234567890],
    "fecha_ultimo_dia_emp": ['20/11/2005'],
    "fecha_verificacion_emp": ['20/11/2005'],
    "origen_razon_social_emp": ['dos', 'tres'],
}

cd.llenarCeldasEmpleo(datos_emp, libro, hoja)

#-------------------------Recibiendo datos y llamando función Datos cuenta---------------------------------------------
datos_dc = {
    "clave_actual_dc": [12345678901111, 123],
    "nombre_otorgante_dc": ['name1', 'name2', "name3"],
    "cuenta_actual_dc": ['abc123', 'cde456'],
    "tipo_responsabilidad_dc": ['ab'],
    "tipo_cuenta_dc": ['jkl'],
    "tipo_contrato_dc": ['zxy'],
    "clave_uni_monetaria_dc": ['jkl', 'frt'],
    "valor_activo_dc": [123456789000, 1234],
    "num_pagos_dc": [12345, 123],
    "frecuencia_pagos_dc": ['ABC', "DEF"],
    "monto_pagar_dc": [1234, 567, 8910],
    "fecha_apertura_dc": ["05/11/2023"],
    "fecha_ultimo_pago_dc": ["18/03/2024"],
    "fecha_utlima_compra_dc": ["29/09/2022", "22/11/2022"],
    "fecha_cierre_cuenta_dc": ["07/08/2023", "03/07/2023"],
    "fecha_corte_dc": ["05/11/2023"],
    "garantia_dc": ["ABCDEFG12345kkkkkkkkkkkkkkkkkkkAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaaaaaaaaaaaaaaaaaaaaaa"],
    "credito_maximo_dc": [1234, 5678, 123456789000],
    "saldo_actual_dc": [1234567890111],
    "limite_credito_dc": [123456789000],
    "saldo_vencido_dc": [123456789000],
    "num_pagos_vencidos_dc": [12345555, 98765444],
    "pago_actual_dc": ["ABC", "DEF"],
    "historico_pagos_dc": ["AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"],
    "clave_prevencion_dc": ["AMA", "TKM"],
    "total_pagos_rep_dc": [123, 4],
    "clave_anterior_otor_dc": [1234567890111],
    "nombre_anterior_otor_dc": ["nombre1", "nombre2", "nombre3"],
    "num_cuenta_anterior_dc": [1234567890111],
    "fecha_primer_incump_dc": ['27/04/2006'],
    "saldo_insoluto_dc": [1234567890111, 123],
    "monto_ultimo_pago": [123456789011, 12345],
    "fecha_ingreso_carterav_dc": ['20/12/2009', '16/01/2025'],
    "monto_intereses_dc": [123456789000],
    "forma_pago_ac_int_dc": [987],
    "dias_vencimiento_dc": [123, 456, 8],
    "plazo_meses_dc": [12345666, ''],
    "monto_creditoOri_dc": [12345678990111],
    "correo_consumidor_dc": ["correo@gmail.com"],
}

cd.llenarCeldasDC(datos_dc, libro, hoja)

#-------------------------Recibiendo datos y llamando función Cifras Control---------------------------------------------
datos_cc = {
    "total_saldos_act_cc": [8465498615, 6545254312065120],
    "total_saldos_venc_cc": [123456789123456666],
    "total_elementosNR_cc": [12345678900],
    "total_elementosDR_cc": [987654321000],
    "total_elementosER_cc": [123456789000],
    "total_elementosCR_cc": [1234567890000],
    "nombre_otorgante_cc": ["Arturo", "Arturito"],
    "domicilio_devolucion": ["Mi domicilio muy largo para comprobar el salto de linea"],
}
cd.llenarCeldasCifrasControl(datos_cc, libro, hoja)
#-----------------------------------------------------------------------------------------------------------------------
#Se le da un tamaño automático a cada celda con este método
hoja.autofit()
#Se cierra el libro
libro.close()