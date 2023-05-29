from estilos_documentoPM import ClaseEstilos
from controller_documentopm import ControllerDocumento

#Instanciamos las diferentes funciones que usaremos
ce = ClaseEstilos()
libro = ce.crearLibro()
hoja = ce.crearHoja(libro)
cd = ControllerDocumento()

#---------------------------------------LLAMADA DE FUNCIONES DE ESTILO------------------------------------------------------
ce.agregarEstiloFondo(libro, hoja)
ce.agregarEncabezadoEncabezado(libro, hoja)
ce.agregarEncabezadoEmpresa(libro, hoja)
ce.agregarEncabezadoAccionista(libro, hoja)
ce.agregarEncabezadoCredito(libro, hoja)
ce.agregarEncabezadoDetalleCredito(libro, hoja)
ce.agregarEncabezadoAval(libro, hoja)
ce.agregarColumnasEncabezado(libro, hoja)

#--------------------------------Datos y para llenar celdas----------------------------------------------
datos_encabezado = {
    "otorgante": [1234, 4232, 123456, 4],
    "otorgante_anterior": [1324, 3, 12345],
    "nombre_otorgante": ['Nombre otorgante 1'],
    "institucion": ['Institucion 1'],
    "formato": ['A', 'BCD'],
    "fecha": ['24/05/2023555'],
    "periodo_reporta": ['10/04/2022555'],
    "version": [123, 879]
}

datos_empresa = {
    "rfc_empresa": ['wwwwwwwwwwwwjkl'],
    "curp_empresa": ['saddddddsssssdhssskkk'],
    "compania_empresa": [],
    "nombre1_empresa": ['nombre 1'],
    "nombre2_empresa": ['nombre 2', 'nombre21', 'nombre22'],
    "apellido_paterno_empresa": ['eg34567890123456789012345yyyy'],
    "apellido_materno_empresa": ['ff'],
    "nacionalidad_empresa": ['mexico'],
    "cal_cartera_empresa": ['un'],
    "clave_banxico1": [12345678901222],
    "clave_banxico2": [1234],
    "clave_banxico3": [67843],
    "direccion1_empresa": ['asdfghj'],
    "direccion2_empresa": [],
    "colonia_empresa": ['Leon 1'], 
    "deleg_mun_empresa": ['Leon', 'Toluca', 'Tijuana'],
    "ciudad_empresa": ['Monterrey'],
    "estado_empresa": ['Nuevo Leon'],
    "cp_empresa": ['37006', '76215'],
   "telefono_empresa": ['4231023847', '21307982'],
    "extension_empresa": ['54132'],
    "fax_empresa": ['9864531'],
    "tipo_cliente_empresa": [1, 2],
    "edo_extranjero_empresa": ['Si', 'No'],
    "pais_empresa": ['Filipinas', 'España'],
    "tel_movil_empresa": ['12378921'],
    "correo_empresa": ['correo@gmail.com']
}

datos_accionista = {
    "rfc_accionista": ['123456789098'],
    "curp_accionista": ['123456789098765432'],
    "compania_accionista": [],
    "nombre1_accionista": ['nombre1', 'nombre2', 'nombre3'],
    "nombre2_accionista": ['si'],
    "apellido_paterno_accionista": ['ap1', 'ap2', 'ap3'],
    "apellido_materno_accionista": ['ap1', 'ap2'],
    "porcentaje_accionista": [10, 156],
    "direccion1_accionista": ['direccion 1'],
    "direccion2_accionista": [' direccion 2'],
    "colonia_accionista": ['Valle de León'], 
    "deleg_mun_accionista": ['Leon', 'Toluca', 'Tijuana'],
    "ciudad_accionista": ['Ciudad Victoria'],
    "estado_accionista": ['GRO'],
    "cp_accionista": ['3700', '76215'],
    "telefono_accionista": ['46548645'],
    "extension_accionista": ['123', '456'],
    "fax_accionista": ['789452'],
    "tipo_cliente_accionista": [1, 456],
    "edo_extranjero_accionista": ['Edo extranjero 1', 'Edo extranjero 2'],
    "pais_accionista": ['PO', 'Brasil']
}

datos_credito = {
    "rfc_credito": ['asdfghjklapo2'],
    "experiencias_crediticias": [1, 2, 3],
    "num_contrato": [1, 2],
    "num_contrato_anterior": [5],
    "fecha_apertura": ['20/10/2012'],
    "plazo_meses": [67],
    "tipo_credito": [1],
    "saldo_inicial": [54000],
    "moneda": ['MXN', 'EUR', 'USD'],
    "num_pagos": [7, 10],
    "frecuencia_pagos": [7],
    "importe_pagos": [875],
    "fecha_ultimo_pago": ['08/10/2002'],
    "fecha_reestructura": ['05/11/2013'],
    "pago_efectivo": [],
    "fecha_liquidacion": ['23/11/2023'],
    "quita": [123, 4645],
    "dacion": [76547, 98651],
    "quebranto_castigo": [1234, 4567],
    "clave_observacion": ['Hola', 'Ejemplo'],
    "especiales": ['F'],
    "fecha_primer_cumplimiento": ['22/02/2000'],
    "saldo_insoluto": [200000],
    "credito_maximo_utilizado": [123098],
    "fecha_ingreso_cv": ['23/05/2023']
}

datos_detalle_credito = {
    "rfc_detalle_credito": ['asdf'],
    "num_contrato_detalle_credito": ["2", "3", "4"],
    "dias_vencidos_detalle_credito": [34, 45, 68],
    "cantidad_detalle_credito": [1, 2],
    "intereses_detalle_credito": [123456]
}

datos_aval = {
    "rfc_aval": ['asv', 'asd', '123'],
    "curp_aval": ['asdfghjklñzxcvbnml'],
    "compania_aval": ['uno'],
    "nombre1_aval": ['nombre1', 'nombre2', 'nombre3', 'nombre4', 'nombre5'],
    "nombre2_aval": ['nombre21'],
    "apellido_paterno_aval": ['as', 'rea'],
    "apellido_materno_aval": ['gomez', 'lira'],
    "direccion1_aval": ['asfghjklñzxcvbnmasdqwertyuiopasdfghjklñzx'],
    "direccion2_aval": ['dire2'],
    "colonia_aval": ['La india'],
    "deleg_mun_aval": ['Leon', 'CDMX', 'Tijuana'],
    "ciudad_aval": ['city1', 'city2'],
    "estado_aval": ['CHIHUAHUA'],
    "cp_aval": ['37100', '123456'],
    "telefono_aval": ['4776009669', '4778263456'],
    "extension_aval": ['123'],
    "fax_aval": ['123908'],
    "tipo_cliente_aval": [2, 3, 1],
    "edo_extranjero_aval": ['texto tan grande que supera los 40 caracteres'],
    "pais_aval": ['Brasil', 'Alemania']
}

#-----------------------Llamado de funciones----------------------------------------------------------------
cd.llenarDatosEncabezado(datos_encabezado, datos_empresa, datos_accionista, datos_credito, datos_detalle_credito, datos_aval, libro, hoja)

cd.llenarDatosEmpresa(datos_encabezado, datos_empresa, datos_accionista, datos_credito, datos_detalle_credito, datos_aval, libro, hoja)

cd.llenarDatosAccionista(datos_encabezado, datos_empresa, datos_accionista, datos_credito, datos_detalle_credito, datos_aval, libro, hoja)

cd.llenarDatosCredito(datos_encabezado, datos_empresa, datos_accionista, datos_credito, datos_detalle_credito, datos_aval, libro, hoja)

cd.llenarDatosDetalleCredito(datos_encabezado, datos_empresa, datos_accionista, datos_credito, datos_detalle_credito, datos_aval, libro, hoja)

cd.llenarDatosAval(datos_encabezado, datos_empresa, datos_accionista, datos_credito, datos_detalle_credito, datos_aval, libro, hoja)

#Se le da un tamaño automático a cada celda con este método
hoja.autofit()
#Se cierra el libro
libro.close()
