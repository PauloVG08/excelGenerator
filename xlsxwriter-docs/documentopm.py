import xlsxwriter
from datetime import datetime
from estilos_documentoPM import ClaseEstilos
from controller_documentopm import ControllerDocumento

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

#--------------------------------Datos y llamado para llenar encabezado----------------------------------------------
datos_encabezado = {
    "otorgante": [1234, 4232, 3],
    "otorgante_anterior": [1324, 3],
    "nombre_otorgante": ['Nombre otorgante 1',''],
    "institucion": ['Institucion 1', ''],
    "formato": [1,3],
    "fecha": ['24/05/2023'],
    "periodo_reporta": ['10/04/2022'],
    "version": []
}

cd.almacenarDatosEncabezado(datos_encabezado, libro, hoja)

#-----------------------Datos y llamado para llenar Empresa----------------------------------------------------------------
datos_empresa = {
    "rfc_empresa": ['wwwwwwwwwwww'],
    "curp_empresa": ['saddddddsssssdhsss'],
    "compania_empresa": [],
    "nombre1_empresa": ['nombre 1'],
    "nombre2_empresa": ['nombre 2', 'nombre21', 'nombre22'],
    "apellido_paterno_empresa": ['eg', 'ahs'],
    "apellido_materno_empresa": ['ff'],
    "nacionalidad_empresa": ['mx'],
    "cal_cartera_empresa": ['un'],
    "clave_banxico1": [],
    "clave_banxico2": [],
    "clave_banxico3": [],
    "direccion1_empresa": [],
    "direccion2_empresa": [],
    "colonia_empresa": [], 
    "deleg_mun_empresa": ['Leon', 'Toluca', 'Tijuana'],
    "ciudad_empresa": [],
    "estado_empresa": [],
    "cp_empresa": ['37006', '76215'],
    "telefono_empresa": [],
    "extension_empresa": [],
    "fax_empresa": [],
    "tipo_cliente_empresa": [1, 2],
    "edo_extranjero_empresa": [],
    "pais_empresa": [],
    "tel_movil_empresa": [],
    "correo_empresa": ['correo@gmail.com']
}

cd.almacenarDatosEmpresa(datos_empresa, libro, hoja)

#-----------------------Datos y llamado para llenar Accionista--------------------------------------------------------------
datos_accionista = {
    "rfc_accionista": ['123456789098'],
    "curp_accionista": ['123456789098765432'],
    "compania_accionista": [],
    "nombre1_accionista": ['nombre1', 'nombre2', 'nombre3'],
    "nombre2_accionista": ['si'],
    "apellido_paterno_accionista": ['ap1', 'ap2', 'ap3'],
    "apellido_materno_accionista": ['ap1', 'ap2'],
    "porcentaje_accionista": [10, 1],
    "direccion1_accionista": [],
    "direccion2_accionista": [],
    "colonia_accionista": [], 
    "deleg_mun_accionista": ['Leon', 'Toluca', 'Tijuana'],
    "ciudad_accionista": [],
    "estado_accionista": [],
    "cp_accionista": ['3700', '76215'],
    "telefono_accionista": [],
    "extension_accionista": ['123', '456'],
    "fax_accionista": [],
    "tipo_cliente_accionista": [1],
    "edo_extranjero_accionista": [],
    "pais_accionista": ['PO']
}

cd.almacenarDatosAccionista(datos_accionista, libro, hoja)
#-----------------------Datos y llamado para llenar Crédito--------------------------------------------------------------
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
    "importe_pagos": [],
    "fecha_ultimo_pago": ['08/10/2002'],
    "fecha_reestructura": ['05/11/2013'],
    "pago_efectivo": [],
    "fecha_liquidacion": ['23/11/2023'],
    "quita": [],
    "dacion": [],
    "quebranto_castigo": [1234, 4567],
    "clave_observacion": [],
    "especiales": ['F'],
    "fecha_primer_cumplimiento": ['22/02/2000'],
    "saldo_insoluto": [],
    "credito_maximo_utilizado": [],
    "fecha_ingreso_cv": ['23/05/2023']
}

cd.almacenarDatosCredito(datos_credito, libro, hoja)
#-----------------------Datos y llamado para llenar Detalle de crédito------------------------------------------------------
datos_detalle_credito = {
    "rfc_detalle_credito": ['asdf'],
    "num_contrato_detalle_credito": [2, 3, 4],
    "dias_vencidos_detalle_credito": [34, 45, 68],
    "cantidad_detalle_credito": [1, 2],
    "intereses_detalle_credito": []
}

cd.almacenarDatosDetalleCredito(datos_detalle_credito, libro, hoja)

#Se le da un tamaño automático a cada celda con este método
hoja.autofit()
#Se cierra el libro
libro.close()
