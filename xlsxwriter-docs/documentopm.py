import xlsxwriter
from datetime import datetime
from estilos_documentoPM import ClaseEstilos
from controller_documentopm import ControllerDocumento

ce = ClaseEstilos()
libro = ce.crearLibro()
hoja = ce.crearHoja(libro)
cd = ControllerDocumento()

#--------------------------------Datos y llamado para llenar encabezado----------------------------------------------
datos_encabezado = {
    "otorgante": [1234, 4232, 3],
    "otorgante_anterior": [1324, 3],
    "nombre_otorgante": ['Nombre otorgante 1',''],
    "institucion": ['Institucion 1', ''],
    "formato": [1,3],
    "fecha": ['24/05/2023'],
    "periodo_reporta": ['10/04/2022'],
    "version": [12, 13]
}

cd.almacenarDatosEncabezado(datos_encabezado, libro, hoja)

#-----------------------Datos y llamado para llenar Empresa----------------------------------------------------------------
datos_empresa = {
    "rfc_empresa": [],
    "curp_empresa": [],
    "compania_empresa": [],
    "nombre1_empresa": [],
    "nombre2_empresa": [],
    "apellido_paterno_empresa": [],
    "apellido_materno_empresa": [],
    "nacionalidad_empresa": [],
    "cal_cartera_empresa": [],
    "clave_banxico1": [],
    "clave_banxico2": [],
    "clave_banxico3": [],
    "direccion1_empresa": [],
    "direccion2_empresa": [],
    "colonia_empresa": [], 
    "deleg_mun_empresa": [],
    "ciudad_empresa": [],
    "estado_empresa": [],
    "cp_empresa": [],
    "telefono_empresa": [],
    "extension_empresa": [],
    "fax_empresa": [],
    "tipo_cliente_empresa": [],
    "edo_extranjero_empresa": [],
    "pais_empresa": [],
    "tel_movil_empresa": [],
    "correo_empresa": []
}

#---------------------------------------LLAMADA DE FUNCIONES DE ESTILO------------------------------------------------------
ce.agregarEstiloFondo(libro, hoja)
ce.agregarEncabezadoEncabezado(libro, hoja)
ce.agregarEncabezadoEmpresa(libro, hoja)
ce.agregarEncabezadoAccionista(libro, hoja)
ce.agregarEncabezadoCredito(libro, hoja)
ce.agregarEncabezadoDetalleCredito(libro, hoja)
ce.agregarEncabezadoAval(libro, hoja)
ce.agregarColumnasEncabezado(libro, hoja)

#Se le da un tamaño automático a cada celda con este método
hoja.autofit()
#Se cierra el libro
libro.close()
