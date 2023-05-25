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
    "clave_otorgante": [1,2,3],
    "nombre_otorgante": ['hola', 'si'],
    "identificador_medio": ['1qds'],
    "fecha_extraccion": ['28/04/200366'],
    "nota_otorgante": [],
    "version": [1, 23]
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
    "numero_seguridad_social_dp": [123, 456],
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

#Se le da un tamaño automático a cada celda con este método
hoja.autofit()
#Se cierra el libro
libro.close()