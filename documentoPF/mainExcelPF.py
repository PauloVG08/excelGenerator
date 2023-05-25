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

#Se le da un tamaño automático a cada celda con este método
hoja.autofit()
#Se cierra el libro
libro.close()