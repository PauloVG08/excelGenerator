from estilosPF import ClaseEstilos
#from controller_documentopm import ControllerDocumento

ce = ClaseEstilos()
libro = ce.crearLibro()
hoja = ce.crearHoja(libro)
#cd = ControllerDocumento()

#---------------------------------------LLAMADA DE FUNCIONES DE ESTILO------------------------------------------------------
ce.agregarEstiloFondo(libro, hoja)
ce.agregarEncabezadoEncabezado(libro, hoja)
ce.agregarEncabezadoDatoPersonales(libro, hoja)
ce.agregarEncabezadoDomicilio(libro, hoja)
ce.agregarEncabezadoEmpleo(libro, hoja)
ce.agregarEncabezadoDetalleCuenta(libro, hoja)
ce.agregarEncabezadoCifrasControl(libro, hoja)
ce.agregarColumnasEncabezado(libro, hoja)


#Se le da un tamaño automático a cada celda con este método
hoja.autofit()
#Se cierra el libro
libro.close()