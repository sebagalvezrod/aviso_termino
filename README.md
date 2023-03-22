# aviso_termino
# Sencillo código para automatizar envio de correos cumpliendo un criterio (fechas) desde listado excel.
# Se acompaña una planilla excel (puedes usarla para testear). 
# Pensado para la gestión de cartas de aviso por no renovación de contratos, pero el template sirve para otras tareas.
# En la fecha de ejecución del código, ser revisa si "hoy" está dentro de un rango definido. 
# Este rango lo puedes definir tú modifcando la planilla que se adjunta como ejemplo. 
# En el código, debes reemplazar la ruta con la información del archivo excel (el "donde" guardas la planilla).
# Llena los campos de dirección de correos con direcciones válidas. Te sugiero usar tu dirección personal para los test.
# El código crea una instancia de outlook para el envío de correos desde la dirección de outlook que tienes activa en tu equipo.
