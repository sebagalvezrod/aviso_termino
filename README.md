# aviso_termino
# Sencillo código para ayudar a automatizar envio correos desde excel.
# Está pensado para la gestión de cartas de aviso por no renovación de contratos.
# Esta gestión debe comenzar algunos días previos a una fecha límite, definida por XX días previos al fin de contrato.
# En la fecha de ejecución del código, ser revisa si "hoy" está dentro de un rango definido. 
# Este rango lo puedes definir tú modifcando la planilla que se adjunta como ejemplo. 
# Se acompaña una planilla excel. 
# En el código, debes reemplazar la ruta con la información de dónde guardes la planilla.
# Llena los campos de dirección de correos con direcciones válidas. Te sugiero usar tu dirección personal para los tests.
# El código crea una instancia de outlook para el envío de correos desde la dirección de outlook que tienes activa en tu equipo.
