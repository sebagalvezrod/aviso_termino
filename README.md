# aviso_termino
Código en python para enviar correos desde excel.
Al momento de la ejecucín del código, se evalúa si la fecha de "hoy" (el día en que se ejecuta el correo) está dentro de un rango de fechas (rango_ini - rango_fin). Si está dentro del rango, envía correos a las direcciones definidas.

Se define un título de correo y mensaje tipo con campos que van siendo llenados con la data de las filas de la tabla si es que se cumple la condición de fechas.

Se acompaña una planilla excel para la prueba del código.

Columnas de la planilla:
mail_1: dirección de correo a la que llegará una alerta para gestionar los avisos de termino de contrato.
mail_2: dirección de correo de backup del ejecutivo a cargo para gestionar los avisos de término de contrato.
mail_3: dirección de correo de otra instancia de la compañía (puede ser el área de contratos, o área legal)
cliente: nombre del cliente
fin: fecha de fin de contrato
dias: días de aviso previo al término de contrato
limite: límite contractual para envío de carta por no renovación (es la resta entre "fin" y "dias").
dias rango: días previo a la fecha límite para el inicio de la gestión (define el rango_ini)
rango_ini: definición de límite inferior de rango para envío de recordatorio para gestionar carta de aviso (es la resta entre "limite" y "dias rango")
rango_fin: definición de límite superior de rango para envío de recordatorio para gestionar carta de aviso (es la suma entre "rango_ini" y "dias rango")
