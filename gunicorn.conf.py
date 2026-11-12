# Configuración de gunicorn. Se carga sola al arrancar con `gunicorn app:app`
# desde esta carpeta (salvo que GUNICORN_CMD_ARGS o el comando de Render la pisen).

# Dos workers: si una subida se queda colgada, el otro sigue atendiendo.
workers = 2

# Segundos sin respuesta antes de matar un worker. Cubre la conversión de un
# lote normal (cada unidad tarda unos segundos) y libera antes un worker atascado
# esperando un cuerpo de petición que no llega (antes: 10 minutos).
timeout = 180

# Una línea por petición en el log de Render: estado, tamaño de la respuesta y duración.
accesslog = "-"
access_log_format = '%(h)s "%(r)s" %(s)s %(B)s bytes %(L)s s'
