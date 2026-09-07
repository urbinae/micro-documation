#!/bin/bash
set -e

# Levanta LibreOffice como un proceso "servidor" escuchando en el socket UNO.
# Se mantiene vivo entre requests para evitar el costo de arranque (varios
# segundos) en cada conversión.
soffice --headless --invisible --nocrashreport --nodefault --nologo \
  --nofirststartwizard --norestore \
  --accept="socket,host=0.0.0.0,port=2002;urp;" &

# Esperar a que el socket UNO esté disponible antes de aceptar tráfico.
for i in $(seq 1 30); do
  if python3 -c "
import uno
ctx = uno.getComponentContext()
resolver = ctx.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', ctx)
resolver.resolve('uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext')
" 2>/dev/null; then
    echo "LibreOffice listo."
    break
  fi
  echo "Esperando a LibreOffice... ($i/30)"
  sleep 1
done

exec gunicorn -b 0.0.0.0:8080 -w 1 --timeout 180 app:app