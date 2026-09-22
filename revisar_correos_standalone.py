import sys
sys.path.insert(0, '.')
import bot

# Reusa exactamente la misma logica de /revisar-correos (consulta a Correo
# Argentino, actualizacion en Firestore, aviso por WhatsApp de novedades) via
# el test_client de Flask, para no duplicar el codigo del endpoint. Corre
# desde GitHub Actions en vez de Render porque Correo Argentino bloquea las
# consultas que llegan desde la IP del servidor de Render.
with bot.app.test_client() as client:
    resp = client.get('/revisar-correos')
    print("Status:", resp.status_code)
    print("Body:", resp.get_json())
    if resp.status_code != 200:
        sys.exit(1)
