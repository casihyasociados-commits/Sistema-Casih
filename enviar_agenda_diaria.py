import os
import sys
import json
import time
import requests
from datetime import datetime
import pytz
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

def obtener_eventos_hoy():
    creds_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not creds_json:
        raise ValueError("No se encontró GOOGLE_SERVICE_ACCOUNT_JSON")
    
    info = json.loads(creds_json)
    scopes = ['https://www.googleapis.com/auth/calendar.readonly']
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    
    service = build('calendar', 'v3', credentials=creds)
    
    tz = pytz.timezone('America/Argentina/Buenos_Aires')
    ahora = datetime.now(tz)
    inicio_dia = ahora.replace(hour=0, minute=0, second=0, microsecond=0).isoformat()
    fin_dia = ahora.replace(hour=23, minute=59, second=59, microsecond=0).isoformat()
    
    events_result = service.events().list(
        calendarId='casihyasociados@gmail.com',
        timeMin=inicio_dia,
        timeMax=fin_dia,
        singleEvents=True,
        orderBy='startTime'
    ).execute()
    
    return events_result.get('items', []), ahora.strftime("%d/%m/%Y")

def get_render(path):
    ultimo=None
    for intento in range(3):
        try:
            r = requests.get("https://sistema-casih.onrender.com" + path, timeout=60)
            r.raise_for_status()
            return r.json()
        except Exception as e:
            ultimo = e
            print("Intento %d fallido en %s: %s" % (intento + 1, path, str(e)))
            time.sleep(5)
    raise ultimo

def obtener_recordatorios_hoy():
    try:
        return get_render("/recordatorios-del-dia")
    except Exception as e:
        print("No se pudieron obtener los recordatorios del día:", str(e))
        return []

# A diferencia de los recordatorios, las tareas pendientes no tienen fecha:
# se avisan TODOS los días hasta que alguien las marca como hechas desde el
# sistema (ahí se borran y dejan de aparecer acá).
def obtener_tareas_pendientes():
    try:
        return get_render("/tareas-pendientes")
    except Exception as e:
        print("No se pudieron obtener las tareas pendientes:", str(e))
        return []

def obtener_pasos_seguimiento_hoy():
    try:
        return get_render("/seguimiento-pasos-del-dia")
    except Exception as e:
        print("No se pudieron obtener los pasos de seguimiento:", str(e))
        return []

def revisar_seguimientos_tareas():
    try:
        r = requests.get("https://sistema-casih.onrender.com/revisar-seguimientos-tareas", timeout=15)
        r.raise_for_status()
        print("Seguimientos de tareas revisados:", r.json())
    except Exception as e:
        print("No se pudo revisar seguimientos de tareas:", str(e))

def enviar_whatsapp(mensaje):
    instance_id = os.environ.get("GREEN_API_INSTANCE_ID")
    token = os.environ.get("GREEN_API_TOKEN")
    group_id = os.environ.get("GREEN_API_GROUP_ID")
    
    url = f"https://api.green-api.com/waInstance{instance_id}/sendMessage/{token}"
    
    payload = {
        "chatId": group_id,
        "message": mensaje
    }
    
    headers = {'Content-Type': 'application/json'}
    response = requests.post(url, json=payload, headers=headers, timeout=30)
    print("Green API respondió:", response.status_code, response.text[:300])
    response.raise_for_status()
    return response.json()

def diagnosticar_envio(res):
    # Green API acepta el mensaje (200 + idMessage) aunque despues no se entregue:
    # consultamos el estado de la instancia y el estado del mensaje para saber por que.
    instance_id = os.environ.get("GREEN_API_INSTANCE_ID")
    token = os.environ.get("GREEN_API_TOKEN")
    group_id = os.environ.get("GREEN_API_GROUP_ID") or ""
    base = f"https://api.green-api.com/waInstance{instance_id}"
    instancia_ok = True
    try:
        estado = requests.get(f"{base}/getStateInstance/{token}", timeout=20)
        print("Estado de la instancia:", estado.status_code, estado.text[:200])
        if '"authorized"' not in estado.text:
            print("⚠️ La instancia de Green API NO está autorizada: hay que volver a vincular el WhatsApp escaneando el QR en console.green-api.com")
            instancia_ok = False
    except Exception as e:
        print("No se pudo consultar el estado de la instancia:", str(e))
    try:
        time.sleep(6)
        msg = requests.post(f"{base}/getMessage/{token}", json={"chatId": group_id, "idMessage": (res or {}).get("idMessage")}, timeout=20)
        print("Estado del mensaje:", msg.status_code, msg.text[:400])
    except Exception as e:
        print("No se pudo consultar el estado del mensaje:", str(e))
    print("Destino termina en:", group_id[-12:] if group_id else "(vacío)")
    return instancia_ok

if __name__ == "__main__":
    try:
        calendario_ok = True
        try:
            eventos, fecha_str = obtener_eventos_hoy()
        except Exception as e:
            print("No se pudo leer el calendario:", str(e))
            calendario_ok = False
            eventos = []
            fecha_str = datetime.now(pytz.timezone('America/Argentina/Buenos_Aires')).strftime("%d/%m/%Y")
        recordatorios = obtener_recordatorios_hoy()
        tareas = obtener_tareas_pendientes()
        pasos = obtener_pasos_seguimiento_hoy()

        texto_mensaje = f"📅 *AGENDA DEL DÍA - {fecha_str}*\n\n"

        if not calendario_ok:
            texto_mensaje += "⚠️ No se pudo leer el calendario hoy, revisá la agenda a mano.\n"
        elif not eventos:
            texto_mensaje += "🎉 ¡No hay eventos ni reuniones agendadas para hoy!\n"
        else:
            for evento in eventos:
                start = evento['start'].get('dateTime', evento['start'].get('date'))
                summary = evento.get('summary', 'Sin título')

                if 'T' in start:
                    hora = start.split('T')[1][:5]
                    texto_mensaje += f"⏰ *{hora} hs* - {summary}\n"
                else:
                    texto_mensaje += f"📌 *Todo el día* - {summary}\n"

        if recordatorios:
            texto_mensaje += "\n🔔 *RECORDATORIOS DE HOY*\n\n"
            for r in recordatorios:
                if r.get('hora'):
                    texto_mensaje += f"🕐 *{r['hora']} hs* - {r['texto']}\n"
                else:
                    texto_mensaje += f"• {r['texto']}\n"

        if pasos:
            texto_mensaje += "\n📋 *PASOS DE SEGUIMIENTO*\n\n"
            for p in pasos:
                fecha_corta = "/".join(reversed(p['fecha'].split('-')[1:]))
                estado = "vencido, previsto el %s" % fecha_corta if p.get('vencido') else "para hoy"
                texto_mensaje += f"• {p['cliente']} — {p['paso']} ({estado})\n"

        if tareas:
            texto_mensaje += "\n✅ *TAREAS PENDIENTES*\n\n"
            for t in tareas:
                texto_mensaje += f"• {t['texto']}\n"

        res = enviar_whatsapp(texto_mensaje)
        print("Mensaje enviado con éxito:", res)
        if not diagnosticar_envio(res):
            revisar_seguimientos_tareas()
            sys.exit(1)

        revisar_seguimientos_tareas()

    except Exception as e:
        print("Error al ejecutar el script:", str(e))
        sys.exit(1)
