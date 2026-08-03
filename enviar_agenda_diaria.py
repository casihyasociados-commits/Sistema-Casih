import os
import json
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
    response = requests.post(url, json=payload, headers=headers)
    return response.json()

if __name__ == "__main__":
    try:
        eventos, fecha_str = obtener_eventos_hoy()
        
        texto_mensaje = f"📅 *AGENDA DEL DÍA - {fecha_str}*\n\n"
        
        if not eventos:
            texto_mensaje += "🎉 ¡No hay eventos ni reuniones agendadas para hoy!"
        else:
            for evento in eventos:
                start = evento['start'].get('dateTime', evento['start'].get('date'))
                summary = evento.get('summary', 'Sin título')
                
                if 'T' in start:
                    hora = start.split('T')[1][:5]
                    texto_mensaje += f"⏰ *{hora} hs* - {summary}\n"
                else:
                    texto_mensaje += f"📌 *Todo el día* - {summary}\n"
        
        res = enviar_whatsapp(texto_mensaje)
        print("Mensaje enviado con éxito:", res)
        
    except Exception as e:
        print("Error al ejecutar el script:", str(e))
