import os
import json
import re
import pytz
from datetime import datetime
from flask import Flask, request, jsonify
import firebase_admin
from firebase_admin import credentials, firestore
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import requests

app = Flask(__name__)

# --- CATEGORÍA POR DEFECTO ---
# Podés cambiar 'varios' por 'procesales', 'oficina' o 'servicios' si preferís otra.
CATEGORIA_DEFECTO_CODE = 'varios'
CATEGORIA_DEFECTO_NOMBRE = '📋 Varios'

# --- INICIALIZAR FIREBASE ---
firebase_json = os.environ.get("FIREBASE_SERVICE_ACCOUNT_JSON")
if firebase_json:
    cred_dict = json.loads(firebase_json)
    cred = credentials.Certificate(cred_dict)
    firebase_admin.initialize_app(cred)
    db = firestore.client()

# --- FUNCION PARA CREAR EVENTO EN GOOGLE CALENDAR ---
def agregar_evento_calendar(resumen, hora_str):
    gcal_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not gcal_json:
        return False
    
    info = json.loads(gcal_json)
    scopes = ['https://www.googleapis.com/auth/calendar']
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    service = build('calendar', 'v3', credentials=creds)
    
    tz = pytz.timezone('America/Argentina/Buenos_Aires')
    ahora = datetime.now(tz)
    
    horas, minutos = map(int, hora_str.split(':'))
    inicio = ahora.replace(hour=horas, minute=minutos, second=0, microsecond=0)
    fin = inicio.replace(hour=horas+1)
    
    evento = {
        'summary': resumen,
        'start': {'dateTime': inicio.isoformat()},
        'end': {'dateTime': fin.isoformat()},
    }
    
    service.events().insert(calendarId='casihyasociados@gmail.com', body=evento).execute()
    return True

# --- RESPONDER POR WHATSAPP ---
def responder_whatsapp(chat_id, texto):
    instance_id = os.environ.get("GREEN_API_INSTANCE_ID")
    token = os.environ.get("GREEN_API_TOKEN")
    url = f"https://api.green-api.com/waInstance{instance_id}/sendMessage/{token}"
    requests.post(url, json={"chatId": chat_id, "message": texto})

# --- RUTA PRINCIPAL (WEBHOOK GREEN API) ---
@app.route('/webhook', methods=['POST'])
def webhook():
    data = request.json
    try:
        type_webhook = data.get('typeWebhook')
        if type_webhook == 'incomingMessageReceived':
            message_data = data.get('messageData', {})
            chat_id = data.get('senderData', {}).get('chatId')
            
            text_message = ""
            if 'textMessageData' in message_data:
                text_message = message_data['textMessageData'].get('textMessage', '').strip()
            elif 'extendedTextMessageData' in message_data:
                text_message = message_data['extendedTextMessageData'].get('text', '').strip()

            if not text_message:
                return jsonify({"status": "ignored"}), 200

            text_lower = text_message.lower()

            # 1. COMANDO: GASTO
            if text_lower.startswith('gasto'):
                match = re.search(r'\d+', text_message)
                if match:
                    monto = float(match.group())
                    
                    # Extraer solo la descripción removiendo la palabra 'gasto' y el monto
                    descripcion = re.sub(r'gasto|\d+', '', text_message, flags=re.IGNORECASE).strip()
                    if not descripcion:
                        descripcion = "Gasto registrado por WhatsApp"
                    
                    tz = pytz.timezone('America/Argentina/Buenos_Aires')
                    fecha_str = datetime.now(tz).strftime("%Y-%m-%d")
                    
                    # Guardar siempre en la categoría por defecto
                    db.collection('gastos').add({
                        'monto': monto,
                        'descripcion': descripcion,
                        'categoria': CATEGORIA_DEFECTO_CODE,
                        'fecha': fecha_str,
                        'timestamp': firestore.SERVER_TIMESTAMP
                    })
                    
                    responder_whatsapp(chat_id, f"✅ *Gasto registrado:* ${monto:.2f}\n📂 *Categoría:* {CATEGORIA_DEFECTO_NOMBRE}\n📝 *Detalle:* {descripcion}")

            # 2. COMANDO: INGRESO
            elif text_lower.startswith('ingreso'):
                match = re.search(r'\d+', text_message)
                if match:
                    monto = float(match.group())
                    descripcion = re.sub(r'ingreso|\d+', '', text_message, flags=re.IGNORECASE).strip()
                    if not descripcion:
                        descripcion = "Ingreso general"
                    
                    tz = pytz.timezone('America/Argentina/Buenos_Aires')
                    fecha_str = datetime.now(tz).strftime("%Y-%m-%d")
                    
                    db.collection('ingresos').add({
                        'monto': monto,
                        'descripcion': descripcion,
                        'fecha': fecha_str,
                        'timestamp': firestore.SERVER_TIMESTAMP
                    })
                    
                    responder_whatsapp(chat_id, f"💵 *Ingreso registrado:* ${monto:.2f}\n📝 *Detalle:* {descripcion}")

            # 3. COMANDO: EVENTO (ej: evento 16:30 Reunion)
            elif text_lower.startswith('evento'):
                match_hora = re.search(r'\b\d{1,2}:\d{2}\b', text_message)
                if match_hora:
                    hora_str = match_hora.group()
                    resumen = re.sub(r'evento|\b\d{1,2}:\d{2}\b', '', text_message, flags=re.IGNORECASE).strip()
                    if not resumen:
                        resumen = "Reunión / Evento"
                    
                    if agregar_evento_calendar(resumen, hora_str):
                        responder_whatsapp(chat_id, f"📅 *Evento agendado en Google Calendar:*\n⏰ *Hora:* {hora_str} hs\n📌 *Título:* {resumen}")

    except Exception as e:
        print("Error procesando Webhook:", str(e))
        
    return jsonify({"status": "success"}), 200

@app.route('/', methods=['GET'])
def health():
    return "Servidor del Bot activo y funcionando correctamente.", 200

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
