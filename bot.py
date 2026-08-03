import os
import json
import re
import pytz
from datetime import datetime, timedelta
from flask import Flask, request, jsonify
import firebase_admin
from firebase_admin import credentials, firestore
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import requests

app = Flask(__name__)

# --- CONFIGURACIÓN Y CATEGORÍAS ---
CATEGORIA_DEFECTO = 'varios'

# --- INICIALIZAR FIREBASE ---
firebase_json = os.environ.get("FIREBASE_SERVICE_ACCOUNT_JSON")
if firebase_json:
    cred_dict = json.loads(firebase_json)
    cred = credentials.Certificate(cred_dict)
    firebase_admin.initialize_app(cred)
    db = firestore.client()

# --- FUNCION INTELIGENTE PARA CALCULAR LA FECHA DEL EVENTO ---
def parsear_fecha_hora(texto):
    tz = pytz.timezone('America/Argentina/Buenos_Aires')
    ahora = datetime.now(tz)
    fecha_evento = ahora
    txt = texto.lower()

    # 1. Detección por texto relativo
    if 'pasado mañana' in txt:
        fecha_evento = ahora + timedelta(days=2)
    elif 'mañana' in txt:
        fecha_evento = ahora + timedelta(days=1)
    else:
        # 2. Buscar fechas numéricas (ej: 15/08 o 15/08/2026)
        match_fecha = re.search(r'\b(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?\b', txt)
        if match_fecha:
            dia = int(match_fecha.group(1))
            mes = int(match_fecha.group(2))
            
            if match_fecha.group(3):
                anio = int(match_fecha.group(3))
                if anio < 100:
                    anio += 2000
            else:
                # Si solo se pone DD/MM, se asume el año actual
                anio = ahora.year
                try:
                    # Si la fecha de este año ya pasó, asume el año siguiente
                    fecha_temp = tz.localize(datetime(anio, mes, dia))
                    if fecha_temp.date() < ahora.date():
                        anio += 1
                except ValueError:
                    pass

            try:
                fecha_evento = tz.localize(datetime(anio, mes, dia))
            except ValueError:
                pass
        else:
            # 3. Buscar días de la semana (ej: "el viernes")
            dias_semana = {
                'lunes': 0, 'martes': 1, 'miercoles': 2, 'miércoles': 2,
                'jueves': 3, 'viernes': 4, 'sabado': 5, 'sábado': 5, 'domingo': 6
            }
            for dia_nombre, dia_num in dias_semana.items():
                if dia_nombre in txt:
                    dias_hasta = (dia_num - ahora.weekday() + 7) % 7
                    if dias_hasta == 0:
                        dias_hasta = 7  # Si es hoy, refiere al mismo día de la próxima semana
                    fecha_evento = ahora + timedelta(days=dias_hasta)
                    break

    # Detección de hora (ej: 16:30 o 9:00)
    match_hora = re.search(r'\b(\d{1,2}):(\d{2})\b', txt)
    if match_hora:
        horas = int(match_hora.group(1))
        minutos = int(match_hora.group(2))
        inicio = fecha_evento.replace(hour=horas, minute=minutos, second=0, microsecond=0)
        fin = inicio + timedelta(hours=1)
        return inicio, fin, match_hora.group(0)
    
    return None, None, None

# --- CREAR EVENTO EN GOOGLE CALENDAR ---
def agregar_evento_calendar(resumen, inicio, fin):
    gcal_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not gcal_json:
        return False
    
    info = json.loads(gcal_json)
    scopes = ['https://www.googleapis.com/auth/calendar']
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    service = build('calendar', 'v3', credentials=creds)
    
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

# --- RUTA PRINCIPAL (WEBHOOK) ---
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
            tz = pytz.timezone('America/Argentina/Buenos_Aires')
            fecha_hoy_str = datetime.now(tz).strftime("%Y-%m-%d")

            # 1. COMANDO: GASTO
            if text_lower.startswith('gasto'):
                match = re.search(r'\d+', text_message)
                if match:
                    monto = float(match.group())
                    descripcion = re.sub(r'gasto|\d+', '', text_message, flags=re.IGNORECASE).strip()
                    if not descripcion:
                        descripcion = "Gasto WhatsApp"
                    
                    db.collection('gastos').add({
                        'monto': monto,
                        'descripcion': descripcion,
                        'categoria': CATEGORIA_DEFECTO,
                        'fecha': fecha_hoy_str,
                        'timestamp': firestore.SERVER_TIMESTAMP
                    })
                    
                    responder_whatsapp(chat_id, f"✅ *Gasto registrado:* ${monto:.2f}\n📂 *Categoría:* Gastos varios\n📝 *Detalle:* {descripcion}")

            # 2. COMANDO: INGRESO / HONORARIOS
            elif text_lower.startswith('ingreso') or text_lower.startswith('honorario'):
                match = re.search(r'\d+', text_message)
                if match:
                    monto = float(match.group())
                    # Elimina el comando y el monto para aislar el cliente/detalle
                    detalle = re.sub(r'ingreso|honorarios|honorario|\d+', '', text_message, flags=re.IGNORECASE).strip()
                    
                    cliente = detalle if detalle else "General"
                    descripcion = f"Honorarios - {cliente}" if cliente != "General" else "Ingreso Honorarios"

                    db.collection('ingresos').add({
                        'monto': monto,
                        'descripcion': descripcion,
                        'cliente': cliente,
                        'fecha': fecha_hoy_str,
                        'timestamp': firestore.SERVER_TIMESTAMP
                    })
                    
                    responder_whatsapp(chat_id, f"💵 *Ingreso registrado:* ${monto:.2f}\n👤 *Cliente/Detalle:* {cliente}\n📅 *Fecha:* {fecha_hoy_str}")

            # 3. COMANDO: EVENTO
            elif text_lower.startswith('evento'):
                inicio, fin, hora_str = parsear_fecha_hora(text_message)
                if inicio and fin:
                    # Quitar palabras clave y fechas del resumen del evento
                    resumen = re.sub(r'evento|mañana|pasado mañana|lunes|martes|miercoles|miércoles|jueves|viernes|sabado|sábado|domingo|\b\d{1,2}:\d{2}\b|\b\d{1,2}/\d{1,2}(?:/\d{2,4})?\b', '', text_message, flags=re.IGNORECASE).strip()
                    if not resumen:
                        resumen = "Reunión / Evento"
                    
                    if agregar_evento_calendar(resumen, inicio, fin):
                        fecha_formateada = inicio.strftime("%d/%m/%Y")
                        responder_whatsapp(chat_id, f"📅 *Evento agendado en Google Calendar:*\n📆 *Fecha:* {fecha_formateada}\n⏰ *Hora:* {hora_str} hs\n📌 *Título:* {resumen}")

    except Exception as e:
        print("Error procesando Webhook:", str(e))
        
    return jsonify({"status": "success"}), 200

@app.route('/', methods=['GET'])
def health():
    return "Servidor del Bot activo y funcionando correctamente.", 200

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
