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

# --- NOMBRES DE MESES EN ESPAÑOL ---
MESES_ESP = {
    1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril',
    5: 'Mayo', 6: 'Junio', 7: 'Julio', 8: 'Agosto',
    9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
}

# --- INICIALIZAR FIREBASE ---
firebase_json = os.environ.get("FIREBASE_SERVICE_ACCOUNT_JSON")
if firebase_json:
    cred_dict = json.loads(firebase_json)
    cred = credentials.Certificate(cred_dict)
    firebase_admin.initialize_app(cred)
    db = firestore.client()

# --- FUNCION PARA CALCULOS DE FECHA Y HORA DE EVENTOS ---
def parsear_fecha_hora(texto):
    tz = pytz.timezone('America/Argentina/Buenos_Aires')
    ahora = datetime.now(tz)
    fecha_evento = None
    txt = texto.lower()

    if 'pasado mañana' in txt:
        fecha_evento = ahora + timedelta(days=2)
    elif 'mañana' in txt:
        fecha_evento = ahora + timedelta(days=1)
    elif 'hoy' in txt:
        fecha_evento = ahora
    else:
        match_fecha = re.search(r'\b(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?\b', txt)
        if match_fecha:
            dia = int(match_fecha.group(1))
            mes = int(match_fecha.group(2))
            anio = int(match_fecha.group(3)) if match_fecha.group(3) else ahora.year
            if anio < 100:
                anio += 2000
            try:
                fecha_temp = tz.localize(datetime(anio, mes, dia))
                if not match_fecha.group(3) and fecha_temp.date() < ahora.date():
                    anio += 1
                fecha_evento = tz.localize(datetime(anio, mes, dia))
            except ValueError:
                fecha_evento = None
        else:
            dias_semana = {
                'lunes': 0, 'martes': 1, 'miercoles': 2, 'miércoles': 2,
                'jueves': 3, 'viernes': 4, 'sabado': 5, 'sábado': 5, 'domingo': 6
            }
            for dia_nombre, dia_num in dias_semana.items():
                if dia_nombre in txt:
                    dias_hasta = (dia_num - ahora.weekday() + 7) % 7
                    if dias_hasta == 0:
                        dias_hasta = 7
                    fecha_evento = ahora + timedelta(days=dias_hasta)
                    break

    match_hora = re.search(r'\b(\d{1,2}):(\d{2})\b', txt)
    hora_str = None
    inicio = None
    fin = None
    hora_default = False

    if match_hora and fecha_evento:
        horas = int(match_hora.group(1))
        minutos = int(match_hora.group(2))
        if 0 <= horas <= 23 and 0 <= minutos <= 59:
            inicio = fecha_evento.replace(hour=horas, minute=minutos, second=0, microsecond=0)
            fin = inicio + timedelta(hours=1)
            hora_str = match_hora.group(0)
    elif fecha_evento and not match_hora:
        # No se especificó hora -> se asume 08:30 por defecto
        inicio = fecha_evento.replace(hour=8, minute=30, second=0, microsecond=0)
        fin = inicio + timedelta(hours=1)
        hora_str = "08:30"
        hora_default = True

    return inicio, fin, hora_str, fecha_evento, hora_default

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

# --- FUNCION PARA INTERPRETAR MONTOS EN CUALQUIER FORMATO ---
# Entiende: "1000000" / "1,000,000" / "1.000.000" / "1 millon" / "1.5 millones" / "500 mil"
def extraer_monto(texto):
    txt = texto.lower()

    # 1) "X millon(es)" -> multiplica por 1.000.000
    m = re.search(r'\d+(?:[.,]\d+)?\s*(?:millones|mill[oó]n|mill)\b(?:\s*de\s*pesos)?', txt)
    if m:
        numero = re.search(r'\d+(?:[.,]\d+)?', m.group(0)).group(0)
        monto = float(numero.replace(',', '.')) * 1_000_000
        texto_sin = texto[:m.start()] + texto[m.end():]
        return monto, texto_sin

    # 2) "X mil" -> multiplica por 1.000
    m = re.search(r'\d+(?:[.,]\d+)?\s*mil\b(?:\s*de\s*pesos)?', txt)
    if m:
        numero = re.search(r'\d+(?:[.,]\d+)?', m.group(0)).group(0)
        monto = float(numero.replace(',', '.')) * 1_000
        texto_sin = texto[:m.start()] + texto[m.end():]
        return monto, texto_sin

    # 3) numero plano, con o sin separadores de miles/decimales
    m = re.search(r'\d[\d.,]*', txt)
    if not m:
        return None, texto
    token = m.group(0)
    sep_count = token.count('.') + token.count(',')
    if sep_count == 0:
        monto = float(token)
    elif sep_count > 1:
        # mas de un separador -> todos son de miles (1.000.000 / 1,000,000)
        monto = float(re.sub(r'[.,]', '', token))
    else:
        sep_char = '.' if '.' in token else ','
        entero, frac = token.split(sep_char)
        if len(frac) == 3:
            # separador de miles (1.000 / 1,000)
            monto = float(entero + frac)
        else:
            # separador decimal (1250.50)
            monto = float(entero + '.' + frac)
    texto_sin = texto[:m.start()] + texto[m.end():]
    return monto, texto_sin

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
            ahora = datetime.now(tz)

            # Mismo formato que usa index.html: fecha "YYYY-MM-DD", mes de 2 dígitos,
            # año como string, y mesKey = "YYYY-MM" (es el campo que filtra Administración)
            fecha_hoy_str = ahora.strftime("%Y-%m-%d")
            mes_str = ahora.strftime("%m")
            anio_str = ahora.strftime("%Y")
            mes_key = f"{anio_str}-{mes_str}"

            # 0. COMANDO: AYUDA
            if text_lower.startswith('ayuda'):
                responder_whatsapp(
                    chat_id,
                    "🤖 *Comandos disponibles*\n\n"
                    "💰 *Gasto*\n`gasto MONTO CONCEPTO`\n_Ej: gasto 12000 librería_\n\n"
                    "💵 *Ingreso / Honorario*\n`ingreso MONTO CLIENTE`\n_Ej: ingreso 150000 Garcia_\n\n"
                    "📅 *Evento* (se crea en Google Calendar)\n`evento DÍA [HORA] TÍTULO`\n_Ej: evento miercoles 17:00 cita cliente Walter Figueroa_\n\n"
                    "⏰ *Recordatorio* (llega por WhatsApp a esta misma hora)\n`recordame DÍA [HORA] TEXTO`\n_Ej: recordame mañana 15:00 llamar a García_\n\n"
                    "📆 *Formas de indicar el día:* hoy, mañana, pasado mañana, un día de la semana (lunes, martes...), o una fecha exacta (15/08 o 15/08/2026).\n\n"
                    "🕐 Si no indicás la hora, se asume automáticamente las *08:30 hs*.\n\n"
                    "⚠️ El comando va siempre *al principio* del mensaje, y la hora en formato 24hs (15:00, no 3pm)."
                )

            # 1. COMANDO: GASTO
            elif text_lower.startswith('gasto'):
                monto, resto = extraer_monto(text_message)
                concepto = re.sub(r'gasto', '', resto, flags=re.IGNORECASE).strip()
                concepto = re.sub(r'\s{2,}', ' ', concepto).strip()

                if monto is None:
                    responder_whatsapp(chat_id, "⚠️ *Error al registrar gasto:*\nFalta indicar el *monto* numérico.\n\n💡 *Ejemplo:* `gasto 12000 librería`")
                elif not concepto:
                    responder_whatsapp(chat_id, "⚠️ *Error al registrar gasto:*\nFalta indicar el *concepto o detalle*.\n\n💡 *Ejemplo:* `gasto 12000 resma de hojas`")
                else:
                    doc_gasto = {
                        'fecha': fecha_hoy_str,
                        'concepto': concepto,
                        'cat': 'varios',            # clave usada por catLbl/catCls en index.html
                        'cargadoPor': 'Bot WhatsApp',
                        'monto': monto,
                        'mes': mes_str,              # "08" (2 dígitos, no "Agosto")
                        'anio': anio_str,            # "2026"
                        'mesKey': mes_key,           # "2026-08" -> esto es lo que filtra Administración
                        'creadoEn': firestore.SERVER_TIMESTAMP
                    }

                    db.collection('gastos').add(doc_gasto)

                    responder_whatsapp(chat_id, f"✅ *Gasto registrado correctamente*\n💰 *Monto:* ${monto:.2f}\n📝 *Concepto:* {concepto}\n📅 *Fecha:* {fecha_hoy_str}")

            # 2. COMANDO: INGRESO / HONORARIOS
            elif text_lower.startswith('ingreso') or text_lower.startswith('honorario'):
                monto, resto = extraer_monto(text_message)
                cliente = re.sub(r'ingreso|honorarios|honorario', '', resto, flags=re.IGNORECASE).strip()
                cliente = re.sub(r'\s{2,}', ' ', cliente).strip()

                if monto is None:
                    responder_whatsapp(chat_id, "⚠️ *Error al registrar ingreso:*\nFalta indicar el *monto* numérico.\n\n💡 *Ejemplo:* `ingreso 150000 Garcia`")
                else:
                    # El cliente es opcional (igual que en el formulario web) para no trabar la carga
                    concepto_txt = f"Cobro honorarios {cliente}" if cliente else "Cobro de honorarios"

                    doc_ingreso = {
                        'fecha': fecha_hoy_str,
                        'concepto': concepto_txt,
                        'cliente': cliente,          # puede quedar como ''
                        'cat': 'honorarios',          # clave usada por catLbl/catCls en index.html
                        'monto': monto,
                        'mes': mes_str,               # "08"
                        'anio': anio_str,              # "2026"
                        'mesKey': mes_key,            # "2026-08"
                        'creadoEn': firestore.SERVER_TIMESTAMP
                    }

                    db.collection('ingresos').add(doc_ingreso)

                    linea_cliente = f"\n👤 *Cliente:* {cliente}" if cliente else ""
                    responder_whatsapp(chat_id, f"💵 *Ingreso registrado correctamente*\n💰 *Monto:* ${monto:.2f}{linea_cliente}\n📝 *Concepto:* {concepto_txt}\n📅 *Fecha:* {fecha_hoy_str}")

            # 3. COMANDO: EVENTO
            elif text_lower.startswith('evento'):
                inicio, fin, hora_str, fecha_detectada, hora_default = parsear_fecha_hora(text_message)
                
                resumen = re.sub(
                    r'evento|mañana|pasado mañana|hoy|lunes|martes|miercoles|miércoles|jueves|viernes|sabado|sábado|domingo|\b\d{1,2}:\d{2}\b|\b\d{1,2}/\d{1,2}(?:/\d{2,4})?\b', 
                    '', 
                    text_message, 
                    flags=re.IGNORECASE
                ).strip()

                if not fecha_detectada:
                    responder_whatsapp(
                        chat_id, 
                        "⚠️ *Error al crear evento:*\nNo se detectó el *día o fecha*.\n\n"
                        "💡 *Ejemplo:* `evento miercoles 17:00 cita cliente Walter Figueroa`"
                    )
                elif not resumen:
                    responder_whatsapp(
                        chat_id, 
                        "⚠️ *Error al crear evento:*\nFalta indicar el *motivo o título*.\n\n"
                        "💡 *Ejemplo:* `evento mañana 16:30 Reunion with Client`"
                    )
                else:
                    try:
                        if agregar_evento_calendar(resumen, inicio, fin):
                            fecha_formateada = inicio.strftime("%d/%m/%Y")
                            nota_default = " _(hora no indicada, se asumió por defecto)_" if hora_default else ""
                            responder_whatsapp(
                                chat_id, 
                                f"📅 *Evento agendado en Google Calendar:*\n"
                                f"📆 *Fecha:* {fecha_formateada}\n"
                                f"⏰ *Hora:* {hora_str} hs{nota_default}\n"
                                f"📌 *Título:* {resumen}"
                            )
                        else:
                            responder_whatsapp(chat_id, "❌ *Error:* No se pudo conectar con Google Calendar.")
                    except Exception as err:
                        responder_whatsapp(chat_id, f"❌ *Error al crear evento:* {str(err)}")

            # 4. COMANDO: RECORDATORIO
            elif text_lower.startswith('recordame') or text_lower.startswith('recordá') or text_lower.startswith('recorda'):
                inicio, fin, hora_str, fecha_detectada, hora_default = parsear_fecha_hora(text_message)

                texto_recordatorio = re.sub(
                    r'recordame|recorda|recordá|mañana|pasado mañana|hoy|lunes|martes|miercoles|miércoles|jueves|viernes|sabado|sábado|domingo|\b\d{1,2}:\d{2}\b|\b\d{1,2}/\d{1,2}(?:/\d{2,4})?\b',
                    '',
                    text_message,
                    flags=re.IGNORECASE
                ).strip()

                if not fecha_detectada:
                    responder_whatsapp(
                        chat_id,
                        "⚠️ *Error al crear recordatorio:*\nNo se detectó el *día o fecha*.\n\n"
                        "💡 *Ejemplo:* `recordame mañana 15:00 llamar a García`"
                    )
                elif not texto_recordatorio:
                    responder_whatsapp(
                        chat_id,
                        "⚠️ *Error al crear recordatorio:*\nFalta indicar *qué* recordar.\n\n"
                        "💡 *Ejemplo:* `recordame mañana 15:00 llamar a García`"
                    )
                else:
                    doc_recordatorio = {
                        'chatId': chat_id,           # a este mismo chat (privado o grupo) se le responde
                        'fecha': inicio.strftime("%Y-%m-%d"),
                        'hora': inicio.strftime("%H:%M"),
                        'texto': texto_recordatorio,
                        'enviado': False,
                        'creadoEn': firestore.SERVER_TIMESTAMP
                    }
                    db.collection('recordatorios_bot').add(doc_recordatorio)

                    fecha_formateada = inicio.strftime("%d/%m/%Y")
                    nota_default = " _(hora no indicada, se asumió por defecto)_" if hora_default else ""
                    responder_whatsapp(
                        chat_id,
                        f"⏰ *Recordatorio guardado*\n"
                        f"📆 *Fecha:* {fecha_formateada}\n"
                        f"🕐 *Hora:* {hora_str} hs{nota_default}\n"
                        f"📝 *Texto:* {texto_recordatorio}"
                    )

    except Exception as e:
        print("Error procesando Webhook:", str(e))
        
    return jsonify({"status": "success"}), 200

# --- RUTA PARA REVISAR Y ENVIAR RECORDATORIOS PENDIENTES ---
# La llama cron-job.org cada 5 minutos (GET simple). De paso, ese mismo
# ping mantiene despierto el servicio en Render, que se duerme tras 15
# minutos sin actividad en el plan gratuito.
@app.route('/revisar-recordatorios', methods=['GET'])
def revisar_recordatorios():
    tz = pytz.timezone('America/Argentina/Buenos_Aires')
    ahora = datetime.now(tz)
    enviados = 0
    errores = 0

    try:
        pendientes = db.collection('recordatorios_bot').where('enviado', '==', False).stream()
        for doc in pendientes:
            data = doc.to_dict()
            try:
                fecha_hora_str = f"{data['fecha']} {data['hora']}"
                programado = tz.localize(datetime.strptime(fecha_hora_str, "%Y-%m-%d %H:%M"))

                if ahora >= programado:
                    responder_whatsapp(
                        data['chatId'],
                        f"⏰ *Recordatorio:*\n{data['texto']}"
                    )
                    doc.reference.update({'enviado': True, 'enviadoEn': firestore.SERVER_TIMESTAMP})
                    enviados += 1
            except Exception as e_inner:
                print("Error procesando recordatorio", doc.id, str(e_inner))
                errores += 1
    except Exception as e:
        print("Error revisando recordatorios:", str(e))
        return jsonify({"status": "error", "detalle": str(e)}), 500

    return jsonify({"status": "ok", "enviados": enviados, "errores": errores, "hora_chequeo": ahora.isoformat()}), 200

@app.route('/', methods=['GET'])
def health():
    return "Servidor del Bot activo y funcionando correctamente.", 200

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
