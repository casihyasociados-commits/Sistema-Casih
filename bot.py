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
import anthropic

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

# --- SEGUIMIENTO CORREO ARGENTINO ---
# El formulario publico de seguimiento no pide captcha: es un POST simple que
# devuelve un fragmento HTML con la tabla de movimientos (mas nuevo primero).
CA_URL = "https://www.correoargentino.com.ar/sites/all/modules/custom/ca_forms/api/wsFacade.php"

def consultar_correo_ca(prefijo, numero):
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'X-Requested-With': 'XMLHttpRequest',
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36',
        'Referer': 'https://www.correoargentino.com.ar/formularios/ondnc'
    }
    payload = {'action': 'ondnc', 'id': numero, 'producto': prefijo, 'pais': 'AR'}
    resp = requests.post(CA_URL, data=payload, headers=headers, timeout=30)
    resp.raise_for_status()
    html = resp.content.decode('utf-8-sig', errors='replace')

    if 'No se encontraron resultados' in html:
        return []

    filas = re.findall(
        r'data-title="Fecha:">(.*?)</td>\s*'
        r'<td data-title="Planta:">(.*?)</td>\s*'
        r'<td data-title="Historia:">(.*?)</td>\s*'
        r'<td data-title="Estado:">(.*?)</td>',
        html, re.S
    )
    return [
        {'fecha': f.strip(), 'planta': p.strip(), 'historia': h.strip(), 'estado': e.strip()}
        for f, p, h, e in filas
    ]

def procesar_pieza_correo(doc, data, grupo):
    """Consulta una pieza, actualiza Firestore y devuelve que paso.
    Resultado: 'error' | 'sin_resultados' | 'primera_vez' | 'sin_cambios' | 'con_novedad'
    """
    prefijo = data.get('caPrefijo')
    numero = data.get('caNumero')

    try:
        movs = consultar_correo_ca(prefijo, numero)
    except Exception as e_req:
        print("Error consultando", prefijo, numero, str(e_req))
        doc.reference.update({'caError': 'No se pudo consultar', 'caUltimaConsulta': firestore.SERVER_TIMESTAMP})
        return 'error'

    if not movs:
        doc.reference.update({'caError': 'Sin resultados en Correo Argentino', 'caUltimaConsulta': firestore.SERVER_TIMESTAMP})
        return 'sin_resultados'

    ultimo = movs[0]
    previos = data.get('caMovimientos') or 0

    update = {
        'caEstado': ultimo['estado'],
        'caHistoria': ultimo['historia'],
        'caPlanta': ultimo['planta'],
        'caFechaMov': ultimo['fecha'],
        'caMovimientos': len(movs),
        'caError': '',
        'caUltimaConsulta': firestore.SERVER_TIMESTAMP
    }

    # Una vez entregada no se consulta mas
    if ultimo['estado'].upper() == 'ENTREGADO':
        update['caFinalizado'] = True

    # La primera consulta solo deja registrado el estado actual: avisar de
    # piezas recien cargadas seria ruido, no novedad.
    primera_vez = not data.get('caConsultado')
    update['caConsultado'] = True

    doc.reference.update(update)

    if primera_vez:
        return 'primera_vez'

    if len(movs) <= previos:
        return 'sin_cambios'

    if grupo:
        cliente = data.get('cliente', 'Sin cliente')
        pieza = "%s-%s-AR" % (prefijo, numero)
        estado = ultimo['estado'].upper()

        if estado == 'RECHAZADO':
            msg = ("⚠️ *ENVÍO RECHAZADO*\n\n"
                   "👤 *Cliente:* %s\n"
                   "📦 *Pieza:* %s\n"
                   "📍 *Movimiento:* %s\n"
                   "🏢 *Planta:* %s\n"
                   "📅 *Fecha:* %s\n\n"
                   "_El destinatario rechazó la pieza. Revisar si requiere acción._"
                   % (cliente, pieza, ultimo['historia'], ultimo['planta'], ultimo['fecha']))
        elif estado == 'ENTREGADO':
            msg = ("✅ *Envío entregado*\n\n"
                   "👤 *Cliente:* %s\n"
                   "📦 *Pieza:* %s\n"
                   "🏢 *Planta:* %s\n"
                   "📅 *Fecha:* %s"
                   % (cliente, pieza, ultimo['planta'], ultimo['fecha']))
        else:
            msg = ("📦 *Novedad en envío*\n\n"
                   "👤 *Cliente:* %s\n"
                   "📦 *Pieza:* %s\n"
                   "📍 *Movimiento:* %s\n"
                   "🏢 *Planta:* %s\n"
                   "📅 *Fecha:* %s"
                   % (cliente, pieza, ultimo['historia'], ultimo['planta'], ultimo['fecha']))

        responder_whatsapp(grupo, msg)

    return 'con_novedad'

@app.route('/revisar-correos', methods=['GET'])
def revisar_correos():
    contadores = {'error': 0, 'sin_resultados': 0, 'primera_vez': 0, 'sin_cambios': 0, 'con_novedad': 0}
    consultados = 0

    try:
        docs = list(db.collection('correos').stream())
    except Exception as e:
        print("Error leyendo correos:", str(e))
        return jsonify({"status": "error", "detalle": str(e)}), 500

    grupo = os.environ.get("GREEN_API_GROUP_ID")

    for doc in docs:
        data = doc.to_dict()
        # Solo las piezas con codigo cargado y todavia no finalizadas
        if not data.get('caPrefijo') or not data.get('caNumero') or data.get('caFinalizado'):
            continue
        consultados += 1
        resultado = procesar_pieza_correo(doc, data, grupo)
        contadores[resultado] = contadores.get(resultado, 0) + 1

    errores = contadores['error']
    # Si fallaron todas, probablemente cambio el sitio o nos estan bloqueando:
    # conviene enterarse en vez de que el chequeo quede roto en silencio.
    if grupo and consultados > 0 and errores == consultados:
        responder_whatsapp(
            grupo,
            "⚠️ *Seguimiento de correo*\nNo se pudo consultar ninguna pieza hoy (%d intento/s). "
            "Puede que Correo Argentino haya cambiado el sitio o esté bloqueando las consultas." % errores
        )

    return jsonify({
        "status": "ok",
        "consultados": consultados,
        "primera_consulta": contadores['primera_vez'],
        "con_novedad": contadores['con_novedad'],
        "sin_cambios": contadores['sin_cambios'],
        "errores": errores
    }), 200

# Se llama desde el navegador (index.html) justo despues de cargar un
# seguimiento nuevo, para no esperar hasta el proximo chequeo programado.
# Lleva CORS propio porque el resto del bot no lo necesita (solo lo pegan
# el webhook de WhatsApp y cron-job.org, ambos server-to-server).
@app.route('/revisar-correo-individual', methods=['GET'])
def revisar_correo_individual():
    doc_id = request.args.get('id')
    if not doc_id:
        resp = jsonify({"status": "error", "detalle": "Falta el parametro id"})
        resp.headers['Access-Control-Allow-Origin'] = '*'
        return resp, 400

    ref = db.collection('correos').document(doc_id)
    snap = ref.get()
    if not snap.exists:
        resp = jsonify({"status": "error", "detalle": "No existe ese seguimiento"})
        resp.headers['Access-Control-Allow-Origin'] = '*'
        return resp, 404

    data = snap.to_dict()
    if not data.get('caPrefijo') or not data.get('caNumero'):
        resp = jsonify({"status": "error", "detalle": "Ese registro no tiene codigo de Correo Argentino"})
        resp.headers['Access-Control-Allow-Origin'] = '*'
        return resp, 400

    grupo = os.environ.get("GREEN_API_GROUP_ID")
    resultado = procesar_pieza_correo(ref, data, grupo)

    resp = jsonify({"status": "ok", "resultado": resultado})
    resp.headers['Access-Control-Allow-Origin'] = '*'
    return resp, 200

# --- ESCRITOS CON IA ---
# Plantilla anonimizada (sin datos reales de ningun cliente) que le muestra a
# Claude la estructura y el registro formal que usa el estudio. El contenido
# especifico del caso nuevo se agrega aparte, en PROMPT_HECHOS.
PLANTILLA_DENUNCIA_LABORAL = """FORMULA DENUNCIA.

Sr. Director de Reclamaciones Individuales
Departamento Provincial del Trabajo.
S___________________/____________D

[NOMBRE Y APELLIDO DEL DENUNCIANTE], CUIL [CUIL], [EDAD] años de edad, de estado civil [ESTADO CIVIL], de nacionalidad [NACIONALIDAD], actualmente [SITUACION LABORAL ACTUAL], con domicilio real en [DOMICILIO], Ciudad de Córdoba. Celular [TELEFONO], correo electrónico: [EMAIL], acompañado por la Dra. Ciardiello María Paula M.P. 1-41227, paulaciardiello@hotmail.com, CIDI NIVEL 2, CUIT 27-38003502-8, cel. 3517541685, constituyendo domicilio a los efectos procesales en calle Arturo M. Bas Nº 389 1º Piso Oficina "A", ambos de esta ciudad, ante S.S. respetuosamente comparezco y digo:

Que vengo a formular denuncia en contra de mi empleadora [NOMBRE DEL EMPLEADOR] CUIL/CUIT [CUIL EMPLEADOR], titular y/o responsable de [DESCRIPCION DEL COMERCIO/EMPRESA] con domicilio en [DOMICILIO EMPLEADOR], comparecemos y decimos:

Que ingresé a trabajar bajo las órdenes de la denunciada el [FECHA DE INGRESO], cumpliendo tareas acordes con la categoría [CATEGORIA LABORAL].

Que mi jornada de trabajo era [DESCRIPCION DE LA JORNADA: días y horarios].

Que las tareas desempeñadas por mi parte consistían en [DESCRIPCION DETALLADA DE LAS TAREAS].

Que durante toda la relación laboral presté servicios de manera personal, habitual y bajo dependencia de los denunciados, sin que mi vínculo laboral se encontrara debidamente registrado ante los organismos correspondientes, en violación a la normativa laboral vigente. [Omitir este párrafo de "sin registración" si el caso es de un trabajador SÍ registrado.]

Que la relación laboral continuó desarrollándose en los términos precedentemente expuestos hasta [DESARROLLO DE LOS HECHOS QUE MOTIVAN LA DENUNCIA: despido, impedimento de ingreso, falta de pago, acoso, etc. — este es el núcleo narrativo del caso, se redacta a partir de los hechos concretos que aporte el usuario].

[Si corresponde: descripción de intimaciones previas, cartas documento o telegramas enviados, con fecha y número, y su contenido resumido en discurso indirecto o citado entre comillas cuando el usuario aporte el texto exacto.]

[Si corresponde: respuesta de la contraparte, rechazo, o silencio.]

Que al día de la fecha la denunciada no ha dado cumplimiento a [LO RECLAMADO], ni ha hecho entrega de la certificación de servicios y remuneraciones prevista en el art. 80 LCT [si aplica].

Por los motivos expuestos, se solicita la intervención de esta Autoridad de Aplicación.

En definitiva, ante el incumplimiento de la normativa vigente por parte del Empleador, solicitó que el mismo sea citado, quien deberá concurrir con toda la documentación laboral, legajo personal, además de las constancias de pago de los aportes y contribuciones sindicales, sociales y previsionales, y los importes correspondientes a Liquidación Final e indemnizaciones de ley, a cuyo fin deberá fijar día y hora de audiencia de conciliación. Sin otro particular. Saludo a Ud. muy atte."""

TIPOS_ESCRITO = {
    'denuncia_trabajo': {
        'nombre': 'Denuncia laboral — Ministerio de Trabajo',
        'plantilla': PLANTILLA_DENUNCIA_LABORAL
    }
}

@app.route('/generar-escrito', methods=['POST', 'OPTIONS'])
def generar_escrito():
    if request.method == 'OPTIONS':
        resp = jsonify({})
        resp.headers['Access-Control-Allow-Origin'] = '*'
        resp.headers['Access-Control-Allow-Methods'] = 'POST, OPTIONS'
        resp.headers['Access-Control-Allow-Headers'] = 'Content-Type'
        return resp, 200

    body = request.get_json(silent=True) or {}
    tipo = body.get('tipo')
    cliente = body.get('cliente') or {}
    hechos = (body.get('hechos') or '').strip()

    def error(msg, code=400):
        resp = jsonify({"status": "error", "detalle": msg})
        resp.headers['Access-Control-Allow-Origin'] = '*'
        return resp, code

    if tipo not in TIPOS_ESCRITO:
        return error('Tipo de escrito no reconocido.')
    if not cliente.get('nombre'):
        return error('Falta seleccionar un cliente.')
    if not hechos:
        return error('Falta describir los hechos del caso.')

    api_key = os.environ.get('ANTHROPIC_API_KEY')
    if not api_key:
        return error('El servidor no tiene configurada la clave de la API de Claude.', 500)

    datos_cliente = "\n".join([
        "Nombre completo: %s" % cliente.get('nombre', ''),
        "CUIL: %s" % cliente.get('cuil', ''),
        "Edad: %s" % cliente.get('edad', ''),
        "Domicilio: %s" % cliente.get('domicilio', ''),
        "Teléfono: %s" % cliente.get('tel', ''),
        "Empleador: %s" % cliente.get('empleador', ''),
        "Categoría / motivo: %s" % cliente.get('motivo', ''),
        "Antigüedad: %s" % cliente.get('antiguedad', ''),
    ])

    system_prompt = (
        "Sos un asistente de redacción para un estudio jurídico laboralista de Córdoba, Argentina "
        "(CASIH & Asociados). Tu tarea es redactar un escrito legal nuevo siguiendo EXACTAMENTE la "
        "estructura, el registro formal y las fórmulas jurídicas de la plantilla que se te da a "
        "continuación, pero reemplazando los datos del cliente y redactando de nuevo el relato de "
        "los hechos a partir de la información real del caso nuevo que te va a dar el usuario.\n\n"
        "Reglas:\n"
        "- Mantené el mismo tono formal, las mismas fórmulas de estilo ('Que vengo a...', 'Que, "
        "ante ello...', 'Sin otro particular. Saludo a Ud. muy atte.') y la misma estructura de "
        "párrafos que la plantilla.\n"
        "- Los textos entre corchetes [ASI] son instrucciones para vos, no van en el resultado: "
        "reemplazalos por el dato real, o si el dato no fue provisto, redactá la oración sin ese "
        "dato en vez de dejar el corchete o inventar información.\n"
        "- El párrafo del relato de los hechos (motivo de la denuncia, intimaciones previas, "
        "rechazos, etc.) no es un molde a completar: redactalo de cero en el mismo estilo formal, "
        "a partir de los hechos que te cuenta el usuario en su propio lenguaje.\n"
        "- Nunca inventes fechas, montos, números de carta documento ni nombres que el usuario no "
        "haya mencionado.\n"
        "- Devolvé únicamente el texto final del escrito, sin comentarios tuyos antes o después."
    )

    user_prompt = (
        "PLANTILLA DE REFERENCIA (estructura y estilo a seguir):\n\n%s\n\n"
        "DATOS DEL CLIENTE PARA ESTE ESCRITO:\n%s\n\n"
        "HECHOS DEL CASO (tal como los describió quien carga el sistema):\n%s\n\n"
        "Redactá el escrito completo aplicando la plantilla de arriba a este cliente y estos hechos."
    ) % (TIPOS_ESCRITO[tipo]['plantilla'], datos_cliente, hechos)

    try:
        client = anthropic.Anthropic(api_key=api_key)
        response = client.messages.create(
            model="claude-sonnet-5",
            max_tokens=4096,
            system=system_prompt,
            messages=[{"role": "user", "content": user_prompt}]
        )
        texto = "".join([b.text for b in response.content if b.type == "text"])
    except Exception as e:
        print("Error generando escrito:", str(e))
        return error('No se pudo generar el escrito. Intentá de nuevo en un momento.', 500)

    resp = jsonify({"status": "ok", "texto": texto})
    resp.headers['Access-Control-Allow-Origin'] = '*'
    return resp, 200

@app.route('/', methods=['GET'])
def health():
    return "Servidor del Bot activo y funcionando correctamente.", 200

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
