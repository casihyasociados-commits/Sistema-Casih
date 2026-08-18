import os
import json
import re
import pytz
from io import BytesIO
from datetime import datetime, timedelta
from flask import Flask, request, jsonify, send_file
import firebase_admin
from firebase_admin import credentials, firestore
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import requests
import anthropic
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

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
    try:
        r = requests.post(url, json={"chatId": chat_id, "message": texto}, timeout=15)
        print("[GREEN-API sendMessage] status=%s body=%s" % (r.status_code, r.text[:300]))
        return r.status_code == 200
    except Exception as e:
        print("[GREEN-API sendMessage] ERROR:", str(e))
        return False

# --- RUTA PRINCIPAL (WEBHOOK) ---
@app.route('/webhook', methods=['POST'])
def webhook():
    data = request.json
    print("[WEBHOOK recibido]", json.dumps(data)[:600] if data else "sin body/JSON")
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
                    ok = responder_whatsapp(
                        data['chatId'],
                        f"⏰ *Recordatorio:*\n{data['texto']}"
                    )
                    if ok:
                        doc.reference.update({'enviado': True, 'enviadoEn': firestore.SERVER_TIMESTAMP})
                        enviados += 1
                    else:
                        # No se pudo confirmar la entrega (ej. cupo de Green API agotado):
                        # se deja pendiente para reintentar en la proxima corrida del cron.
                        print("Recordatorio no confirmado, se reintenta:", doc.id)
                        errores += 1
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
# Todo lo que es fijo (representacion legal, encabezado, formulas de cierre)
# se arma con texto literal, sin pasar por el modelo: asi los datos duros
# (nombres, CUIT, fechas) nunca corren riesgo de ser alterados por la IA.
# Claude solo redacta dos parrafos puntuales, donde hace falta traducir
# lenguaje informal a estilo formal: la descripcion de tareas y el relato de
# los hechos posteriores al despido.

BLOQUES_REPRESENTACION = {
    'ciardiello': (
        'acompañado por la Dra. Ciardiello María Paula M.P. 1-41227, CIDI NIVEL 2, '
        'CUIT 27-38003502-8, correo electrónico paulaciardiello@hotmail.com, '
        'cel. 3517541685, constituyendo domicilio a los efectos procesales en calle '
        'Arturo M. Bas Nº 389 1º Piso Oficina "A"'
    ),
    'casih': (
        'acompañado por el Dr. Pablo CASIH M.P. 1-31142 CIDI NIVEL 2, '
        'CUIT 20-22220869-7, correo electrónico pcasih@hotmail.com, '
        'cel. 3515102323, constituyendo domicilio a los efectos procesales en calle '
        'Arturo M. Bas Nº 389 1º Piso Oficina "A"'
    ),
    'ambos': (
        'acompañado por la Dra. Ciardiello María Paula M.P. 1-41227, CIDI NIVEL 2, '
        'CUIT 27-38003502-8, correo electrónico paulaciardiello@hotmail.com, '
        'cel. 3517541685 y por el Dr. Pablo CASIH M.P. 1-31142 CIDI NIVEL 2, '
        'CUIT 20-22220869-7, correo electrónico pcasih@hotmail.com, '
        'cel. 3515102323, constituyendo domicilio a los efectos procesales en calle '
        'Arturo M. Bas Nº 389 1º Piso Oficina "A"'
    ),
}

# El formulario solo soporta un denunciado (persona, empresa, o una persona
# titular de una empresa/comercio), asi que siempre es singular. Genero
# correcto: si hay persona fisica, se usa el genero que cargo el usuario; si
# es solo empresa, se trata como "la empresa" (sustantivo femenino).
def denunciado_es_femenino(den):
    if den.get('esPersona'):
        return (den.get('genero') or '').strip().lower() == 'femenino'
    return True

def de_denunciado(den):
    return 'de la denunciada' if denunciado_es_femenino(den) else 'del denunciado'

def articulo_denunciado(den):
    return 'la denunciada' if denunciado_es_femenino(den) else 'el denunciado'

def armar_bloque_denunciado(den):
    persona = (den.get('esPersona') and (den.get('nombre') or '').strip())
    empresa = (den.get('esEmpresa') and (den.get('empresa') or '').strip())
    domicilio = (den.get('domicilio') or '').strip()
    ciudad = (den.get('ciudad') or '').strip()
    fantasia = (den.get('nombreFantasia') or '').strip()
    email = (den.get('email') or '').strip()

    # Persona fisica que opera un comercio con nombre de fantasia propio
    # (sin ser una empresa/CUIT aparte): usamos esta forma en vez de la de
    # "persona + empresa" para no pedirle un segundo CUIT que no existe.
    if persona and fantasia:
        base = 'mi empleador/a %s CUIL %s, titular y/o responsable del local comercial que gira bajo el nombre de fantasía: "%s"' % (
            den.get('nombre', ''), den.get('cuitPersona', ''), fantasia
        )
    elif persona and empresa:
        base = 'mi empleador/a %s CUIL %s, titular y/o responsable de la empresa %s CUIT %s, dedicada a %s' % (
            den.get('nombre', ''), den.get('cuitPersona', ''),
            den.get('empresa', ''), den.get('cuitEmpresa', ''), den.get('rubro', '')
        )
    elif persona:
        base = 'mi empleador/a %s CUIL %s' % (den.get('nombre', ''), den.get('cuitPersona', ''))
    elif empresa:
        base = 'la empresa %s CUIT %s, dedicada a %s' % (
            den.get('empresa', ''), den.get('cuitEmpresa', ''), den.get('rubro', '')
        )
    else:
        base = 'la parte denunciada'

    if domicilio or ciudad:
        base += ', con domicilio en %s%s' % (
            domicilio, (' de la ciudad de %s' % ciudad) if ciudad else ''
        )
    if email:
        base += ', correo electrónico %s' % email
    return base

def armar_bloque_registro(relacion, denunciado):
    texto = (
        "Que durante toda la relación laboral presté servicios de manera personal, "
        "habitual y bajo dependencia %s, sin que mi vínculo laboral "
        "se encontrara debidamente registrado ante los organismos correspondientes, "
        "en violación a la normativa laboral vigente."
    ) % de_denunciado(denunciado)
    extra = (relacion.get('registroParcial') or '').strip()
    if extra:
        texto += '\n\n' + extra
    return texto

# Relato fijo para el motivo "impedimento de ingreso" (el empleador no deja
# volver a trabajar). Otros motivos (accidente, falta de pago, etc.) van a
# necesitar su propio bloque el dia que se agreguen.
def armar_bloque_relato_impedimento(fecha, denunciado):
    return ("""Qué la relación laboral continuó desarrollándose en los términos precedentemente expuestos hasta el día %s, fecha a partir de la cual se me impidió de manera injustificada el ingreso a mi lugar habitual de trabajo, sin mediar causa válida ni comunicación previa, circunstancia que me imposibilitó continuar prestando tareas con normalidad.

Que dicho impedimento se materializó mediante la negativa a permitirme el ingreso y a asignarme tareas propias de mi categoría y especialidad, pese a presentarme en el lugar y horario habitual de trabajo, configurando una situación de hecho que vulnera mis derechos laborales.

Qué, asimismo, como expresé anteriormente, el empleador nunca regularizó correctamente mi situación laboral, encontrándome sin la debida registración ante los organismos correspondientes, pese a haber prestado tareas de manera personal, habitual y bajo relación de dependencia.

Que, ante la falta de respuesta y la persistencia de dicha situación, me vi obligado a intimar fehacientemente a fin de que se aclare mi situación laboral, se regularice mi registración y se me permita retomar mis tareas habituales, bajo apercibimiento de accionar conforme a derecho.

Que, pese a mis reiterados reclamos verbales, %s persistió en su conducta de impedirme trabajar, colocándome en una situación de absoluta incertidumbre respecto de mi continuidad laboral.""") % (fecha, articulo_denunciado(denunciado))

BLOQUE_CIERRE_RECLAMO = """Que conforme lo acontecido, las injurias ocasionadas, intimo a que me abone Liquidación Final, conforme la verdadera relación fáctica laboral y haberes adeudados y diferencias de haberes, SAC y Vacaciones No Gozadas por el plazo de prescripción, Indemnización por despido sin causa, antigüedad (art. 245 LCT), sustitutiva de preaviso (art. 232 LCT), Integración mes de despido (art. 233 LCT) y haga entrega de Certificaciones de Servicios y Remuneraciones estipuladas en Art. 80 LCT, sirviendo el presente como emplazamiento de dicha entrega."""

# Cada item de correspondencia se arma con una formula fija, en orden
# cronologico; el contenido transcripto (cita textual del telegrama) nunca
# pasa por la IA en este paso, solo se inserta tal cual lo cargo/leyo el
# usuario.
def armar_telegrama_nuestro(item):
    fecha = (item.get('fechaEnvio') or '').strip()
    numero = (item.get('numero') or '').strip()
    contenido = (item.get('contenido') or '').strip()
    recepcion = (item.get('fechaRecepcion') or '').strip()

    if numero:
        intro = 'Que, con fecha %s, remití TCL CD Nº %s, en los siguientes términos:' % (fecha, numero)
    else:
        intro = 'Que, con fecha %s, remití TCL, en los siguientes términos:' % fecha

    texto = intro + '\n\n"%s"' % contenido
    if recepcion:
        texto += '\n\nDicha intimación fue debidamente recibida por la denunciada con fecha %s conforme constancia del Correo Argentino.' % recepcion
    return texto

def armar_respuesta_empresa(resp):
    fecha = (resp.get('fechaEnvio') or '').strip()
    numero = (resp.get('numero') or '').strip()
    contenido = (resp.get('contenido') or '').strip()
    recepcion = (resp.get('fechaRecepcion') or '').strip()

    texto = 'Que, no obstante lo intimado, con fecha %s la parte denunciada remitió CD' % fecha
    if numero:
        texto += ' Nº %s' % numero
    if recepcion:
        texto += ', recibida por mi parte el día %s' % recepcion
    texto += ', mediante la cual pretendió responder los reclamos, en los siguientes términos:\n\n"%s"' % contenido
    return texto

def armar_item_correspondencia(item, alguna_respuesta_ref):
    if item.get('emisor') == 'empresa':
        alguna_respuesta_ref[0] = True
        return [armar_respuesta_empresa(item)]

    partes = [armar_telegrama_nuestro(item)]
    if item.get('huboRespuesta') and item.get('respuesta'):
        partes.append(armar_respuesta_empresa(item['respuesta']))
        alguna_respuesta_ref[0] = True
    return partes

def armar_bloque_correspondencia(correspondencia, denunciado):
    partes = []
    alguna_respuesta_ref = [False]
    for item in correspondencia:
        partes.extend(armar_item_correspondencia(item, alguna_respuesta_ref))
    if not alguna_respuesta_ref[0]:
        partes.append('Al día de la fecha %s no ha contestado ninguna de las intimaciones remitidas por esta parte.' % articulo_denunciado(denunciado))
    return '\n\n'.join(partes)

def armar_encabezado_cliente(cliente):
    nombre = (cliente.get('nombre') or '').strip()
    dni = (cliente.get('dni') or '').strip()
    cuit = (cliente.get('cuit') or '').strip()
    edad = (cliente.get('edad') or '').strip()
    estado_civil = (cliente.get('estadoCivil') or '').strip()
    nacionalidad = (cliente.get('nacionalidad') or '').strip()
    situacion = (cliente.get('situacionLaboral') or '').strip()
    domicilio = (cliente.get('domicilio') or '').strip()
    ciudad = (cliente.get('ciudad') or '').strip()
    telefono = (cliente.get('telefono') or '').strip()
    email = (cliente.get('email') or '').strip()
    bloque_rep = BLOQUES_REPRESENTACION.get(cliente.get('representacion'), BLOQUES_REPRESENTACION['ciardiello'])

    texto = nombre
    if dni:
        texto += ' DNI %s' % dni
    texto += ', CUIL %s' % cuit
    if edad:
        texto += ', %s años de edad' % edad
    texto += ', de estado civil %s, de nacionalidad %s' % (estado_civil, nacionalidad)
    if situacion:
        texto += ', actualmente %s' % situacion
    texto += (
        ', con domicilio real en %s, %s. Celular %s, correo electrónico: %s, %s, '
        'ambos de esta ciudad, ante S.S. respetuosamente comparezco y digo:'
    ) % (domicilio, ciudad, telefono, email, bloque_rep)
    return texto

def armar_escrito_denuncia_trabajo(cliente, denunciado, relacion, correspondencia, parrafo_tareas):
    encabezado_cliente = armar_encabezado_cliente(cliente)
    bloque_den = armar_bloque_denunciado(denunciado)
    bloque_registro = armar_bloque_registro(relacion, denunciado)
    bloque_relato = armar_bloque_relato_impedimento(relacion.get('fechaDespido', ''), denunciado)
    bloque_corr = armar_bloque_correspondencia(correspondencia, denunciado)

    return """FORMULA DENUNCIA.

Sr. Director de Reclamaciones Individuales
Departamento Provincial del Trabajo.
S___________________/____________D

%s

Que vengo a formular denuncia en contra de %s, comparecemos y decimos:

Que ingresé a trabajar bajo las órdenes %s %s, cumpliendo tareas acordes con la categoría %s.

Que mi jornada de trabajo era %s, %s.

%s

%s

%s

%s

%s

Por los motivos expuestos, se solicita la intervención de esta Autoridad de Aplicación.

En definitiva, ante el incumplimiento de la normativa vigente por parte del Empleador, solicitó que el mismo sea citado, quien deberá concurrir con toda la documentación laboral, legajo personal, además de las constancias de pago de los aportes y contribuciones sindicales, sociales y previsionales, y los importes correspondientes a Liquidación Final e indemnizaciones de ley, a cuyo fin deberá fijar día y hora de audiencia de conciliación. Sin otro particular. Saludo a Ud. muy atte.""" % (
        encabezado_cliente,
        bloque_den,
        de_denunciado(denunciado), relacion.get('fechaIngreso', ''), relacion.get('categoria', ''),
        relacion.get('jornada', ''), relacion.get('horarios', ''),
        parrafo_tareas,
        bloque_registro,
        bloque_relato,
        bloque_corr,
        BLOQUE_CIERRE_RECLAMO
    )

# --- EXPORTAR A WORD (.docx) ---
# Toma el texto ya generado (y eventualmente editado a mano en el textarea)
# y lo pasa a un documento .docx con tipografia, tamano, justificado y
# sangria acordes al modelo aprobado por el estudio, listo para retocar en
# Word antes de exportar a PDF.

FIRMAS_ABOGADOS = {
    'ciardiello': [('Ma. Paula Ciardiello', 'Abogada', 'M.P. 1-41227')],
    'casih': [('Pablo Casih', 'Abogado', 'M.P. 1-31142')],
    'ambos': [('Ma. Paula Ciardiello', 'Abogada', 'M.P. 1-41227'), ('Pablo Casih', 'Abogado', 'M.P. 1-31142')],
}

def _agregar_firma(celda, nombre, subt1='', subt2=''):
    celda.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    celda.paragraphs[0].add_run('_______________________')
    p_nombre = celda.add_paragraph(nombre)
    p_nombre.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if p_nombre.runs:
        p_nombre.runs[0].bold = True
    for sub in (subt1, subt2):
        if sub:
            p_sub = celda.add_paragraph(sub)
            p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

def construir_docx_denuncia(texto, cliente):
    doc = Document()

    section = doc.sections[0]
    section.left_margin = Cm(3)
    section.right_margin = Cm(2.5)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)

    estilo = doc.styles['Normal']
    estilo.font.name = 'Times New Roman'
    estilo.font.size = Pt(12)
    estilo.paragraph_format.line_spacing = 1.15
    estilo.paragraph_format.space_after = Pt(0)

    marcador_cierre = 'Sin otro particular. Saludo a Ud. muy atte.'
    bloques = [b.strip('\n') for b in texto.split('\n\n') if b.strip()]

    for bloque in bloques:
        if bloque.strip() == 'FORMULA DENUNCIA.':
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run('FORMULA DENUNCIA.').bold = True
            continue

        if bloque.startswith('Sr. Director'):
            p = doc.add_paragraph()
            for i, linea in enumerate(bloque.split('\n')):
                if i > 0:
                    p.add_run().add_break()
                p.add_run(linea)
            continue

        if bloque.startswith('"') and bloque.endswith('"'):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.left_indent = Cm(1.5)
            p.paragraph_format.right_indent = Cm(1.5)
            p.add_run(bloque).italic = True
            continue

        if marcador_cierre in bloque:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Cm(1.25)
            antes, _, despues = bloque.partition(marcador_cierre)
            if antes:
                p.add_run(antes)
            p.add_run(marcador_cierre + despues).bold = True
            continue

        p = doc.add_paragraph(bloque)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = Cm(1.25)

    doc.add_paragraph()
    doc.add_paragraph()

    abogados = FIRMAS_ABOGADOS.get(cliente.get('representacion'), FIRMAS_ABOGADOS['ciardiello'])
    tabla = doc.add_table(rows=1, cols=1 + len(abogados))
    celdas = tabla.rows[0].cells

    nombre_cliente = (cliente.get('nombre') or '').strip()
    dni_cliente = (cliente.get('dni') or '').strip()
    _agregar_firma(celdas[0], nombre_cliente, ('DNI %s' % dni_cliente) if dni_cliente else '')
    for i, (nombre_ab, titulo_ab, mp_ab) in enumerate(abogados):
        _agregar_firma(celdas[1 + i], nombre_ab, titulo_ab, mp_ab)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

@app.route('/generar-escrito-docx', methods=['POST', 'OPTIONS'])
def generar_escrito_docx():
    if request.method == 'OPTIONS':
        resp = jsonify({})
        resp.headers['Access-Control-Allow-Origin'] = '*'
        resp.headers['Access-Control-Allow-Methods'] = 'POST, OPTIONS'
        resp.headers['Access-Control-Allow-Headers'] = 'Content-Type'
        return resp, 200

    body = request.get_json(silent=True) or {}
    texto = (body.get('texto') or '').strip()
    cliente = body.get('cliente') or {}

    if not texto:
        resp = jsonify({"status": "error", "detalle": "Falta el texto del escrito."})
        resp.headers['Access-Control-Allow-Origin'] = '*'
        return resp, 400

    try:
        buffer = construir_docx_denuncia(texto, cliente)
    except Exception as e:
        print("Error generando docx:", str(e))
        resp = jsonify({"status": "error", "detalle": "No se pudo generar el archivo Word."})
        resp.headers['Access-Control-Allow-Origin'] = '*'
        return resp, 500

    nombre_archivo = 'Denuncia - %s.docx' % ((cliente.get('nombre') or 'escrito').strip() or 'escrito')

    resp = send_file(
        buffer,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        as_attachment=True,
        download_name=nombre_archivo
    )
    resp.headers['Access-Control-Allow-Origin'] = '*'
    return resp

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
    denunciado = body.get('denunciado') or {}
    relacion = body.get('relacion') or {}
    correspondencia = body.get('correspondencia') or []

    def error(msg, code=400):
        resp = jsonify({"status": "error", "detalle": msg})
        resp.headers['Access-Control-Allow-Origin'] = '*'
        return resp, code

    if tipo != 'denuncia_trabajo':
        return error('Tipo de escrito no reconocido.')
    if not cliente.get('nombre'):
        return error('Faltan los datos del cliente.')
    if not denunciado.get('esPersona') and not denunciado.get('esEmpresa'):
        return error('Marcá si el denunciado es persona física, empresa, o ambas.')
    if not (relacion.get('tareas') or '').strip():
        return error('Falta la descripción de las tareas.')
    if not correspondencia:
        return error('Agregá al menos un telegrama o carta documento.')
    for item in correspondencia:
        if item.get('huboRespuesta') and not (item.get('respuesta') or {}).get('contenido', '').strip():
            return error('Marcaste que hubo respuesta de la empresa en un telegrama pero falta el contenido de esa respuesta.')

    api_key = os.environ.get('ANTHROPIC_API_KEY')
    if not api_key:
        return error('El servidor no tiene configurada la clave de la API de Claude.', 500)

    system_prompt = (
        "Sos un asistente de redacción para un estudio jurídico laboralista de Córdoba, Argentina. "
        "Vas a recibir una descripción informal de las tareas que hacía un trabajador y tenés que "
        "convertirla en UN párrafo formal, en primera persona, con el registro de una denuncia "
        "laboral ante el Ministerio de Trabajo, arrancando de forma natural (por ejemplo: 'Que las "
        "tareas desempeñadas por mi parte consistían en...').\n\n"
        "Reglas:\n"
        "- Nunca inventes datos que no estén en el texto que te pasaron.\n"
        "- Devolvé únicamente el párrafo final, sin comentarios antes o después."
    )
    user_prompt = "TEXTO INFORMAL DE LAS TAREAS:\n%s" % relacion.get('tareas', '')

    try:
        client = anthropic.Anthropic(api_key=api_key)
        response = client.messages.create(
            model="claude-sonnet-5",
            max_tokens=1024,
            system=system_prompt,
            messages=[{"role": "user", "content": user_prompt}]
        )
        parrafo_tareas = "".join([b.text for b in response.content if b.type == "text"]).strip()
    except Exception as e:
        print("Error generando parrafo de tareas:", str(e))
        return error('No se pudo generar el escrito. Intentá de nuevo en un momento.', 500)

    texto_final = armar_escrito_denuncia_trabajo(cliente, denunciado, relacion, correspondencia, parrafo_tareas)

    resp = jsonify({"status": "ok", "texto": texto_final})
    resp.headers['Access-Control-Allow-Origin'] = '*'
    return resp, 200

# Lee un telegrama/CD adjunto (foto o PDF) y extrae fecha, numero de
# seguimiento y contenido. Es una transcripcion asistida: el frontend siempre
# la muestra en un campo editable para que se revise contra el original antes
# de usarla en un escrito.
@app.route('/extraer-telegrama', methods=['POST', 'OPTIONS'])
def extraer_telegrama():
    if request.method == 'OPTIONS':
        resp = jsonify({})
        resp.headers['Access-Control-Allow-Origin'] = '*'
        resp.headers['Access-Control-Allow-Methods'] = 'POST, OPTIONS'
        resp.headers['Access-Control-Allow-Headers'] = 'Content-Type'
        return resp, 200

    def error(msg, code=400):
        resp = jsonify({"status": "error", "detalle": msg})
        resp.headers['Access-Control-Allow-Origin'] = '*'
        return resp, code

    body = request.get_json(silent=True) or {}
    archivo_b64 = body.get('archivo')
    media_type = body.get('mediaType') or 'application/pdf'

    if not archivo_b64:
        return error('Falta el archivo.')

    api_key = os.environ.get('ANTHROPIC_API_KEY')
    if not api_key:
        return error('El servidor no tiene configurada la clave de la API de Claude.', 500)

    if media_type == 'application/pdf':
        content_block = {"type": "document", "source": {"type": "base64", "media_type": media_type, "data": archivo_b64}}
    elif media_type.startswith('image/'):
        content_block = {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": archivo_b64}}
    else:
        return error('Formato de archivo no soportado. Subí una foto o un PDF.')

    system_prompt = (
        "Sos un asistente que transcribe telegramas y cartas documento (CD) de Correo Argentino "
        "para un estudio jurídico. Te paso una imagen o PDF del documento. Extraé exactamente: "
        "la fecha de envío (formato DD/MM/AAAA si figura), el número de seguimiento/CD, y el "
        "contenido textual completo del mensaje enviado (no incluyas datos del formulario de "
        "Correo Argentino que no sean parte del texto del mensaje, como códigos de barra). "
        "Transcribí el contenido exactamente como está escrito, sin resumir ni corregir. Si algún "
        "dato no se puede leer con confianza, dejalo vacío en vez de adivinar."
    )

    try:
        client = anthropic.Anthropic(api_key=api_key)
        response = client.messages.create(
            model="claude-sonnet-5",
            max_tokens=2048,
            system=system_prompt,
            output_config={"format": {
                "type": "json_schema",
                "schema": {
                    "type": "object",
                    "properties": {
                        "fecha": {"type": "string"},
                        "numero": {"type": "string"},
                        "contenido": {"type": "string"}
                    },
                    "required": ["fecha", "numero", "contenido"],
                    "additionalProperties": False
                }
            }},
            messages=[{"role": "user", "content": [
                content_block,
                {"type": "text", "text": "Extraé fecha, número de seguimiento y contenido de este documento."}
            ]}]
        )
        texto_json = "".join([b.text for b in response.content if b.type == "text"])
        datos = json.loads(texto_json)
    except Exception as e:
        print("Error leyendo telegrama:", str(e))
        return error('No se pudo leer el archivo. Probá con otra foto/escaneo, o cargá los datos a mano.', 500)

    resp = jsonify({"status": "ok", "fecha": datos.get('fecha', ''), "numero": datos.get('numero', ''), "contenido": datos.get('contenido', '')})
    resp.headers['Access-Control-Allow-Origin'] = '*'
    return resp, 200

@app.route('/', methods=['GET'])
def health():
    return "Servidor del Bot activo y funcionando correctamente.", 200

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
