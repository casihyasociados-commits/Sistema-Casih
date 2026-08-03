import os
import json
from flask import Flask, request, jsonify
import firebase_admin
from firebase_admin import credentials, firestore
import requests

app = Flask(__name__)

# Inicializar Firebase
if not firebase_admin._apps:
    cred_json = os.environ.get('FIREBASE_SERVICE_ACCOUNT_JSON')
    if cred_json:
        cred_dict = json.loads(cred_json)
        cred = credentials.Certificate(cred_dict)
        firebase_admin.initialize_app(cred)
    else:
        firebase_admin.initialize_app()

db = firestore.client()

GREEN_API_INSTANCE = os.environ.get('GREEN_API_INSTANCE')
GREEN_API_TOKEN = os.environ.get('GREEN_API_TOKEN')

def responder_whatsapp(chat_id, texto):
    if not GREEN_API_INSTANCE or not GREEN_API_TOKEN:
        return
    url = f"https://api.green-api.com/waInstance{GREEN_API_INSTANCE}/sendMessage/{GREEN_API_TOKEN}"
    payload = {"chatId": chat_id, "message": texto}
    headers = {'Content-Type': 'application/json'}
    requests.post(url, json=payload, headers=headers)

@app.route('/webhook', methods=['POST'])
def webhook():
    data = request.get_json() or {}
    
    # Extraer mensaje y sender
    message_data = data.get('messageData', {})
    text_message_data = message_data.get('textMessageData', {})
    mensaje = text_message_data.get('textMessage', '').strip()
    sender_data = data.get('senderData', {})
    chat_id = sender_data.get('chatId')

    if not mensaje or not chat_id:
        return jsonify({"status": "ignored"}), 200

    # Normalizar texto (ej: "Gasto 1500 taxi" -> ["gasto", "1500", "taxi"])
    partes = mensaje.split(' ', 2)
    comando = partes[0].lower()

    if comando in ['gasto', 'ingreso'] and len(partes) >= 2:
        try:
            # Limpiar el monto de caracteres como '$' o ','
            monto_raw = partes[1].replace('$', '').replace(',', '.')
            monto = float(monto_raw)
            concepto = partes[2] if len(partes) > 2 else 'Sin concepto'

            coleccion = 'gastos' if comando == 'gasto' else 'ingresos'
            
            # Guardar en Firestore
            db.collection(coleccion).add({
                'monto': monto,
                'concepto': concepto,
                'fecha': firestore.SERVER_TIMESTAMP,
                'origen': 'WhatsApp'
            })

            respuesta = f"✅ *{comando.capitalize()} registrado:* ${monto:.2f}\n📝 *Concepto:* {concepto}"
            responder_whatsapp(chat_id, respuesta)
            return jsonify({"status": "success"}), 200

        except ValueError:
            responder_whatsapp(chat_id, "❌ Error: El monto ingresado no es un número válido. Ejemplo: `gasto 1500 taxi`")
            return jsonify({"status": "error", "message": "Monto invalido"}), 400
        except Exception as e:
            responder_whatsapp(chat_id, f"❌ Error al guardar en la base de datos: {str(e)}")
            return jsonify({"status": "error", "message": str(e)}), 500

    return jsonify({"status": "ignored"}), 200

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)
