from flask import Flask, request, jsonify
import os
import datetime
from pytz import timezone
import re
import io
import requests
import json  # Importamos json para manejar las credenciales de Google
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import docx
import openpyxl
import PyPDF2
from google.oauth2 import service_account
from flask_cors import CORS  # Importamos CORS para permitir solicitudes desde diferentes orígenes

app = Flask(__name__)
CORS(app)  # Configuramos CORS en la aplicación

# Ruta para la página de inicio
@app.route('/')
def index():
    return "¡Bienvenido a la aplicación de transcripción y carga de documentos!"

# Configurar los SCOPES de Google Drive, Docs y Sheets
SCOPES = ['https://www.googleapis.com/auth/documents',
          'https://www.googleapis.com/auth/drive',
          'https://www.googleapis.com/auth/spreadsheets.readonly']

# Función para autenticarse usando solo la cuenta de servicio
def service_account_login():
    credentials = None
    if 'GOOGLE_CREDENTIALS' in os.environ:
        # Autenticación mediante la cuenta de servicio configurada en Heroku
        try:
            credentials_info = json.loads(os.getenv('GOOGLE_CREDENTIALS'))
            credentials = service_account.Credentials.from_service_account_info(
                credentials_info, scopes=SCOPES)
        except Exception as e:
            print(f"Error al cargar las credenciales de servicio: {e}")
            raise EnvironmentError("Las credenciales de servicio de Google no están configuradas correctamente.")
    else:
        raise EnvironmentError("Las credenciales de servicio de Google no están configuradas en las variables de entorno.")

    drive_service = build('drive', 'v3', credentials=credentials)
    docs_service = build('docs', 'v1', credentials=credentials)
    sheets_service = build('sheets', 'v4', credentials=credentials)
    return drive_service, docs_service, sheets_service

# Función para obtener la ruta completa de una carpeta
def obtener_ruta_completa(service, file_id):
    parents = service.files().get(
        fileId=file_id, fields="parents").execute().get('parents', [])
    if not parents:
        return "Raíz"

    path = []
    while parents:
        parent_id = parents[0]
        parent = service.files().get(
            fileId=parent_id, fields="name, parents").execute()
        path.insert(0, parent['name'])  # Insertar al inicio para formar la ruta
        parents = parent.get('parents', [])

    return '/'.join(path)

# Buscar archivos por nombre en Google Drive y mostrar detalles para elegir
def buscar_archivos_por_nombre(service, file_name):
    results = service.files().list(
        q=f"name contains '{file_name}'",
        spaces='drive',
        fields='files(id, name, mimeType, modifiedTime, parents)',
        pageSize=10
    ).execute()

    items = results.get('files', [])
    if not items:
        print('No se encontraron archivos.')
        return []

    archivos = []
    for item in items:
        ruta_completa = obtener_ruta_completa(service, item['id'])
        archivo = {
            'id': item['id'],
            'nombre': item['name'],
            'tipo_mime': item['mimeType'],
            'modificado': item['modifiedTime'],
            'ruta': ruta_completa
        }
        archivos.append(archivo)
    return archivos

# Función para procesar Google Sheets
def procesar_gsheet(sheets_service, file_id):
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=file_id, range="A1:Z1000").execute()
    values = result.get('values', [])
    print(f"Contenido extraído de Google Sheets: {values[:5]}...")
    return {"nombre_documento": "Google Sheet", "contenido": values}

# Función para procesar documentos PDF
def procesar_pdf(service, file_id):
    request = service.files().get_media(fileId=file_id)
    file_stream = io.BytesIO()
    downloader = MediaIoBaseDownload(file_stream, request)

    done = False
    while not done:
        status, done = downloader.next_chunk()

    file_stream.seek(0)
    reader = PyPDF2.PdfReader(file_stream)
    full_text = ""
    for page_num in range(len(reader.pages)):
        page = reader.pages[page_num]
        full_text += page.extract_text()

    print(f"Contenido extraído del archivo PDF: {full_text[:100]}...")
    return {"nombre_documento": "Archivo PDF", "contenido": full_text}

# Función para procesar documentos Word
def procesar_word(service, file_id):
    request = service.files().get_media(fileId=file_id)
    file_stream = io.BytesIO()
    downloader = MediaIoBaseDownload(file_stream, request)

    done = False
    while not done:
        status, done = downloader.next_chunk()

    file_stream.seek(0)
    doc = docx.Document(file_stream)
    full_text = ""
    for para in doc.paragraphs:
        full_text += para.text + "\n"

    print(f"Contenido extraído del archivo Word: {full_text[:100]}...")
    return {"nombre_documento": "Archivo Word", "contenido": full_text}

# Función para procesar documentos Excel
def procesar_excel(service, file_id):
    request = service.files().get_media(fileId=file_id)
    file_stream = io.BytesIO()
    downloader = MediaIoBaseDownload(file_stream, request)

    done = False
    while not done:
        status, done = downloader.next_chunk()

    file_stream.seek(0)
    workbook = openpyxl.load_workbook(file_stream)
    sheet = workbook.active

    content = []
    for row in sheet.iter_rows(values_only=True):
        content.append(list(row))

    print(f"Contenido extraído del archivo Excel: {content[:5]}...")
    return {"nombre_documento": "Archivo Excel", "contenido": content}

# Función para cargar documento por nombre
def cargar_documento(drive_service, nombre_documento, seleccion_usuario=None):
    try:
        print(f"Intentando cargar el documento: {nombre_documento}")

        _, docs_service, sheets_service = service_account_login()
        archivos = buscar_archivos_por_nombre(drive_service, nombre_documento)

        if not archivos:
            print(f"No se encontró ningún archivo con el nombre: {nombre_documento}")
            return {"mensaje": f"No se encontró el documento con el nombre {nombre_documento}"}

        if len(archivos) > 1 and seleccion_usuario is None:
            # Mostrar opciones al usuario con detalles
            mensaje = "Múltiples archivos encontrados:\n"
            for idx, archivo in enumerate(archivos, 1):
                mensaje += f"{idx}. {archivo['nombre']} (MIME: {archivo['tipo_mime']}, Modificado: {archivo['modificado']}, Ruta: {archivo['ruta']})\n"
            print(mensaje)
            return {"mensaje": mensaje, "archivos": archivos}

        # Si se pasa selección del usuario, usar el archivo seleccionado
        if seleccion_usuario is not None and 1 <= seleccion_usuario <= len(archivos):
            archivo_seleccionado = archivos[seleccion_usuario - 1]
            file_id = archivo_seleccionado['id']
            mime_type = archivo_seleccionado['tipo_mime']
        else:
            file_id = archivos[0]['id']
            mime_type = archivos[0]['tipo_mime']

        print(f"Archivo seleccionado: ID={file_id}, MIME={mime_type}")

        # Verificar si es un Google Sheet
        if mime_type == 'application/vnd.google-apps.spreadsheet':
            return procesar_gsheet(sheets_service, file_id)

        # Verificar si es un documento de Google Docs
        elif mime_type == 'application/vnd.google-apps.document':
            document = docs_service.documents().get(documentId=file_id).execute()

            full_text = ""
            content_elements = document.get('body').get('content')
            for element in content_elements:
                if 'paragraph' in element:
                    for text_run in element['paragraph']['elements']:
                        if 'textRun' in text_run:
                            full_text += text_run['textRun']['content']

            print(f"Contenido extraído del documento: {full_text[:100]}...")
            return {"nombre_documento": nombre_documento, "contenido": full_text}

        # Verificar si es un archivo Excel (.xlsx)
        elif mime_type == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
            return procesar_excel(drive_service, file_id)

        # Si es un archivo de Word (.docx), procesarlo directamente
        elif mime_type == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
            return procesar_word(drive_service, file_id)

        # Si es un archivo PDF
        elif mime_type == 'application/pdf':
            return procesar_pdf(drive_service, file_id)

        else:
            return {"nombre_documento": nombre_documento, "mensaje": f"El archivo tiene un tipo MIME desconocido: {mime_type}"}
    except Exception as e:
        print(f"Error al cargar el documento: {e}")
        return {"mensaje": f"Error al cargar el documento: {e}"}

# Buscar archivo por nombre en Google Drive (para guardar transcripciones)
def find_file_by_name(service, file_name):
    try:
        results = service.files().list(
            q=f"name='{file_name}' and mimeType='application/vnd.google-apps.document'",
            spaces='drive',
            fields='files(id, name)',
            pageSize=10
        ).execute()
        items = results.get('files', [])
        if not items:
            print('No se encontraron archivos.')
            return None
        else:
            return items[0]['id']
    except Exception as e:
        print(f"Error al buscar el archivo: {e}")
        return None

# Crear un nuevo documento si no existe y moverlo a la carpeta especificada
def create_new_document(drive_service, docs_service, file_name, folder_id):
    try:
        new_document = docs_service.documents().create(body={
            'title': file_name
        }).execute()

        new_file_id = new_document['documentId']

        drive_service.files().update(
            fileId=new_file_id,
            addParents=folder_id,
            fields='id, parents'
        ).execute()

        print(f"Documento creado: {file_name} (ID: {new_file_id}) en la carpeta {folder_id}")

        return new_file_id
    except Exception as e:
        print(f"Error al crear el documento: {e}")
        return None

# Actualizar el documento con la nueva transcripción y la fecha/hora
def update_document(docs_service, file_id, new_content):
    try:
        document = docs_service.documents().get(documentId=file_id).execute()
        content_elements = document.get('body').get('content')
        end_index = 1
        for element in content_elements:
            if 'endIndex' in element:
                end_index = element['endIndex']

        date_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        date_time = datetime.datetime.now(timezone('Europe/Madrid')).strftime("%Y-%m-%d %H:%M:%S")
        content_to_add = f"\nFecha: {date_time}\nTranscripción:\n{new_content}\n"
        requests_body = [
            {
                'insertText': {
                    'location': {
                        'index': end_index - 1,
                    },
                    'text': content_to_add
                }
            }
        ]
        docs_service.documents().batchUpdate(
            documentId=file_id, body={'requests': requests_body}).execute()
    except Exception as e:
        print(f"Error al actualizar el documento: {e}")

# Actualizar el documento con la transcripción completa y formateada
def update_document_con_formato(docs_service, file_id, new_content):
    try:
        document = docs_service.documents().get(documentId=file_id).execute()
        content_elements = document.get('body').get('content')

        # Obtener el índice final del documento para asegurarnos de no exceder los límites
        end_index = 1
        for element in content_elements:
            if 'endIndex' in element:
                end_index = element['endIndex']

        # Asegurarnos de no exceder el índice del documento
        if end_index > 1:
            end_index -= 1

        date_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        date_time = datetime.datetime.now(timezone('Europe/Madrid')).strftime("%Y-%m-%d %H:%M:%S")

        # Formatear la transcripción correctamente, insertando el contenido real
        contenido_formateado = f"\n\nFecha: {date_time}\n\n{new_content}\n\n"

        # Preparar las solicitudes para insertar el contenido
        requests_body = [
            {
                'insertText': {
                    'location': {
                        'index': end_index,
                    },
                    'text': contenido_formateado
                }
            },
            {
                'updateTextStyle': {
                    'range': {
                        'startIndex': end_index,
                        'endIndex': end_index + len(f"\n\nFecha: {date_time}\n\n")
                    },
                    'textStyle': {
                        'bold': True
                    },
                    'fields': 'bold'
                }
            },
            {
                'updateTextStyle': {
                    'range': {
                        'startIndex': end_index + len(f"\n\nFecha: {date_time}\n\n"),
                        'endIndex': end_index + len(contenido_formateado)
                    },
                    'textStyle': {
                        'bold': False
                    },
                    'fields': 'bold'
                }
            }
        ]

        # Aplicar las actualizaciones al documento
        docs_service.documents().batchUpdate(documentId=file_id, body={'requests': requests_body}).execute()
    except Exception as e:
        print(f"Error al actualizar el documento con formato: {e}")

# Función para detectar el comando y enviar la transcripción
def ejecutar_guardado(comando, transcription_text):
    try:
        print(f"Procesando comando: {comando}, con transcripción: {transcription_text}")
        match = re.search(r"guardar transcripción para (\w+)", comando, re.IGNORECASE)

        if match:
            nombre_nino = match.group(1)
            data = {
                "nombre_nino": nombre_nino,
                "transcripcion": transcription_text
            }
            print(f"Enviando datos: {data}")
            response = requests.post(
                "https://narayan-seguiments-2019-2ad53e4ee809.herokuapp.com/guardar_transcripcion", json=data)
            if response.status_code == 200:
                print(f"Transcripción guardada correctamente para {nombre_nino}.")
            else:
                print(f"Error al guardar la transcripción: {response.status_code} - {response.text}")
        else:
            print("Comando no reconocido. No se pudo extraer el nombre del niño.")
    except Exception as e:
        print(f"Error al ejecutar guardado: {e}")

# Ruta para procesar transcripciones y el comando
@app.route('/procesar_transcripcion', methods=['POST'])
def procesar_transcripcion():
    try:
        # Recibir datos del comando y transcripción
        data = request.get_json(force=True)
        if not data:
            return jsonify({"status": "error", "message": "No se recibieron datos JSON válidos."}), 400

        print(f"Solicitud recibida en /procesar_transcripcion: {data}")
        comando = data.get('comando')
        transcription_text = data.get('transcripcion')

        if not comando or not transcription_text:
            return jsonify({"status": "error", "message": "Faltan campos requeridos 'comando' y/o 'transcripcion'."}), 400

        # Procesar el comando y enviar la transcripción
        ejecutar_guardado(comando, transcription_text)

        return jsonify({"status": "success", "message": "Transcripción procesada correctamente."})
    except Exception as e:
        print(f"Error procesando la solicitud en /procesar_transcripcion: {e}")
        return jsonify({"status": "error", "message": f"Error procesando la transcripción: {e}"}), 500

# Ruta de la API para guardar la transcripción con formato
@app.route('/guardar_transcripcion', methods=['POST'])
def guardar_transcripcion():
    try:
        data = request.get_json(force=True)
        if not data:
            return jsonify({"status": "error", "message": "No se recibieron datos JSON válidos."}), 400

        print(f"Solicitud recibida en /guardar_transcripcion: {data}")
        nombre_nino = data.get('nombre_nino')
        transcription_text = data.get('transcripcion')

        if not nombre_nino or not transcription_text:
            return jsonify({"status": "error", "message": "Faltan campos requeridos 'nombre_nino' y/o 'transcripcion'."}), 400

        # Autenticar Google Drive y Docs
        drive_service, docs_service, _ = service_account_login()

        # Nombre del archivo será el nombre del niño
        file_name = f"informe {nombre_nino} Seguimiento"

        # Buscar el documento en Google Drive
        file_id = find_file_by_name(drive_service, file_name)

        # Si no se encuentra, crear uno nuevo
        if not file_id:
            folder_id = '10EnZp-WPh4o-Dwvpojl3djT9iCx4ZwHE'  # Reemplaza con la ID de tu carpeta
            file_id = create_new_document(drive_service, docs_service, file_name, folder_id)

        # Actualizar el documento con la transcripción y formato
        update_document_con_formato(docs_service, file_id, transcription_text)
        return jsonify({"status": "success", "message": f"Transcripción guardada para {nombre_nino}."})
    except Exception as e:
        print(f"Error procesando la solicitud en /guardar_transcripcion: {e}")
        return jsonify({"status": "error", "message": f"Error guardando la transcripción: {e}"}), 500

# Ruta para procesar la carga de un documento
@app.route('/cargar_documento', methods=['POST'])
def cargar_documento_route():
    try:
        data = request.get_json(force=True)
        if not data:
            return jsonify({"status": "error", "message": "No se recibieron datos JSON válidos."}), 400

        nombre_documento = data.get('nombre_documento')
        seleccion_usuario = data.get('seleccion_usuario')  # Número seleccionado por el usuario (si lo hay)

        if not nombre_documento:
            return jsonify({"status": "error", "message": "El campo 'nombre_documento' es requerido."}), 400

        drive_service, _, _ = service_account_login()
        documento = cargar_documento(drive_service, nombre_documento, seleccion_usuario)
        return jsonify(documento)
    except Exception as e:
        print(f"Error procesando la solicitud en /cargar_documento: {e}")
        return jsonify({"status": "error", "message": f"Error cargando el documento: {e}"}), 500

# Configuración para que Flask use el puerto que Heroku asigna
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=True, host='0.0.0.0', port=port)
