import os
import base64
import mimetypes
import time
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def get_service():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    try:
        service = build('gmail', 'v1', credentials=creds)
        return service
    except HttpError as error:
        print(f'An error occurred: {error}')
        return None

def descargar_adjuntos_validos(service, user_id, msg_id):
    archivos_descargados = []
    tipos_validos = ['application/pdf', 'image/jpeg', 'image/png']
    try:
        message = service.users().messages().get(userId=user_id, id=msg_id, format='metadata', metadataHeaders=['From']).execute()
        # Obtener el emisor del correo
        headers = message.get('payload', {}).get('headers', [])
        emisor = 'desconocido'
        for h in headers:
            if h.get('name', '').lower() == 'from':
                emisor = h.get('value', '').split('<')[0].strip().replace(' ', '_').replace('"', '').replace("'", '')
                break
        # Obtener los adjuntos
        full_message = service.users().messages().get(userId=user_id, id=msg_id).execute()
        payload = full_message.get('payload', {})
        parts = payload.get('parts', [])
        # Crear carpeta si no existe
        if not os.path.exists(emisor):
            os.makedirs(emisor)
        # Primero, revisar si todos los adjuntos válidos ya existen
        todos_existen = True
        adjuntos_info = []
        for part in parts:
            filename = part.get('filename')
            mime_type = mimetypes.guess_type(filename)[0] if filename else None
            if filename and mime_type in tipos_validos:
                file_path = os.path.join(emisor, filename)
                adjuntos_info.append((file_path, part, filename))
                if not os.path.exists(file_path):
                    todos_existen = False
        if todos_existen or not adjuntos_info:
            # Todos los adjuntos válidos ya existen, omitir este correo
            return []
        # Si hay al menos uno nuevo, descargar solo los que no existen
        nuevos_descargados = []
        for file_path, part, filename in adjuntos_info:
            if os.path.exists(file_path):
                continue
            attachmentId = part.get('body', {}).get('attachmentId')
            if attachmentId:
                attachment = service.users().messages().attachments().get(
                    userId=user_id, messageId=msg_id, id=attachmentId
                ).execute()
                data = attachment.get('data')
                file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
                with open(file_path, 'wb') as f:
                    f.write(file_data)
                print(f"Descargado: {file_path}")
                nuevos_descargados.append((file_path, emisor, filename))
        return nuevos_descargados
    except HttpError as error:
        print(f"Ocurrió un error: {error}")
        return []

def procesar_archivo_con_vision(archivo):
    import pytesseract
    from PIL import Image
    # Ruta al ejecutable de tesseract (ajusta si es necesario)
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    ext = os.path.splitext(archivo)[1].lower()
    texto_extraido = ""
    if ext == '.pdf':
        import PyPDF2
        try:
            with open(archivo, 'rb') as f:
                reader = PyPDF2.PdfReader(f)
                for page in reader.pages:
                    texto_extraido += page.extract_text() or ''
        except Exception as e:
            texto_extraido = f"Error extrayendo texto del PDF: {e}"
        # Si no se extrajo nada, convertir la primera página a imagen y pasarla por Tesseract
        if not texto_extraido.strip():
            try:
                from pdf2image import convert_from_path
                import tempfile
                with tempfile.TemporaryDirectory() as path:
                    images = convert_from_path(archivo, output_folder=path, first_page=1, last_page=1)
                    for image in images:
                        temp_img = os.path.join(path, 'temp_page.png')
                        image.save(temp_img, 'PNG')
                        texto_extraido = pytesseract.image_to_string(Image.open(temp_img), lang='spa+eng')
                        break
            except Exception as e:
                texto_extraido = f"No se pudo extraer texto del PDF ni con OCR: {e}"
    elif ext in ['.jpg', '.jpeg', '.png']:
        try:
            texto_extraido = pytesseract.image_to_string(Image.open(archivo), lang='spa+eng')
        except Exception as e:
            texto_extraido = f"Error extrayendo texto de la imagen: {e}"
    else:
        texto_extraido = "Tipo de archivo no soportado."
    return texto_extraido

def main():
    service = get_service()
    if not service:
        return
    archivo_procesados = 'procesados.txt'
    print("Iniciando monitoreo de correos con adjuntos PDF/JPG/PNG. Presiona Ctrl+C para detener.")
    while True:
        if os.path.exists(archivo_procesados):
            with open(archivo_procesados, 'r', encoding='utf-8') as f:
                procesados = set(line.strip() for line in f if line.strip())
        else:
            procesados = set()
        nuevos_procesados = set()
        try:
            query = 'is:unread has:attachment (filename:pdf OR filename:jpg OR filename:jpeg OR filename:png)'
            results = service.users().messages().list(userId='me', q=query, maxResults=50).execute()
            messages = results.get('messages', [])
            archivos_nuevos_encontrados = False
            if not messages:
                print("No se encontraron correos con adjuntos válidos.")
            else:
                print(f"Se encontraron {len(messages)} correos con adjuntos válidos.")
                for message in messages:
                    msg_id = message['id']
                    archivos = descargar_adjuntos_validos(service, 'me', msg_id)
                    if not archivos:
                        continue  # No hubo adjuntos nuevos, no marcar como leído
                    procesado_este_correo = False
                    for archivo_info in archivos:
                        archivo, emisor, filename = archivo_info
                        if archivo in procesados or archivo in nuevos_procesados:
                            continue
                        archivos_nuevos_encontrados = True
                        print(f"Procesando {archivo}...")
                        texto = procesar_archivo_con_vision(archivo)
                        txt_filename = os.path.splitext(filename)[0] + '.txt'
                        txt_path = os.path.join(emisor, txt_filename)
                        with open(txt_path, 'w', encoding='utf-8') as f:
                            f.write(texto)
                        print(f"Texto extraído guardado en {txt_path}")
                        nuevos_procesados.add(archivo)
                        procesado_este_correo = True
                    # Marcar como leído solo si se descargó y procesó algún adjunto nuevo
                    if procesado_este_correo:
                        try:
                            service.users().messages().modify(userId='me', id=msg_id, body={'removeLabelIds': ['UNREAD']}).execute()
                        except Exception as e:
                            print(f"No se pudo marcar como leído el correo {msg_id}: {e}")
            if nuevos_procesados:
                with open(archivo_procesados, 'a', encoding='utf-8') as f:
                    for archivo in nuevos_procesados:
                        f.write(archivo + '\n')
            if not archivos_nuevos_encontrados:
                print("No hay archivos nuevos por procesar. Esperando nuevos correos...")
        except HttpError as error:
            print(f"Ocurrió un error: {error}")
        except Exception as e:
            print(f"Error inesperado: {e}")
        print("Esperando 10 segundos antes de volver a revisar...")
        time.sleep(10)

if __name__ == '__main__':
    main()