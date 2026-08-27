import os
import base64
import json
from datetime import datetime
from dotenv import load_dotenv
import gspread
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import openai

# Cargar variables de entorno (API Keys de manera segura)
load_dotenv()
openai.api_key = os.getenv("OPENAI_API_KEY")

# Scopes requeridos por Google
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
SPREADSHEET_URL = os.getenv("SPREADSHEET_URL") # El link de tu Google Sheet

def authenticate_gmail():
    """Autentica y retorna el servicio de Gmail."""
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
    return build('gmail', 'v1', credentials=creds)

def fetch_daily_job_emails(service):
    """Busca correos de LinkedIn o confirmaciones recibidas en el último día."""
    query = "subject:solicitud OR subject:postulación newer_than:1d"
    try:
        results = service.users().messages().list(userId='me', q=query).execute()
        messages = results.get('messages', [])
        email_contents = []

        for msg in messages:
            msg_data = service.users().messages().get(userId='me', id=msg['id'], format='full').execute()
            payload = msg_data.get('payload', {})
            
            # Extraer cuerpo del correo (simplificado para texto plano)
            body = ""
            if 'parts' in payload:
                for part in payload['parts']:
                    if part['mimeType'] == 'text/plain':
                        data = part['body'].get('data')
                        if data:
                            body = base64.urlsafe_b64decode(data).decode('utf-8')
                            break
            
            if body:
                email_contents.append(body)
                
        return email_contents
    except Exception as e:
        print(f"Error al obtener correos: {e}")
        return []

def extract_data_with_ai(email_text):
    """Usa OpenAI para extraer la información exacta en formato JSON."""
    prompt = f"""
    Eres un asistente experto en extracción de datos. Analiza el siguiente correo de confirmación de postulación laboral y extrae los datos en formato JSON con estas claves exactas:
    - "Fecha" (en formato DD/MM/YYYY, usa la fecha de hoy si no se especifica)
    - "Nombre_empresa"
    - "Nombre_rol"
    - "Link_vacante" (si no está el link explícito, pon "No disponible")
    - "Medio_postulacion" (por ejemplo, LinkedIn, Computrabajo, Directo)

    Correo:
    {email_text}
    """
    try:
        response = openai.ChatCompletion.create(
            model="gpt-4o-mini", # Rápido, económico y altamente eficiente
            messages=[{"role": "user", "content": prompt}],
            temperature=0.0
        )
        # Limpiar y parsear la respuesta
        raw_json = response.choices[0].message.content.replace("```json", "").replace("```", "").strip()
        return json.loads(raw_json)
    except Exception as e:
        print(f"Error en extracción de IA: {e}")
        return None

def update_google_sheet(data):
    """Añade una nueva fila al Google Sheet usando la cuenta de servicio."""
    try:
        gc = gspread.service_account(filename='service_account.json')
        sh = gc.open_by_url(SPREADSHEET_URL)
        worksheet = sh.sheet1 # Selecciona la primera hoja
        
        # Preparar la fila basándonos en tu 'Captura de pantalla 2026-08-21 a la(s) 12.55.14 p.m..png'
        row = [
            data.get("Fecha", datetime.now().strftime("%d/%m/%Y")),
            data.get("Nombre_empresa", ""),
            data.get("Nombre_rol", ""),
            data.get("Link_vacante", ""),
            data.get("Medio_postulacion", "")
        ]
        
        worksheet.append_row(row)
        print(f"Éxito: Se agregó {data.get('Nombre_empresa')} al Excel.")
    except Exception as e:
        print(f"Error al actualizar Sheets: {e}")

def main():
    print("Iniciando revisión diaria de correos...")
    gmail_service = authenticate_gmail()
    emails = fetch_daily_job_emails(gmail_service)
    
    if not emails:
        print("No se encontraron nuevas postulaciones hoy.")
        return

    for email in emails:
        extracted_data = extract_data_with_ai(email)
        if extracted_data:
            update_google_sheet(extracted_data)

if __name__ == '__main__':
    main()