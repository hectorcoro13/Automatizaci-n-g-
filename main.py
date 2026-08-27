import base64
import re
from datetime import datetime, timedelta, date
from email.utils import parsedate_to_datetime
from pathlib import Path
from typing import Optional
from zoneinfo import ZoneInfo

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# ============================================================
# CONFIGURACIÓN
# ============================================================

SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]

BASE_DIR = Path(__file__).resolve().parent
CREDENTIALS_FILE = BASE_DIR / "credentials.json"
TOKEN_FILE = BASE_DIR / "token.json"

TIMEZONE = "America/Bogota"

# ID tomado directamente de tu URL:
# https://docs.google.com/spreadsheets/d/13r6.../edit
SPREADSHEET_ID = "13r6Nt0JVQ9lCLUE_HPHJGl50LlAfobyLsPeGYoJho5M"


# ============================================================
# GOOGLE AUTH
# ============================================================

def authenticate_google():
    creds = None

    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(
            str(TOKEN_FILE),
            SCOPES,
        )

    # El token anterior fue creado solo con gmail.readonly.
    # Al agregar spreadsheets, necesitamos reautorizar.
    if not creds or not creds.valid or not creds.has_scopes(SCOPES):
        if creds and creds.expired and creds.refresh_token:
            try:
                creds = Credentials(
                    token=None,
                    refresh_token=creds.refresh_token,
                    token_uri=creds.token_uri,
                    client_id=creds.client_id,
                    client_secret=creds.client_secret,
                    scopes=SCOPES,
                )
                creds.refresh(Request())
            except Exception:
                creds = None

        if not creds or not creds.valid or not creds.has_scopes(SCOPES):
            if not CREDENTIALS_FILE.exists():
                raise FileNotFoundError(
                    f"No existe {CREDENTIALS_FILE}"
                )

            # Borra token viejo manualmente si esta etapa falla.
            flow = InstalledAppFlow.from_client_secrets_file(
                str(CREDENTIALS_FILE),
                SCOPES,
            )
            creds = flow.run_local_server(
                port=0,
                access_type="offline",
                prompt="consent",
            )

        TOKEN_FILE.write_text(creds.to_json(), encoding="utf-8")

    gmail = build("gmail", "v1", credentials=creds)
    sheets = build("sheets", "v4", credentials=creds)

    return gmail, sheets


# ============================================================
# SHEETS - LECTURA DE HISTORIAL
# ============================================================

def get_sheet_metadata(sheets):
    response = (
        sheets.spreadsheets()
        .get(
            spreadsheetId=SPREADSHEET_ID,
            includeGridData=False,
        )
        .execute()
    )

    title = response.get("properties", {}).get("title", "")
    sheets_info = response.get("sheets", [])

    print(f"\n📊 Spreadsheet: {title}")
    print("📄 Hojas:")

    for sheet in sheets_info:
        props = sheet.get("properties", {})
        print(
            f"   - {props.get('title')} "
            f"(gid={props.get('sheetId')})"
        )

    return sheets_info


def read_first_sheet(sheets):
    metadata = get_sheet_metadata(sheets)

    if not metadata:
        raise RuntimeError("El spreadsheet no tiene hojas.")

    first_title = metadata[0]["properties"]["title"]

    result = (
        sheets.spreadsheets()
        .values()
        .get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{first_title}'!A:Z",
        )
        .execute()
    )

    values = result.get("values", [])

    print(f"\n🧾 Hoja usada: {first_title}")
    print(f"Filas encontradas: {len(values)}")

    if not values:
        print("La hoja está vacía.")
        return first_title, []

    headers = values[0]
    print("\nENCABEZADOS:")
    for i, header in enumerate(headers, start=1):
        print(f"  {i}. {header}")

    print("\nÚLTIMAS 10 FILAS:")
    for row in values[-10:]:
        print("  ", row)

    return first_title, values


# ============================================================
# GMAIL
# ============================================================

def decode_body(data: str) -> str:
    decoded = base64.urlsafe_b64decode(
        data + "=" * (-len(data) % 4)
    )
    return decoded.decode("utf-8", errors="replace")


def html_to_text(html: str) -> str:
    text = re.sub(
        r"(?is)<(script|style).*?>.*?</\1>",
        " ",
        html,
    )
    text = re.sub(r"(?i)<br\s*/?>", "\n", text)
    text = re.sub(r"(?i)</p\s*>", "\n", text)
    text = re.sub(r"(?s)<[^>]+>", " ", text)
    text = re.sub(r"&nbsp;", " ", text)
    text = re.sub(r"&amp;", "&", text)
    text = re.sub(r"&lt;", "<", text)
    text = re.sub(r"&gt;", ">", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def extract_text(payload: dict) -> str:
    plain_parts = []
    html_parts = []

    def walk(part):
        mime = (part.get("mimeType") or "").lower()
        data = (part.get("body") or {}).get("data")

        if data:
            decoded = decode_body(data)
            if mime == "text/plain":
                plain_parts.append(decoded)
            elif mime == "text/html":
                html_parts.append(decoded)

        for child in part.get("parts", []) or []:
            walk(child)

    walk(payload)

    if plain_parts:
        return "\n".join(plain_parts).strip()

    if html_parts:
        return html_to_text("\n".join(html_parts))

    return ""


def get_header(headers, name: str) -> str:
    wanted = name.lower()

    for header in headers:
        if header.get("name", "").lower() == wanted:
            return header.get("value", "")

    return ""


def colombia_date_from_message(message: dict) -> Optional[date]:
    try:
        ms = int(message.get("internalDate", "0"))
        if ms:
            dt = datetime.fromtimestamp(
                ms / 1000,
                tz=ZoneInfo(TIMEZONE),
            )
            return dt.date()
    except (ValueError, TypeError, OSError):
        pass

    try:
        raw = get_header(
            message.get("payload", {}).get("headers", []),
            "Date",
        )
        return parsedate_to_datetime(raw).astimezone(
            ZoneInfo(TIMEZONE)
        ).date()
    except Exception:
        return None


# ============================================================
# FILTRO DE CONFIRMACIÓN
# ============================================================

POSITIVE_PATTERNS = [
    r"\bse ha enviado tu solicitud\b",
    r"\bse envió tu solicitud\b",
    r"\btu solicitud fue enviada\b",
    r"\bhemos recibido tu solicitud\b",
    r"\bhemos recibido su solicitud\b",
    r"\bgracias por tu solicitud\b",
    r"\bgracias por postularte\b",
    r"\bgracias por aplicar\b",
    r"\bapplication received\b",
    r"\bapplication submitted\b",
    r"\bthank you for applying\b",
    r"\bwe (have )?received your application\b",
    r"\byour application (has been|was) submitted\b",
]

NEGATIVE_SENDERS = [
    "newsletters-noreply@linkedin.com",
    "updates-noreply@linkedin.com",
    "jobalerts-noreply@linkedin.com",
]

NEGATIVE_PATTERNS = [
    r"\bjob alert\b",
    r"\balerta de empleo\b",
    r"\bnuevos empleos\b",
    r"\bnew jobs\b",
    r"\bempleos similares\b",
    r"\bjobs similar\b",
    r"\bnewsletter\b",
    r"\bbolet[ií]n\b",
    r"\bver anuncio de empleo\b",
    r"\bsolicitar con perfil y cv\b",
]


def matches_any(patterns, text: str) -> bool:
    return any(
        re.search(pattern, text, flags=re.IGNORECASE)
        for pattern in patterns
    )


def classify_email(message: dict):
    subject = message["subject"].strip()
    sender = message["from"].lower()
    body = message["body"]

    sample = f"{subject}\n{body}"

    if any(addr in sender for addr in NEGATIVE_SENDERS):
        return False, "newsletter/alerta de LinkedIn"

    if not matches_any(POSITIVE_PATTERNS, sample):
        return False, "no confirma una postulación enviada"

    if matches_any(NEGATIVE_PATTERNS, sample):
        # Una confirmación explícita tiene prioridad.
        if not matches_any(POSITIVE_PATTERNS, sample):
            return False, "oferta/alerta/newsletter"

    return True, "confirmación de postulación"


def find_today_candidates(gmail, max_results=100):
    today = datetime.now(
        ZoneInfo(TIMEZONE)
    ).date()
    tomorrow = today + timedelta(days=1)

    query = (
        f"after:{today.strftime('%Y/%m/%d')} "
        f"before:{tomorrow.strftime('%Y/%m/%d')}"
    )

    response = (
        gmail.users()
        .messages()
        .list(
            userId="me",
            q=query,
            maxResults=min(max_results, 100),
        )
        .execute()
    )

    candidates = []

    for item in response.get("messages", []):
        message = (
            gmail.users()
            .messages()
            .get(
                userId="me",
                id=item["id"],
                format="full",
            )
            .execute()
        )

        if colombia_date_from_message(message) != today:
            continue

        payload = message.get("payload", {})
        headers = payload.get("headers", [])

        email = {
            "id": message["id"],
            "thread_id": message.get("threadId"),
            "subject": get_header(headers, "Subject"),
            "from": get_header(headers, "From"),
            "date": get_header(headers, "Date"),
            "body": extract_text(payload),
        }

        is_candidate, reason = classify_email(email)

        print("\n" + "=" * 90)
        print(f"SUBJECT: {email['subject']}")
        print(f"FROM:    {email['from']}")
        print(
            f"RESULT:  "
            f"{'✅ CANDIDATO' if is_candidate else '❌ IGNORADO'}"
        )
        print(f"MOTIVO:  {reason}")

        if is_candidate:
            candidates.append(email)

    return candidates


# ============================================================
# MAIN - SOLO PRUEBA, NO ESCRIBE EN SHEETS
# ============================================================

def main():
    print("\n🔐 Autenticando Gmail + Google Sheets...")
    gmail, sheets = authenticate_google()
    print("✅ Google autenticado.")

    # 1. Confirmar acceso a tu Sheet.
    _, history = read_first_sheet(sheets)

    # 2. Revisar correos de hoy.
    candidates = find_today_candidates(gmail)

    print("\n" + "=" * 90)
    print("RESUMEN")
    print(f"📊 Filas históricas: {max(len(history) - 1, 0)}")
    print(f"📬 Candidatos de hoy: {len(candidates)}")
    print("=" * 90)

    print(
        "\n⚠️ ESTA VERSIÓN NO ESCRIBE NADA EN LA HOJA."
        "\nPrimero verificamos que lea correctamente tu historial."
    )


if __name__ == "__main__":
    main()
