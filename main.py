import os
import base64
from pathlib import Path

from dotenv import load_dotenv
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

load_dotenv()

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

BASE_DIR = Path(__file__).resolve().parent
CREDENTIALS_FILE = BASE_DIR / "credentials.json"
TOKEN_FILE = BASE_DIR / "token.json"


def authenticate_gmail():
    creds = None

    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(
            str(TOKEN_FILE),
            SCOPES,
        )

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not CREDENTIALS_FILE.exists():
                raise FileNotFoundError(
                    f"No existe {CREDENTIALS_FILE}"
                )

            flow = InstalledAppFlow.from_client_secrets_file(
                str(CREDENTIALS_FILE),
                SCOPES,
            )

            creds = flow.run_local_server(
                port=0,
                access_type="offline",
                prompt="consent",
            )

        TOKEN_FILE.write_text(creds.to_json())

    return build(
        "gmail",
        "v1",
        credentials=creds,
    )


def decode_body(data: str) -> str:
    decoded = base64.urlsafe_b64decode(
        data + "=" * (-len(data) % 4)
    )
    return decoded.decode("utf-8", errors="replace")


def extract_text(payload):
    body = ""

    if payload.get("body", {}).get("data"):
        body = decode_body(payload["body"]["data"])

    for part in payload.get("parts", []):
        mime_type = part.get("mimeType", "")

        if mime_type == "text/plain":
            data = part.get("body", {}).get("data")
            if data:
                return decode_body(data)

        nested = extract_text(part)
        if nested:
            body = nested

    return body


def get_header(headers, name):
    name = name.lower()

    for header in headers:
        if header.get("name", "").lower() == name:
            return header.get("value", "")

    return ""


def find_job_emails(service):
    query = (
        'newer_than:1d '
        '(from:linkedin.com OR '
        'subject:(application OR solicitud OR postulación OR postulaste))'
    )

    response = (
        service.users()
        .messages()
        .list(
            userId="me",
            q=query,
            maxResults=50,
        )
        .execute()
    )

    messages = response.get("messages", [])

    results = []

    for item in messages:
        message = (
            service.users()
            .messages()
            .get(
                userId="me",
                id=item["id"],
                format="full",
            )
            .execute()
        )

        payload = message.get("payload", {})
        headers = payload.get("headers", [])

        subject = get_header(headers, "Subject")
        sender = get_header(headers, "From")
        date = get_header(headers, "Date")
        body = extract_text(payload)

        results.append(
            {
                "id": message["id"],
                "thread_id": message.get("threadId"),
                "subject": subject,
                "from": sender,
                "date": date,
                "body": body,
            }
        )

    return results


def main():
    print("\n🔐 Autenticando Gmail...")

    service = authenticate_gmail()

    print("✅ Gmail conectado.")

    print("\n📬 Buscando correos de postulaciones...\n")

    emails = find_job_emails(service)

    print(f"Encontrados: {len(emails)}\n")

    for email in emails:
        print("=" * 80)
        print(f"ID:      {email['id']}")
        print(f"FROM:    {email['from']}")
        print(f"SUBJECT: {email['subject']}")
        print(f"DATE:    {email['date']}")
        print("-" * 80)
        print(email["body"][:3000])
        print("=" * 80)


if __name__ == "__main__":
    main()