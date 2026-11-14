import os
from pathlib import Path
from dotenv import load_dotenv
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

load_dotenv()

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
BASE_DIR = Path(__file__).resolve().parent
CREDENTIALS_FILE = BASE_DIR / "credentials.json"
TOKEN_FILE = BASE_DIR / "token.json"

SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "").strip()

def get_service():
    """Authenticate a local desktop user with read-only Sheets access."""
    if not SPREADSHEET_ID:
        raise ValueError("Set SPREADSHEET_ID to use SOURCE_TYPE=google_sheets.")
    if not CREDENTIALS_FILE.exists():
        raise FileNotFoundError(
            "credentials.json is missing. Create a Desktop OAuth client in "
            "Google Cloud Console and place the downloaded file beside app.py."
        )

    creds = None
    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                str(CREDENTIALS_FILE), SCOPES
            )
            creds = flow.run_local_server(port=0)
        TOKEN_FILE.write_text(creds.to_json(), encoding="utf-8")

    return build("sheets", "v4", credentials=creds, cache_discovery=False)

def read_spreadsheet():
    """Return spreadsheet title and a dict of worksheet DataFrames."""
    service = get_service()
    metadata = service.spreadsheets().get(
        spreadsheetId=SPREADSHEET_ID,
        fields="properties(title),sheets(properties(sheetId,title,gridProperties))",
    ).execute()

    title = metadata.get("properties", {}).get("title", "Google Spreadsheet")
    tabs = metadata.get("sheets", [])
    results = {}

    for tab in tabs:
        tab_title = tab["properties"]["title"]
        # Quote sheet names safely, including names containing apostrophes.
        quoted = "'" + tab_title.replace("'", "''") + "'"
        response = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{quoted}!A:Z",
            valueRenderOption="UNFORMATTED_VALUE",
            dateTimeRenderOption="FORMATTED_STRING",
        ).execute()
        results[tab_title] = response.get("values", [])

    return title, results
