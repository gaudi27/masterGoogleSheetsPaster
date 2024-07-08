import os
import re
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError



#TODO - make sure the duplicate function is actually working

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]



def extract_sheet_id(url):
    """
    Extracts the Google Sheets ID from the provided URL.
    
    Args:
        url (str): The URL of the Google Sheet.
        
    Returns:
        str: The extracted Google Sheets ID.
    """
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    if match:
        return match.group(1)
    else:
        raise ValueError("Invalid Google Sheets URL")


def get_credentials():
    """Get user credentials for Google Sheets API."""
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return creds

def read_sheet(service, spreadsheet_id, range_name):
    """Read data from a Google Sheet."""
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    return result.get("values", [])

def write_to_sheet(service, spreadsheet_id, range_name, values):
    """Write data to a Google Sheet."""
    body = {
        "values": values  #possible place to specify emails and names
    }
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption="RAW",
        body=body
    ).execute()

def merge_data(existing_data, new_data):
    """Merge new data into existing data without duplicates."""
    merged_data = existing_data[:]
    existing_set = set(tuple(row) for row in existing_data)
    
    for row in new_data:
        if tuple(row) not in existing_set:
            merged_data.append(row)
            existing_set.add(tuple(row))
    
    return merged_data

def paster(MASTER_SHEET_URL, SHEET_URL):
    creds = get_credentials()
    service = build("sheets", "v4", credentials=creds)
    
    #convert URL to ID
    SHEET_IDS = [extract_sheet_id(url) for url in SHEET_URL]
    MASTER_SHEET_ID = extract_sheet_id(MASTER_SHEET_URL)

    # Read existing data from the master sheet
    existing_data = read_sheet(service, MASTER_SHEET_ID, "A1:Z1000")
    if existing_data is None:
        existing_data = []
    
    # Consolidate data from multiple sheets
    all_data = existing_data[:]
    for sheet_id in SHEET_IDS:
        new_data = read_sheet(service, sheet_id, "A1:Z1000")  # Fetch data from a larger range
        if new_data:
            all_data = merge_data(all_data, new_data)
    
    # Write consolidated data to the master sheet
    write_to_sheet(service, MASTER_SHEET_ID, "Sheet1!A1", all_data)
    print("Data has been successfully consolidated into the master sheet.")

