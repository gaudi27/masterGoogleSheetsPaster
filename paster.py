import os
import sys
import json
import re
import mimetypes
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build, MediaFileUpload
from googleapiclient.errors import HttpError

#TODO - be able to convert xlsx files to google sheets files
#TODO - make sure the duplicate function is working
#TODO - make into a onefile

def get_credentials_path():
    if getattr(sys, 'frozen', False):  # Check if running as compiled executable
        # Running in PyInstaller bundle
        base_path = sys._MEIPASS
    else:
        # Running as a script
        base_path = os.path.abspath(os.path.dirname(__file__))
    
    credentials_path = os.path.join(base_path, 'credentials.json')
    return credentials_path


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


#TODO - finish this function to convert xlsx files to google sheets
def upload_and_convert_to_sheets(service_drive, file_path):
    """Uploads an .xlsx file to Google Drive and converts it to Google Sheets."""
    file_metadata = {
        'name': os.path.basename(file_path),
        'mimeType': 'application/vnd.google-apps.spreadsheet'
    }
    media = MediaFileUpload(file_path, mimetype=mimetypes.guess_type(file_path)[0])
    
    file = service_drive.files().create(
        body=file_metadata,
        media_body=media,
        fields='id'
    ).execute()
    
    return file.get('id')


def read_sheet(service, spreadsheet_id, range_name):
    """Read data from a Google Sheet."""
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    return result.get("values", [])

def write_to_sheet(service, spreadsheet_id, range_name, values):
    """Write data to a Google Sheet."""
    body = {
        "values": values
    }
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption="RAW",
        body=body
    ).execute()

def extract_relevant_data(values):
    """Extract first name, last name, and email from the data based on headers."""
    if not values:
        return []
    
    headers = values[0]
    #names
    first_name_index = headers.index("First Name") if "First Name" in headers else None
    last_name_index = headers.index("Last Name") if "Last Name" in headers else None
    name_index = headers.index("name") if "name" in headers else None
    #emails
    email_index = headers.index("Email") if "Email" in headers else None
    email_index2 = headers.index("email") if "email" in headers else None
    email_index3 = headers.index("Email Address") if "Email Address" in headers else None

    extracted_values = []
    for row in values[1:]:  # Skip the header row
        extracted_row = []

        #cases for names
        if first_name_index is not None:
            extracted_row.append(row[first_name_index] if len(row) > first_name_index else '')
        if last_name_index is not None:
            extracted_row.append(row[last_name_index] if len(row) > last_name_index else '')
        if name_index is not None:
            extracted_row.append(row[name_index] if len(row) > name_index else '')

        #cases for emails
        if email_index is not None:
            extracted_row.append(row[email_index] if len(row) > email_index else '')
        if email_index2 is not None:
            extracted_row.append(row[email_index2] if len(row) > email_index2 else '') 
        if email_index3 is not None:
            extracted_row.append(row[email_index3] if len(row) > email_index3 else '')
        
        if extracted_row:  # Only add rows that have at least one value
            extracted_values.append(extracted_row)
    return extracted_values


'''def merge_data(existing_data, new_data):
    """Merge new data into existing data without duplicates."""
    merged_data = existing_data[:]
    existing_set = set((row[0], row[2]) for row in existing_data)
    
    for row in new_data:
        if len(row) >= 3:  # Ensure row has at least three elements (first name, last name, email)
            key = (row[0], row[2])  # Assuming first name and email are in the first and third columns
            if key not in existing_set:
                merged_data.append(row)
                existing_set.add(key)
        elif len(row) >= 2:  # Handle case where last name is missing
            key = (row[0], row[1])  # Use only first name and email
            if key not in existing_set:
                merged_data.append(row)
                existing_set.add(key)
        else:
            # Handle case where row doesn't have enough elements to extract required fields
            print(f"Ignoring invalid row: {row}")
    
    return merged_data
'''




def merge_data(existing_data, new_data):
    """Merge new data into existing data without duplicates."""
    merged_data = existing_data[:]
    existing_set = set(tuple(row) for row in existing_data)
    
    for row in new_data:
        if tuple(row) not in existing_set:
            merged_data.append(row)
            existing_set.add(tuple(row))
    
    return merged_data


def remove_duplicates(all_data):
    seen_rows = set()
    unique_data = []
    
    for row in all_data:
        # Convert row to tuple to make it hashable and add to set
        row_tuple = tuple(row)
        if row_tuple not in seen_rows:
            unique_data.append(row)
            seen_rows.add(row_tuple)
    
    return unique_data
    

def paster(MASTER_SHEET_URL, SHEET_URL):
    creds = get_credentials()
    service = build("sheets", "v4", credentials=creds)
    service_drive = build("drive", "v3", credentials=creds)

    #convert URL to ID
    SHEET_IDS = [extract_sheet_id(url) for url in SHEET_URL]

    #TODO - convert xlsx files to google sheets
    '''sheet_ids = []
    for url in SHEET_URL:
        try:
            # Attempt to extract Google Sheets ID
            sheet_id = extract_sheet_id(url)
        except ValueError:
            # If it's not a valid Google Sheets URL, assume it's a file path
            if os.path.isfile(url):
                sheet_id = upload_and_convert_to_sheets(service_drive, url)
                #            else:
                #raise ValueError(f"Invalid URL or file path: {url}")
        sheet_ids.append(sheet_id)'''

    
    MASTER_SHEET_ID = extract_sheet_id(MASTER_SHEET_URL)

    # Read existing data from the master sheet
    existing_data = read_sheet(service, MASTER_SHEET_ID, "A1:Z1000")
    if existing_data is None:
        existing_data = []
    
    # Consolidate data from multiple sheets
    all_data = existing_data[:]
    for sheet_id in SHEET_IDS:
        range_name = "A1:Z1000"  # Specify the range to cover a large area
        sheet_data = read_sheet(service, sheet_id, range_name)
        if sheet_data:
            new_data = extract_relevant_data(sheet_data)
            all_data = merge_data(all_data, new_data)
    all_data = remove_duplicates(all_data)
    # Write consolidated data to the master sheet
    write_to_sheet(service, MASTER_SHEET_ID, "Sheet1!A1", all_data)
    print("Data has been successfully consolidated into the master sheet.")








