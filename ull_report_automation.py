from __future__ import print_function
import os
import logging
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import pandas as pd


SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']

ULL_INFO_WEBSITE_PASS = os.getenv('ULL_INFO_WEBSITE_PASS')
ULL_INFO_GOOGLE_PASS = os.getenv('ULL_INFO_GOOGLE_PASS')

REPORT_URL = "https://reporting.bluesombrero.com/25410/admin/saved/160583"
SPORTS_CONNECT_URL = "https://www.unionlittleleaguebaseball.com/Default.aspx?tabid=1888373&isLogin=True"

GMAIL_USERNAME = "info@unionlittleleaguebaseball.com"

REGISTRATION_CHART_ID = "1vvHHbIma2pRyycOfxYIOsZqvsCM1c_dg9Y1TUp1lnK4"

CSV_PATH = 'Enrollment_Details.csv'

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')

def setup_driver():
    """Sets up the Selenium WebDriver with headless Chrome."""
    chrome_options = Options()
    #chrome_options.add_argument("--headless=new")
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    #Add user-agent to fix 500 error from Google when using headless mode
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36")
    prefs = {
        "download.default_directory": os.path.dirname(os.path.realpath(__file__)), # Set your desired download folder
        "download.prompt_for_download": False,  # Disable the download prompt
        "profile.default_content_settings.popups": 0,  # Disable popups
    }
    chrome_options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    return driver

def wait_for_redirects_to_complete(driver, timeout=10):
    current_url = driver.current_url
    WebDriverWait(driver, timeout).until(lambda d: d.current_url != current_url)

def do_sportsconnect_login(driver):
    logging.info("Logging in to SportsConnect...")
    driver.get(SPORTS_CONNECT_URL)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "email"))).send_keys(GMAIL_USERNAME)    
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, "continue"))).click()
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "password"))).send_keys(ULL_INFO_WEBSITE_PASS)    
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, "continue"))).click()
    # Wait for redirects to complete
    wait_for_redirects_to_complete(driver)

def do_csv_download(driver):
    logging.info("Starting data export...")
    driver.get(REPORT_URL)
    sports_connect_window = driver.current_window_handle
    driver.set_window_size(1920, 1080)
    time.sleep(10)

    export_span = driver.find_element(By.XPATH, "//*[text()='Export']")
    export_span.click()
    time.sleep(1)

    export_button = driver.find_element(By.XPATH, "//*[text()=' CSV ']")
    export_button.click()

    time.sleep(2)

def get_credentials():
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
    return creds

def upload_csv(drive_service, sheets_service, today):
    new_spreadsheet = {
        'name': 'Temporary Spreadsheet',
        'mimeType': 'application/vnd.google-apps.spreadsheet'
    }

    file = drive_service.files().create(body=new_spreadsheet, fields='id').execute()
    spreadsheet_id = file.get('id')

    sheet_name = 'Sheet1'  # Replace with the name of your sheet (default is "Sheet1")

    # Read the CSV file
    csv_file = CSV_PATH  # Replace with the path to your CSV file
    data = pd.read_csv(csv_file)
    data_list = [data.columns.tolist()] + data.values.tolist()  # Convert DataFrame to list format

    # Prepare the data for the Sheets API
    body = {
        'values': data_list
    }

    # Write data to the spreadsheet
    result = sheets_service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=sheet_name,
        valueInputOption='RAW',
        body=body
    ).execute()

    return spreadsheet_id

def convert_file(drive_service, excel_sheet_id):
    newfile = {'name': 'Thing', 'mimeType': 'application/vnd.google-apps.spreadsheet'}
    result = drive_service.files().copy(fileId=excel_sheet_id, body=newfile).execute()
    logging.info("Converted to Google Sheets")
    return result['id']

# Update the chart sheet with the data
def update_chart_sheet(sheets_service, sheets_sheet_id, current_signups):
    body = {'values': current_signups.get("values")}
    sheets_service.spreadsheets().values().update(
        spreadsheetId=REGISTRATION_CHART_ID, range='2025 Data!A2:H500',
        valueInputOption='USER_ENTERED', body=body).execute()
    today = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    # Update cell B15 with today's date and time
    body = {
        'values': [[today]]  # Wrap the date and time in a list of lists
    }
    sheets_service.spreadsheets().values().update(
        spreadsheetId=REGISTRATION_CHART_ID, range='Report!B15',
        valueInputOption='USER_ENTERED', body=body).execute()
    logging.info("Chart updated")

# Clean up files
def clean_up_files(drive_service, sheets_sheet_id):
    os.remove(CSV_PATH)
    drive_service.files().delete(fileId=sheets_sheet_id).execute()
    logging.info("Cleanup completed!")


def do_google_sheets_auto():
    logging.info("Starting Google Sheets automation...")
    
    # Step 1: Get credentials
    creds = get_credentials()

    try:
        # Step 2: Build the Drive and Sheets services
        drive_service = build('drive', 'v3', credentials=creds)
        sheets_service = build('sheets', 'v4', credentials=creds)
        
        # Step 3: Upload the CSV as a google sheet
        today = datetime.today() - timedelta(hours=5, minutes=5)
        sheets_sheet_id = upload_csv(drive_service, sheets_service, today)

        # Step 4: Retrieve and update data
        current_signups = sheets_service.spreadsheets().values().get(spreadsheetId=sheets_sheet_id, range='A2:H500').execute()
        logging.info("Retrieved current signup data...")
        update_chart_sheet(sheets_service, sheets_sheet_id, current_signups)

        #Step 6: Clean up resources
        clean_up_files(drive_service, sheets_sheet_id)

    except HttpError as err:
        logging.error(f"An error occurred: {err}")

def main():
    driver = setup_driver()
    do_sportsconnect_login(driver)
    do_csv_download(driver)
    driver.quit()
    do_google_sheets_auto()

if __name__ == '__main__':
    main()
    