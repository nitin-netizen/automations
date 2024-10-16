import pandas as pd
import yagmail
import requests
from requests_oauthlib import OAuth2Session
from oauthlib.oauth2 import BackendApplicationClient

# Replace with your Azure AD credentials
client_id = 'YOUR_CLIENT_ID'
client_secret = 'YOUR_CLIENT_SECRET'
tenant_id = 'YOUR_TENANT_ID'

# Outlook API Base URL
outlook_api_url = 'https://graph.microsoft.com/v1.0/'

# Function to get access token from Microsoft Graph API
def get_access_token(client_id, client_secret, tenant_id):
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    client = BackendApplicationClient(client_id=client_id)
    oauth = OAuth2Session(client=client)

    token_data = {
        'client_id': client_id,
        'scope': 'https://graph.microsoft.com/.default',
        'client_secret': client_secret,
        'grant_type': 'client_credentials'
    }

    response = oauth.fetch_token(token_url=token_url, client_id=client_id, client_secret=client_secret, include_client_id=True)
    return response['access_token']

# Function to get calendar events from Outlook
def get_outlook_calendar_events(access_token):
    headers = {
        'Authorization': f'Bearer {access_token}'
    }
    events_url = outlook_api_url + 'me/events'
    response = requests.get(events_url, headers=headers)
    
    if response.status_code == 200:
        return response.json().get('value', [])
    else:
        print(f"Error fetching events: {response.status_code}")
        return []

# Function to send an email with extracted data
def send_email(data_from_excel, calendar_events):
    # Initialize the yagmail client with your email and app-specific password
    yag = yagmail.SMTP('your_email@example.com', 'your_app_password')

    # Format calendar events into a readable string
    calendar_event_details = "\n".join([f"Event: {event['subject']} on {event['start']['dateTime']}" for event in calendar_events])

    # Create the email body
    body = f"""
    Here is the extracted data from the Excel file:
    {data_from_excel}

    Calendar Events:
    {calendar_event_details}
    """

    # Send the email
    yag.send(
        to='recipient@example.com',
        subject='Automated Report: Excel Data and Calendar Events',
        contents=body
    )
    print('Email sent successfully!')

# Main automation process
def main():
    # Step 1: Load data from Excel file
    excel_file_path = 'your_excel_file.xlsx'
    excel_data = pd.read_excel(excel_file_path)
    
    # Step 2: Get Outlook calendar events
    access_token = get_access_token(client_id, client_secret, tenant_id)
    calendar_events = get_outlook_calendar_events(access_token)

    # Step 3: Send the data via email
    send_email(excel_data.head().to_string(), calendar_events)

if __name__ == "__main__":
    main()
