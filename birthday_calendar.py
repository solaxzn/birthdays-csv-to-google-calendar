from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pandas as pd
import os.path
import csv
import pickle
from datetime import datetime, timedelta

SCOPES = ['https://www.googleapis.com/auth/calendar']

def get_google_calendar_service():
    # Sets up and returns the Google Calendar service.
    creds = None
    
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=8081)
        
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return build('calendar', 'v3', credentials=creds)

def create_birthday_event(service, name, birthday_date):
    #Creates a recurring yearly birthday event
    date_obj = datetime.strptime(birthday_date, "%b %d")
    
    current_year = datetime.now().year
    event_date = date_obj.replace(year=current_year)
    
    event = {
        'summary': f"{name}'s Birthday 🎂",
        'description': f"Happy Birthday {name}! 🎉",
        'start': {
            'date': event_date.strftime('%Y-%m-%d'),
            'timeZone': 'UTC',
        },
        'end': {
            'date': (event_date + timedelta(days=1)).strftime('%Y-%m-%d'),
            'timeZone': 'UTC',
        },
        'recurrence': [
            'RRULE:FREQ=YEARLY'
        ],
        'reminders': {
            'useDefault': False,
            'overrides': [
                {'method': 'popup', 'minutes': 24 * 60},  # 1 day before
                {'method': 'popup', 'minutes': 60}        # 1 hour before
            ],
        },
    }
    
    return service.events().insert(calendarId='primary', body=event).execute()



def main():

    service = get_google_calendar_service()

    file_path = "birthdays.csv"
    df = pd.read_csv(file_path)
    # print(df) # Debug dataframe

    for index, row in df.iterrows():
        if pd.isna(row.get("Added_to_calendar")):
            try:
                event = create_birthday_event(service, row['Name'], row['Birthday'])
                print(f"{row['Name']}\'s birthday added to your calendar!")
                df.at[index, "Added_to_calendar"] = "Yes"
            except Exception as e:
                print(f"Error creating event for {row['Name']}: {str(e)}")

    df.to_csv(file_path, index=False)


if __name__ == '__main__':
    main()