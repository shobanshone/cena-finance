from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os.path
import pickle

credentials=None

#Loading the credentials from pickle file
if os.path.exists(r".\\token.pickle"):
                print("Accessing credentials...")
                with open("token.pickle",'rb') as file:
                    credentials= pickle.load(file)

#if credentials are either expired or not exist , refresh or create credentials
if not credentials or not credentials.valid:
    if credentials and credentials.expired and credentials.refresh_token:
        print("Refreshing access token....")
        credentials.refresh(Request())
    else:
        print('not able to refresh token')
        print("Fetching new access token....")
        flow = InstalledAppFlow.from_client_secrets_file(
            "client_secret.json", scopes=['https://www.googleapis.com/auth/calendar'])
        flow.run_local_server(
                    port=8080, authorization_prompt_message=""
                    )
        
        credentials = flow.credentials
        with open("token.pickle",'wb') as file:
            print("Saving credentials for future use...")
            pickle.dump(credentials,file)

#variable for storing calendar from google calendar api
remainder=build('calendar', 'v3', credentials=credentials)

#getting the calendar list and getting te id of the calendar 'fin'
result=remainder.calendarList().list().execute()
for calendar in result['items']:
                if calendar['summary']== 'fin':
                    calendarId= calendar['id']

page_token = None

#delete all events 
while True:
    event=remainder.events().list(calendarId=calendarId, pageToken=page_token).execute()

    for index,items in enumerate(event['items']):
        remainder.events().delete(calendarId=calendarId,eventId= items['id'] ).execute()
        print(index)
    page_token = event.get('nextPageToken')
    print('iteration ended')
    
    if not page_token:
        print('all deleted')
        break