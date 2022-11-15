from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
import pickle
import pprint
# from datetime import datetime, timedelta
import datetime
from google.auth.transport.requests import Request
import os.path

credentials=None

if os.path.exists(r"D:\Shoban\programming\python\cena fin\token.pickle"):
                print("Accessing credentials...")
                with open("token.pickle",'rb') as file:
                    credentials= pickle.load(file)


if not credentials or not credentials.valid:
    if credentials and credentials.expired and credentials.refresh_token:
        try:
            print("Refreshing access token....")
            credentials.refresh(Request())
        except:
                print('not able to refresh token')
                print("Fetching new access token....")
                flow = InstalledAppFlow.from_client_secrets_file(
                    "client_secret.json", scopes=['https://www.googleapis.com/auth/calendar']
                    )
                flow.run_local_server(
                    port=8080, authorization_prompt_message=""
                    )
                credentials = flow.credentials
                with open("token.pickle",'wb') as file:
                    print("Saving credentials for future use...")
                    pickle.dump(credentials,file)
# flow = InstalledAppFlow.from_client_secrets_file('client_secret.json',scopes=scopes)
# credentials=flow.run_console()
# credentials= '<google.oauth2.credentials.Credentials object at 0x0000021E38BE7F70>'
# pickle.dump(credentials, open("token.pkl",'wb'))
# credentials=pickle.load(open('token.pkl','rb'))



def create_event(start_time, client_name, number, calendarId):
  
    event = {
        'summary': client_name,
        'location': 'Thundukkadu, Tamilnadu',
        'description': f'{client_name} - due {number}',
        'start': {
            'dateTime': start_time.strftime("%Y-%m-%dT08:00:%S"),
            'timeZone': 'Asia/Kolkata',
        },
        'end': {
            'dateTime': start_time.strftime("%Y-%m-%dT22:00:%S"),
            'timeZone': 'Asia/Kolkata',
        },
        'reminders': {
            'useDefault': False,
            'overrides': [
                {'method': 'email', 'minutes': 9 * 60},
                {'method': 'popup', 'minutes': 10},
            ],
        },
    }
    print('ok')
    # return remainder.events().insert(calendarId= calendarId, body=event).execute()
remainder=build('calendar', 'v3', credentials=credentials)
result=remainder.calendarList().list().execute()
for calendar in result['items']:
                if calendar['summary']== 'fin':
                    calendarId= calendar['id']

page_token = None

while True:
    event=remainder.events().list(calendarId=calendarId, pageToken=page_token).execute()
    # if(len(event['items'])==0):
    #     print('all deleted')
    #     break
    for index,items in enumerate(event['items']):
        remainder.events().delete(calendarId=calendarId,eventId= items['id'] ).execute()
        print(index)
    page_token = event.get('nextPageToken')
    print('iteration ended')
    if not page_token:
        print('all deleted')
        break