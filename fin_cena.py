import xlwings as xw
import pandas as pd
from dateutil.relativedelta import relativedelta
from datetime import datetime, timedelta
import math
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pickle
import os.path

pd.options.display.max_rows = 999
pd.set_option('display.max_columns', None)

def buildCalendarApi():
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

            #creating a list for storing the description of all events to identify each event
            eventDescriptionArray=[]
            page_token = None
            print('accessing events...')
            while True:
                event=remainder.events().list(calendarId=calendarId,pageToken=page_token).execute()
                for items in enumerate(event['items']):
                    eventDescriptionArray.append(items['description'])

                page_token=event.get('nextPageToken')

                if not page_token:
                    print('all events accessed')
                    break

            return remainder, calendarId, eventDescriptionArray

#funtion for creating a new event in the google calendar
def create_event(start_time, client_name, number):
  
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
    return event

#Function for Getting all details form the sheet "Client"
def client():
    #storing the sheet to a dataframe
    df = pd.read_excel (filename, sheet_name='Client')
    client_remainder_dict={}
    today=datetime.today()
    isDue=True
    dueMonthList=[]
    
    #creating 30 keys for each client to store due date, principal paid. interest paid
    for i in range(30):
        client_remainder_dict[i+1]={}

    for index,_ in enumerate(df['Date']):
        loanDate=datetime.strptime(df.iloc[index,0],'%d.%m.%Y')
        numberOfMonth=int(df.iloc[index,3])
        principalPaidList=[]

        #creating due dates for each client from date and number of months
        for i in range(numberOfMonth):        
            dueDate=loanDate+relativedelta(months=(i+1))
            if (isinstance(df.loc[index,f'Due Date-{i+1}'],str)== False):
                if(math.isnan(df.loc[index,f'Due Date-{i+1}'])== True):
                    df.at[index,f'Due Date-{i+1}']=dueDate.strftime('%d.%m.20%y')
            client_remainder_dict[i+1][df.iloc[index,1]]=dueDate
            
            # storing the principal paid till date in a list
            if (dueDate<today):
                principalPaidList.append(df.loc[index,f'Principal Paid-{i+1}'])

        position=len(principalPaidList)-1
        isDue=True
        month=0

        #getting the no. of due months unpaid
        while (isDue==True and position>=0):
                if (principalPaidList[position]==0) or (math.isnan(principalPaidList[position])== True):
                    month+=1
                    position-=1
                else:
                    isDue=False
        dueMonthList.append(month)

    return df,client_remainder_dict,dueMonthList

#Function for storing client loan information in the sheet 'client payment'
def client_payment(client,dueMonthList):
    df=pd.DataFrame()
    
    df['Client']=client['Client Name']
    df['Principal Paid'] = client[list(client.filter(regex='Principal Paid-'))].sum(axis=1,numeric_only= True)
    df['Principal Left']=client['Loan Amount']-df['Principal Paid']
    df['Interest Paid'] = client[list(client.filter(regex='Interest Paid-'))].sum(axis=1 ,numeric_only= True)
    df['Interest Left']=(client['Number of Months']*client['Interest'])-df['Interest Paid'] 
    df['Principal + Interest Left']=df['Principal Left']+df['Interest Left']
    df['unpaid due months']=dueMonthList
    return df

#function for getting month intervals and storing respective values
def monthly_accounts(client,expenditure,startDateInFormat,endDate):
    df=pd.DataFrame()

    #exception if staring date is after ending date
    if(startDateInFormat>=endDate):
        df.at[1,'Message']='starting date is after ending date'
        return df

    monthInterval=[]
    loopBreakerMonth=0
    monthCheckDict={}
    for index in df.index:
        monthCheckDict[index]={}
    index=0   
    dueDate3dList=[]

    # storing information for each moth intervals in a dictionary
    while True:
        lastDateInFormat=startDateInFormat+relativedelta(months=1)
        
        #To close the loop when the last date is greater than today's date
        if loopBreakerMonth==1:
            break
        
        # if the given last date is greater than today's date
        if(lastDateInFormat>=endDate):
            loopBreakerMonth=1
            lastDateInFormat=endDate
            string=startDateInFormat.strftime('%d.%m.%Y')+'-'+lastDateInFormat.strftime('%d.%m.%Y')
            monthCheckDict[index]={'x': startDateInFormat,'y': lastDateInFormat, 'z':[],'dc':[],'loan':[],'exp':[]}
         
        #if the last date is within today's date 
        else:
            oneReducedLastDateInFormat=lastDateInFormat-timedelta(days=1)
            string=startDateInFormat.strftime('%d.%m.%Y')+'-'+oneReducedLastDateInFormat.strftime('%d.%m.%Y')
            monthCheckDict[index]={'x': startDateInFormat,'y': oneReducedLastDateInFormat, 'z':[],'dc':[],'loan':[],'exp':[]}
        #storing month intervals as a list of string
        monthInterval.append(string)
        startDateInFormat=lastDateInFormat
        index += 1

    #storing the duedates of each momth interval in a 3D list
    for index,row in client.iterrows():
        for i in range(30):
            dueDateOneByOne=str(row[f'Due Date-{i+1}'])
            if (dueDateOneByOne !='nan') :
                dueDate3dList.append([index,client.columns.get_loc(f'Due Date-{i+1}'),datetime.strptime(dueDateOneByOne,'%d.%m.%Y')])

    #storing the duedates in the respective month intervals in the dictionary
    for xIndex,yIndex,date in dueDate3dList:
        for key in monthCheckDict:
            if(date >=monthCheckDict[key]['x'] and date<=monthCheckDict[key]['y']):
                monthCheckDict[key]['z'].append([xIndex,yIndex])

    #storing the dc, loan in the respective month intervals  in the dictionary
    for index,loanDate in enumerate(client['Date']):
        loanDateInFormat=datetime.strptime(loanDate,'%d.%m.%Y')
        for key in monthCheckDict:
            if(loanDateInFormat >=monthCheckDict[key]['x'] and loanDateInFormat<=monthCheckDict[key]['y']):
                monthCheckDict[key]['dc'].append([index,client.columns.get_loc("DC")])
                monthCheckDict[key]['loan'].append([index,client.columns.get_loc("Loan Amount")])

    #storing the  in the respective month intervals  in the dictionary
    for index,expenditureDate in enumerate(expenditure['Date']):
        expenditureDateInFormat=datetime.strptime(expenditureDate,'%d.%m.%Y')
        for key in monthCheckDict:
            if(expenditureDateInFormat >=monthCheckDict[key]['x'] and expenditureDateInFormat<=monthCheckDict[key]['y']):
                monthCheckDict[key]['exp'].append([index,client.columns.get_loc("Loan Amount")])

    #computation using the dictionary
    principalReceived=[]
    interestReceived=[]
    dcReceived=[]
    expenditureAmount=[]
    loanGiven=[]
    for key in monthCheckDict:
        principalReceived.append(0)
        interestReceived.append(0)
        dcReceived.append(0)
        expenditureAmount.append(0)
        loanGiven.append(0)
        for i in monthCheckDict[key]['z']:
            if(math.isnan(client.iloc[i[0],i[1]+1]) == False):
                principalReceived[key]+=client.iloc[i[0],i[1]+1]
            if(math.isnan(client.iloc[i[0],i[1]+2]) == False):
                interestReceived[key]+=client.iloc[i[0],i[1]+2]
        for j in monthCheckDict[key]['dc']:
            dcReceived[key]+=client.iloc[j[0],j[1]]
        for k in monthCheckDict[key]['exp']:
            expenditureAmount[key]+=expenditure.iloc[k[0],k[1]]
        for m in monthCheckDict[key]['loan']:
            loanGiven[key]+=client.iloc[m[0],m[1]]

    df['Month']=monthInterval
    df['Principal Received']=principalReceived
    df['Interest Received']=interestReceived
    df['DC Received']=dcReceived
    df['Expenditure']=expenditureAmount
    df['Profit']=df['Interest Received']+df['DC Received']-df['Expenditure']
    df['Loan Given']=loanGiven
   
    return df

#function for storing the result information in the dataframe of 'monthly accounts' and 'result'
def result(client,expenditure,monthly_accounts):

    df=pd.DataFrame()

    #exception if staring date is after ending date
    if(monthly_accounts.iloc[0,0] =='starting date is after ending date'):
        df.at[1,'Message']='engaluke alvava'
        return df

    resultArray=['Loan Given', 'Principal Received','Interest Received','DC Received','Expenditure', 'profit','Amount In Hand','Principal to be Received']
    resultMonthValueArray=[]
    resultTotalValueArray=[]

    #storing month wise data
    givenMonthLoanGiven=monthly_accounts['Loan Given'].sum()
    givenMonthPrincipalReceived=monthly_accounts['Principal Received'].sum()
    givenMonthInterestReceived=monthly_accounts['Interest Received'].sum() 
    givenMonthDCReceived=monthly_accounts['DC Received'].sum() 
    givenMonthExpenditure=monthly_accounts['Expenditure'].sum()
    givenMonthProfit=monthly_accounts['Profit'].sum()
    givenMonthAmountInHand=givenMonthPrincipalReceived+givenMonthProfit

    #storing total data
    totalLoanGiven=client['Loan Amount'].sum()
    totalPrincipalReceived=client_payment['Principal Paid'].sum()
    totalInterestReceived=client_payment['Interest Paid'].sum()
    totalDCReceived=client['DC'].sum()
    totalExpenditure=expenditure['Amount'].sum()
    totalProfit=totalInterestReceived+totalDCReceived-totalExpenditure
    totalAmountInHand= totalPrincipalReceived + totalProfit
    totalPrincipalToBeReceived= totalLoanGiven - totalPrincipalReceived

    resultMonthValueArray.extend([givenMonthLoanGiven,givenMonthPrincipalReceived,givenMonthInterestReceived,givenMonthDCReceived,givenMonthExpenditure,givenMonthProfit,givenMonthAmountInHand,'nil'])
    resultTotalValueArray.extend([totalLoanGiven,totalPrincipalReceived,totalInterestReceived,totalDCReceived,totalExpenditure,totalProfit,totalAmountInHand,totalPrincipalToBeReceived])

    df['Result']=resultArray
    df['Values For Given Months']=resultMonthValueArray
    df['Total Values']=resultTotalValueArray

    return df

#function for getting the input dates from the sheet 'Date'
def getDate():
    df = pd.read_excel (filename, sheet_name='Date').fillna('')
    if(len(df['Start Date'])>0  and df.loc[0,'Start Date']!=''):
        startDate=datetime.strptime(df.loc[0,'Start Date'],'%d.%m.%Y')
    else:
        startDate=datetime.strptime('07.01.2022','%d.%m.%Y')
    if(len(df['End Date'])>0 and df.loc[0,'End Date']!=''):
        endDate=datetime.strptime(df.loc[0,'End Date'],'%d.%m.%Y')
    else:
        endDate=datetime.today()
    return startDate,endDate

# main funtion starts
filename=r'.\\cena.xlsx'
client,client_remainder_dict,dueMonthList=client()
remainder, calendarId, eventDescriptionArray=buildCalendarApi()

#Adding events in the google calendar
for month in client_remainder_dict:
    for clientName in client_remainder_dict[month]:
        descriptionString=f'{clientName} - due {month}'
        if descriptionString not in eventDescriptionArray:
            event=create_event(client_remainder_dict[month][clientName], clientName , month)
            remainder.events().insert(calendarId= calendarId, body=event).execute()
            print('event added')
print('all events added')
expenditure= pd.read_excel (filename, sheet_name='Expenditure')
client_payment=client_payment(client,dueMonthList)
startDate,endDate=getDate()
print(startDate,endDate)
monthly_accounts=monthly_accounts(client,expenditure,startDate,endDate)
result=result(client,expenditure,monthly_accounts)

#storing all the sheets in the excel file
print('updating excel file')
#opening excel writer app
app = xw.App(visible=False)
xw.Book(filename).close()
wb = xw.Book(filename)

ws1 = wb.sheets["Client"]
ws2 = wb.sheets["Expenditure"]
ws3 = wb.sheets["Client Payment"]
ws4 = wb.sheets["Monthly Accounts"]
ws5 = wb.sheets["Result"]

#clearing all the sheets for once
ws1.clear()
ws2.clear()
ws3.clear()
ws4.clear()
ws5.clear()

#stroing the dataframe into sheets
ws1["A1"].options(pd.DataFrame, header=1, index=False, expand='table').value = client
ws2["A1"].options(pd.DataFrame, header=1, index=False, expand='table').value = expenditure
ws3["A1"].options(pd.DataFrame, header=1, index=False, expand='table').value =client_payment
ws4["A1"].options(pd.DataFrame, header=1, index=False, expand='table').value =monthly_accounts
ws5["A1"].options(pd.DataFrame, header=1, index=False, expand='table').value =result
wb.save(r'.\\cena.xlsx')
print('file saved')

app.quit()



