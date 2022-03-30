#Author: Gurarsh Singh
#Date: March 30, 2022

import requests
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

#Tokens
slackAPIToken = "Enter your Token"

#
#Google Sheets API Authentication Starts..........
#

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of the spreadsheets.
Spreadsheet_ID = 'Enter Your Spreadsheet ID'

Write_RANGE = 'Enter Range'

Read_RANGE =  'Example = IDs!A1:10'

# Getting the required credentials for the API
creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.json', 'w') as token:
        token.write(creds.to_json())

sheets_service = build('sheets', 'v4', credentials=creds)

#
#End.........................
#

# Data to be enetered in Sheets 
new_values =[
    # cell values
    ['Channel ID','Result']
]

# Function for exproting data to Google Sheet
def exportSheet(values):
    # Sending the data to selected sheet
    body={
        'values': values
    }

    try:
        new_result = sheets_service.spreadsheets().values().append(spreadsheetId=Spreadsheet_ID,
                                    range=Write_RANGE,valueInputOption='RAW', body=body).execute()
        print('Google Sheet Exported Successfully\n')
    except:
            print("Error Occured")
    return 0

# Getting Channel IDs from Google Sheet using Sheets API and Calling Slack APIs to Archieve them
def getChannelIds_Archieve():
    request = sheets_service.spreadsheets().values().get(spreadsheetId=Spreadsheet_ID, range=Read_RANGE,)
    response = request.execute()
    

    #Printing Values
    for data in response['values']:
        print(data[0])
        # Considering the IT Alerts app is not integrated with Channel
        try:
            result = joinChannel(data[0])

            if result is True:
                archieveChannel(data[0])
            else:
                # For Writing Logs and saving in Sheets 
                new_values.append([data[0],"Error While Joining the Channel in ChannelID Function"])
        # Considering the IT Alerts app is integrated with Channel
        except:
                try:
                    archieveChannel(data[0])
                except:
                    # For Writing Logs and saving in Sheets 
                    new_values.append([data[0],"Error While Joining the Channel in ChannelID Function"])


    
    # Wrinting and saving Logs in Google Sheet
    exportSheet(new_values)
    print("Logs Exported Successfully")


#URL for Slack conservation APIs Call
slackJoinChannel = "https://slack.com/api/conversations.join"
slackArchieveChannel = "https://slack.com/api/conversations.archive"

# Function to connect your Slack App to the Public Channel
def joinChannel(channelID):
    #Header, Parameters and payload data
    payload={
        'channel':channelID
    }
    headers={
        'Content-Type':'application/x-www-form-urlencoded',
        'limit':'1',
        'Authorization':'Bearer '+slackAPIToken,
    }
    params={}
    try:
        # Making POST request for channel joining
        response = requests.request("POST", slackJoinChannel, headers=headers, data=payload,params=params)
        response_data = response.json()
        result=(response_data['ok'])
        return result
    except:
        print('Error Occured while joining the channel in Join Channel Function')

def archieveChannel(channelID):
    #Header, Parameters and payload data
    payload={
        'channel':channelID
    }
    headers={
        'Content-Type':'application/x-www-form-urlencoded',
        'limit':'1',
        'Authorization':'Bearer '+slackAPIToken,
    }
    params={}
    try:
        response = requests.request("POST", slackArchieveChannel, headers=headers, data=payload,params=params)
        print("Channel Archieved")
        # For Writing Logs and saving in Sheets 
        new_values.append([channelID,"Channel Archieved"])
    except:
        new_values.append([channelID,"Error while archieving the channel"])


# Calling the main function
def main():
    # Calling the APIs requests
    getChannelIds_Archieve()


if __name__ == "__main__":
    main()