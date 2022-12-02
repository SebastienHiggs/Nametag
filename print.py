from __future__ import print_function

from googleapiclient import discovery
from google.oauth2 import service_account

from datetime import datetime,timedelta
import time
import random
import sys

from win32com.client import Dispatch
import pathlib

filePath = pathlib.Path('./Label/Label2/print.label')




CLIENT_SECRET_FILE = './Label/Label2/keys.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SAMPLE_SPREADSHEET_ID = '1vTEk_45hmLh-1r5_rivIBP7BJD2sHmZJGYF9J1PstnM'

creds = None
creds = service_account.Credentials.from_service_account_file(
        CLIENT_SECRET_FILE, scopes=SCOPES)


service = discovery.build('sheets', 'v4', credentials=creds)

spreadsheets = service.spreadsheets()

def add_sheets(gsheet_id, sheet_name): #ONLY MAke new sheet if there isn't one
    try:
        print("win")
        request_body = {
            'requests': [{
                'addSheet': {
                    'properties': {
                        'title': sheet_name,
                    }
                }
            }]
        }

        response = spreadsheets.batchUpdate(
            spreadsheetId=gsheet_id,
            body=request_body
        ).execute()

        return response
    except Exception as e:
        print("fail")
        print(e)

resultDate = datetime.now()

todate = resultDate.strftime("%d/%m")

add_sheets(SAMPLE_SPREADSHEET_ID, todate)

#RUN first time code where it checks the first 1000 lines
cell = todate + "!A1:C1000"
print(cell)
sheet = service.spreadsheets()
result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,range=cell).execute()
v = list(result.values())
if len(v) > 2:
    currentRow = len(v[2]) + 1
else:
    currentRow = 1
print("Starting at row " + str(currentRow))
count = 0

while True:
    printerCOM = Dispatch('Dymo.DymoAddIn')
    printers = printerCOM.getDymoPrinters()
    #print(printers)
    theList = []
    temp = ''
    onlinePrinters = []
    for x in range(len(printers)):
        if printers[x] != '|':
            temp = temp + printers[x]
        else:
            theList.append(temp)
            temp = ''
    theList.append(temp)
    for x in theList:
        if printerCOM.isPrinterOnline(x) == True:
            onlinePrinters.append(x)
    
    print(len(onlinePrinters))
    printerNumber = random.randrange(len(onlinePrinters))
    print(printerNumber)

    myPrinter = onlinePrinters[printerNumber] #'DYMO LabelWriter 450 TURBO'
    print(myPrinter)
    printerCOM.selectPrinter(myPrinter)
    printerCOM.Open2(filePath)
    printerLabel = Dispatch('Dymo.DymoLabels')
    # Set what row we're looking at
    cell = todate + "!A" + str(currentRow) + ":C" + str(currentRow)
    print(cell)
    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=cell).execute()
    print(result)
    if "values" in result.keys():
        # Read the dict
        firstName = result['values'][0][0]
        lastName = result['values'][0][1]
        Date = result['values'][0][2]
        print(result)

        resultDate = datetime.strptime(Date,'%m/%d/%Y %H:%M:%S')

        #subtract 4 seconds from the current time
        minAgo = datetime.now() - timedelta(seconds=4)
        
        print("Today's date is", resultDate, minAgo, Date)

        #See if the label was submitted less than 15 seconds ago
        if resultDate > minAgo and firstName != "":
            # Insert CODE HERE
            print("printing the label", firstName, lastName)
            rename = printerLabel.SetField('TEXT', firstName)
            rename = printerLabel.SetField('TEXT_1', lastName)
            printerCOM.StartPrintJob()
            printerCOM.Print(1,False)
            printerCOM.EndPrintJob()
            # Insert CODE HERE
            currentRow = currentRow + 1
        else:
            currentRow = currentRow + 1
            print("OLD NAME")
    else:
        print("NO DATA")
    
    count = count+1
    print(count)
    time.sleep(1)
