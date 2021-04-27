from __future__ import print_function  ## this one must occur at the beginning of the file
import pandas as pd
import numpy as np
import pyodbc
from datetime import date
from cryptography.fernet import Fernet
import smtplib
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from fuzzywuzzy import fuzz


# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID_input = '1HuTEXHI16sepY06e6HSbNrdQSE6mBhUb1TPOnrCNsj8'
SAMPLE_RANGE_NAME = 'request!A1:I'


def main():
    global values_input, service
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES) #  JSON file downloaded after enabling the google sheet API,store in the same folder where our code will be saved.
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result_input = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID_input,
                                range=SAMPLE_RANGE_NAME).execute()
    values_input = result_input.get('values', [])

    
    if not values_input and not values_expansion:
        print('No data found.')

main()

# get dataframe from the table in google sheet
df=pd.DataFrame(values_input[1:], columns=values_input[0])

# remove middle names and change to upper cases
df["Vehicle Coordinator's Name:"]=(df["Vehicle Coordinator's Name:"].str.split(' ').str[0].str.upper()+' '+df["Vehicle Coordinator's Name:"].str.split(' ').str[-1].str.upper())

# get coordinator names from email and change to upper case
df['coordinator']=df["Vehicle Coordinator's Email:"].str.split('@').str[0].str.replace('.',' ')
df['coordinator']=df['coordinator'].str.replace('_',' ').str.upper()


# set up sql server connection
sql_conn = pyodbc.connect('DRIVER={ODBC Driver 13 for SQL Server}; SERVER=server; DATABASE=database;UID=prsbiuser;PWD=pwd; Trusted_Connection=no')

# access HR dataset in sql server to get a list of correct employees'names
sql_query = " select * from [db].[Hierarchy_permanent]"
sql_query1 = " select * from [db].[PSEMPLOYEE]"
df_sql = pd.read_sql(sql_query, sql_conn)
df_sql_ora=pd.read_sql(sql_query1, sql_conn)

# remove middle name in HR dataset
df_sql_ora['full']=df_sql_ora.Firstname.str.split(' ').str[0]+' '+df_sql_ora.LastName


# make a list of correct employee names and sort alphabetically
right_name=(df_sql.Employee_name.str.split(' ').str[0]+' '+df_sql.Employee_name.str.split(' ').str[-1]).unique().tolist()
right_name.sort()

# get a list of employee names on google sheet that presents with typos 
google_name=(df["coordinator"].str.split(' ').str[0]+' '+df["coordinator"].str.split(' ').str[-1]).unique()


pairs_full=dict()
for i in google_name:
    for j in right_name:
        Str2=j
        # iterate through google name list and HR names list and use fuzzy ratio to determine the similarity between the two names
        Ratio = fuzz.ratio(Str1,Str2)

        # if a name in google sheet matches a nmae in HR name list 100%, we add that name pair to the pairs_full dictionary  
        if Ratio==100 :
            pairs_full[i]=j
            break

        # if the a name matches 70%, but part of the first name and part of the last name match HR name list, we also add it to the dictionary   
        if Ratio>=70 
            and ((Str1[:3]==Str2[:3] and Str1.split(' ')[1][:3]==Str2.split(' ')[1][:3]) 
            or (Str1[:2]==Str2[:2] and Str1[-3:]==Str2[-3:]) 
            or (Str1.split(' ')[0][-3:]==Str2.split(' ')[0][-3:] and Str1[-3:]==Str2[-3:])):
            pairs_full[i]=j
            break

# replace names with the correct spellings
df["coordinator"]=df["coordinator"].map (pairs_full)

# Allows Python code to execute sql command in a database session.
cursor = sql_conn.cursor()

# set up error handling 
try:
    # create a table in sql database if this table doesn't exist
    cursor.execute("Drop table  [db].[vehicle request]\
                create table [db].[vehicle request] ([Timestamp] datetime, \
        [Select Your Business Unit] varchar(200), [Coordinator] varchar (100), [Email]\
        varchar(100), [Select Rental Start Date] date,[Select Rental End Date] date,[Unit Number]\
        varchar(100))")
    # insert data in each row
    for index,row in df.iterrows():
        cursor.execute("INSERT INTO [db].[vehicle request]([Timestamp],[Select Your Business Unit],\
        [Coordinator], [Email],  [Select Rental Start Date]\
        ,[Select Rental End Date],[Unit Number]) values (?,?,?,?,?,?,?)", row['Timestamp'],  row['Select Your Business Unit'], \
            row["coordinator"],row["Vehicle Coordinator's Email:"],row['Select Rental Start Date'],\
            row['Select Rental End Date'], row['Unit Number']) 
        sql_conn.commit()

    sql_conn.close()
    
# send an email for error notification   
except Exception as e:
    msg="""From: Test <test@gmail.com>
    To: abc <abc@gmail.com>
    Subject: This for SMTP e-mail test

    This is a test e-mail message.
    """
    smtp=smtplib.SMTP ("unixmail.ab.ca",25) 
    smtp.sendmail('error@gmail.ca',['abc@gmail.ca'],msg)
  





