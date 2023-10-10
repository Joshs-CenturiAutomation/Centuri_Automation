#!/usr/bin/env python
# coding: utf-8
import pandas as pd
import win32com.client
import numpy as np

emaildata = pd.read_excel(r"\\AZ99\Accounting\Private\JBusser\00 Automation Production\Email_Que.xlsx")
df = pd.DataFrame(emaildata)
for i in range(0, len(df)):
    row= df.iloc[i]
    ol=win32com.client.Dispatch("outlook.application")
    olmailitem=0x0 #size of the new email
    newmail=ol.CreateItem(olmailitem)
    newmail.Subject= row['Subject']
    newmail.To=row['To']
#     newmail.CC=row['CC']
    newmail.Body= row['Body']
    #Need admin rights to send attachments so the code below will not work.
    # attach='P:\Accounting\Private\JBusser\01 Python\Python SAP\logs\20230228_python.log'
    # newmail.Attachments.Add(attach)
    # To display the mail before sending it
#     newmail.Display() 
    newmail.Send()
    df.loc[i] = np.nan

df.to_excel(r"\\AZ99\Accounting\Private\JBusser\00 Automation Production\Email_Que.xlsx", index =False)    

