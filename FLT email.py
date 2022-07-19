from sys import float_info
from unicodedata import name
from pyad import *
from email.message import EmailMessage
from email.mime.multipart import MIMEMultipart
import pyad.adquery
import smtplib
import re

pyad.set_defaults(ldap_server="", username="", password="")
user = pyad.aduser.ADUser.from_cn("")

StoreLookUp = (input)

q = pyad.adquery.ADQuery(StoreLookUp)

q.execute_query(
    attributes = ["name", "Manager"],
    where_clause = "objectClass = '*'",
    base_dn = "OU=,DC=,DC=,DC=)"
)
adStoreMail = []
ManagerName = []
for row in q.get_results():
    for row in q.get_results():
        ManagerName.append(row["Manager"])
        adStoreData.append(row['name'])
        adStoreData.append(row["mail"])
    adStoreData = [x for x in adStoreData if x != None]
    
z = pyad.adquery.ADQuery(ManagerName)

z.execute_query(
    attributes = ["name", "mail", "Manager"],
    where_clause = "objectClass = '*'",
    base_dn = "OU=,OU=,DC=,DC=,DC="
)

adDMData = []

for row in z.get_results():
    for row in z.get_results():
        adDMData.append(row["mail"])
    adDMData = [x for x in adDMData if x != None]

recipients = [adStoreData(1),adDMData]

#--------------------------------------------------------------------------------------------------------

conn = smtplib.SMTP('smtp-mail.outlook.com',587)
type (conn)
conn
conn.ehlo()
conn.starttls
conn.login('','')


fromaddr = 'sdteam@vitaminshoppe.com'
toaddr = recipients
cc = ['']
Escalation_Description = (str)

Greeting = ("Greetings all \n , I hope all is well. I wanted to provide you with an update regarding " + adStoreData)

Issue_Description = input('')
print(Issue_Description)

Business_Impact = input('')
print(Business_Impact)

Next_Action = input('')
print(Next_Action)

EV_Ticket = input(int)
print("EV Ticket #:" +(EV_Ticket)) 

VSN_Ticket = input(int)
print('Vector Ticket #:'+ (VSN_Ticket))
 
message = ( (Greeting) + '\n' + (Issue_Description) + '\n' + (Next_Action) + '\n' + (EV_Ticket) + '\n' + (VSN_Ticket) )

#This variable is going to need to read data from an excel file
#Store, District Manager, Operators, Guia Auguas, Store systems, Sales audit

conn.sendmail(fromaddr, [recipients,cc] (message))
