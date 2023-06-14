#import pandas as pd
#import glob
import smtplib
from datetime import datetime
from datetime import timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
#import sqlalchemy
import os
#import psycopg2
#from psycopg2 import Error
#from connect_database import run_query_snowflake


sender_address= 'stephane.nichanian@uk.dk.com'
sender_pass = 'xmctsaucgfmvfqnj'



#Create SMTP session for sending the mail
session = smtplib.SMTP("smtp.gmail.com", 587) #use gmail with port
session.ehlo()
session.starttls() #enable security


# In[77]:


session.login(sender_address, sender_pass) #login with mail_id and password
session.select('Inbox')

type, data = session.search(None, 'ALL')
mail_ids = data[0]
id_list = mail_ids.split()

print(id_list)
session.quit()