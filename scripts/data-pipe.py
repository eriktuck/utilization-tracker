import datetime
import os
import win32com.client
import gspread
from oauth2client.service_account import ServiceAccountCredentials

file_path = os.path.expanduser("~/Desktop/Utilization_Report")
today = datetime.date.today()

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) 
messages = inbox.Items

for message in messages:
    if (message.Subject == "Utilization Report from Replicon." 
        and message.Senton.date() == today):
        # body_content = message.body
        attachments = message.Attachments
        attachment = attachments.Item(1)
        attachment.SaveAsFile(os.path.join(file_path, str(attachment)))

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

creds = ServiceAccountCredentials.from_json_keyfile_name(
            'secrets/gs_credentials.json', scope
        )
client = gspread.authorize(creds)

gfile = client.open("Utilization-Hours").id  
report_path = os.path.join(file_path, "Utilization Report Daily.csv")
data = open(report_path, 'r').read()

client.import_csv(gfile, data)
