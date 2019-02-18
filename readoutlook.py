import win32com.client
import re
from datetime import datetime


outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# GetDefaultFolder with index = 6 is the Inbox
inbox = outlook.GetDeFaultFolder(6)

# go to subfolder named Route Pad Sales in Inbox
Routepad = inbox.Folders("Route Pad sales")

# filter emails where received time from 01/01/2018
messages = Routepad.Items.Restrict("[ReceivedTime] >= '01/01/2019' And [ReceivedTime] <= '12/31/2019'")

# sort by received time
messages.Sort("[ReceivedTime]")

msg = []

# loop through messages
if messages:
	for message in messages:
		if "MCNi Remote Sales Processing for JT" in message.subject:
			body_content = message.body
			invoice_dates = re.findall(r'\d{2}\/\d{2}\/\d{2}',body_content)
			sent_date = re.findall(r'\d{2}\/\d{2}\/\d{4}',message.subject)
			for i in sent_date:	
				for y in invoice_dates:
					invoice_date = datetime.strptime(y,'%m/%d/%y').date()
					sent_date = datetime.strptime(i,'%m/%d/%Y').date()
					send_month = sent_date.strftime('%B-%Y')
					invoice_month = invoice_date.strftime('%B-%Y')
					day_diff = (invoice_date - sent_date)

					if send_month != invoice_month:
						print (message.subject, f'Sent Month: {send_month}, Invoice Month: {invoice_month}, Invoice Date: {invoice_date}, Days late: {day_diff}')
					
				
    		
    			
    		
