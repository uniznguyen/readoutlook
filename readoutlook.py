import win32com.client
import re



outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# GetDefaultFolder with index = 6 is the Inbox
inbox = outlook.GetDeFaultFolder(6)

# go to subfolder named Route Pad Sales in Inbox
Routepad = inbox.Folders("Route Pad sales")

# filter emails where received time from 01/01/2018
messages = Routepad.Items.Restrict("[ReceivedTime] >= '01/01/2018'")

# sort by received time
messages.Sort("[ReceivedTime]")

# loop through messages
if messages:
    for message in messages:
        
        #get the body of each emails
        body_content = message.body

        # use regular expression to find invoice date of 2017
        result = re.search(r'\d{2}\/\d{2}\/(17)',body_content)

        # if found the sync in 2018 but invoice dated 2017, print the subject line. 
        if result:
            print (message.subject, result.group())
