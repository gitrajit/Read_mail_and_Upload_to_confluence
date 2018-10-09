import json
import requests
import time
import os
from datetime import datetime
now = datetime.now()
weekno= now.strftime('%U')
import win32com.client

import logging
logFormatter = logging.Formatter('%(asctime)s %(name)s %(levelname)-8s %(message)s' ,datefmt='%a %d %b %Y %H:%M:%S', )
log = logging.getLogger("ReportAutomation")
log.setLevel(logging.INFO)
consoleHandler = logging.StreamHandler()
consoleHandler.setFormatter(logFormatter)
log.addHandler(consoleHandler)

def writingToConfluence(body, VendorName):
    value = body
    print VendorName
    log.debug("Processing the mail body to pass to the confluence API")
    vendor_mapping = {'CISCO':662524717, 'JUNIPER':662524722, 'PALO':662524727, 'F5':687445269}
    Vendor_id=vendor_mapping.get(VendorName.upper(), 691930641)


    #print body, VendorName
    try:
        list1 = list()
        val1= value.replace('\r','')
        val1 = val1.replace('&', 'and')
        val1 = val1.replace('@', ' ')
        val1 = val1.replace('<', ' ')
        val1 = val1.replace('>', ' ')
        val1 = val1.replace('=', ' ')
        list1 = val1.split('\n')

        try:
            for i in range(0,len(list1)):

                if "https://" in list1[i]:

                    list1.remove(list1[i])

                else:
                    list1[i] = "<p>" + list1[i] + "</p>"

        except Exception as e:
            log.warn(e)

        log.info("Forming a single string")
        str2 =''.join(e.encode('utf-8').strip() for e in list1)
        #str2 = ''.join(str(e) for e in list1)

        #print str2
        log.debug("The string : "+ str2)

    except Exception as e:
        log.error(e)

    try:
        now = datetime.now()
        url = "https://localhost/confluence/rest/api/content"
        headers = {"content-type": "application/json"}
        payload = {"type": "page",
                   "title": "Weekly Reports WC:" + weekno + " Dated:" + now.strftime('%d-%m-%Y %H:%M:%S'),
                   "space": {
                       "key": "ICS"
                   },
                   "ancestors":
                       [{"id": Vendor_id}],
                   "body": {
                       "storage": {
                           "representation": "storage",
                           "value": "<p><table>"
                                    "<tr><th>Reports</th></tr>"
                                    "<tr><td><p>" + str2 + "</p></td></tr>"

                                                                        "</table></p>"
                       }
                   },
                   # 'version': {'number': 15}
                   # "history": {
                   #    "createdDate": '2017-01-10T14:00:00.000Z'
                   # },

                   }
        log.info("Calling confluence API for weekno " + weekno + " and dated " + now.strftime('%d-%m-%Y %H:%M:%S'))
        response = requests.post(url, data=json.dumps(payload),
                                 headers=headers,
                                 auth=('username', 'password'))
        #print response.json()
        # print response.json()
        log.debug(json.dumps(response.json(), indent=4, sort_keys=True))
        if response.status_code == 200:
            val = response.text
            pageid = int(val[2:30].split(",")[0].split(":")[1].replace('"', ''))
            log.info("***********New child page created Successfully !!!!! with id " + str(pageid))

            return pageid

        else:
            log.error("Got Error while creating new child page. Error code : " + str(response.status_code))
            exit()



    except Exception as e:
        log.error(e)



def attach_to_confluence(name,newpageno):
    try:
        now = datetime.now()
        filename=name
        pageno = newpageno
        #vendor_mapping = {'CISCO': 20873235, 'JUNIPER': 20873239, 'PALO': 20873241}
        #Vendor_id = vendor_mapping.get(vendorname, 20316182)
        log.info("Preparing the attachment to be uploaded to the confluence")
        with open(filename, 'rb') as f:
            cont = f
            url = "https://localhost/confluence/rest/api/content/{}/child/attachment".format(pageno)

           
            response = requests.post(
                url,
                files={
                    'file': (filename, cont,),
                    'comment': "WeeKly Report for week Number " + str(weekno) + " Dated:" + str(now),
                    'minorEdit': "True"
                },
                auth=('Username', 'Password'),
                headers={'X-Atlassian-Token': 'no-check'})

            #print response.json()
            #print json.dumps(response.json(), indent=4, sort_keys=True)
            log.debug(json.dumps(response.json(), indent=4, sort_keys=True))
            if response.status_code == 200:
                log.info("*********** Attachment uploaded successfully!!!!!")
            else:
                log.error("Got Error while uploading attachment. Error code : " + str(response.status_code))


    except Exception as e:
        log.error(e)



try:
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    messages = inbox.Items
    #print 'total messages: ', len(messages)
    log.info("Reading Mail from outlook")
    message = messages.GetFirst()
    #Test is sub folder created in the outlook
    subfolder = inbox.Folders.Item("Test")
    log.debug("Reading mail from inbox.folder")
    subfolderitems = subfolder.Items
    message = subfolderitems.GetFirst()
    for message in subfolderitems:

        if message.Unread == True:
            log.info("New Unread mail found! and Reading the content.")
            subject = message.Subject.split(" ")[0]
            log.info("This mail is from Vendor " + subject)
            #received_time = str(message.ReceivedTime)
            #print subject, datetime.strptime(received_time, '%m/%d/%y %H:%M:%S')
            body = message.Body
            #message.Unread = False
            log.info("Calling Confluence API to create a new page with the mail body.")
            newpageno = writingToConfluence(body, subject)
            log.info("Checking for the attachment.")
            if message.Attachments:
                for attachments in message.Attachments:
                    if "image" not in str(attachments.Filename):
                        attachments.SaveASFile(os.getcwd() + '\\' + str(attachments))
                        filename= attachments.Filename
                        log.info("Attachment found!!! with name " + filename )
                        attach_to_confluence(filename,newpageno)

            else:
                log.info("No attachment found")
            #message.Unread = False



finally:
    log.info('*********************connection closed*****************************')
