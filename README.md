# Read_mail_and_Upload_to_confluence
Read mail from outlook and upload the content and attachement (if any) to the confluence 


Required packages:

```
pip install pywin32com
pip install requests

```

output:

```
Tue 09 Oct 2018 16:02:58 ReportAutomation INFO     Reading Mail from outlook
Tue 09 Oct 2018 16:02:58 ReportAutomation INFO     New Unread mail found! and Reading the content.
Tue 09 Oct 2018 16:02:58 ReportAutomation INFO     This mail is from Vendor PALO
PALO
Tue 09 Oct 2018 16:02:58 ReportAutomation INFO     Calling Confluence API to create a new page with the mail body.
Tue 09 Oct 2018 16:02:58 ReportAutomation WARNING  list index out of range
Tue 09 Oct 2018 16:02:58 ReportAutomation INFO     Forming a single string
Tue 09 Oct 2018 16:02:58 ReportAutomation INFO     Calling confluence API for weekno 40 and dated 09-10-2018 16:02:58
Tue 09 Oct 2018 16:03:00 ReportAutomation INFO     ***********New child page created Successfully !!!!! with id 24936598
Tue 09 Oct 2018 16:03:00 ReportAutomation INFO     Checking for the attachment.
Tue 09 Oct 2018 16:03:00 ReportAutomation INFO     Attachment found!!! with name Bi-weekly update Amadeus Palo Alto Networks 03.05.2018.pdf
Tue 09 Oct 2018 16:03:00 ReportAutomation INFO     Preparing the attachment to be uploaded to the confluence
Tue 09 Oct 2018 16:03:03 ReportAutomation INFO     *********** Attachment uploaded successfully!!!!!
Tue 09 Oct 2018 16:03:03 ReportAutomation INFO     No attachment found
Tue 09 Oct 2018 16:03:03 ReportAutomation INFO     connection closed
```
