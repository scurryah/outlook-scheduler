# outlook-scheduler
Simple VBA + form to create drafts or scheduled emails in Outlook

Create Outlook messages and either store as drafts or push directly to outbox with a scheduled email data. Sample excel sheet included as well.

Does not update email dynamically; Sends one template to multiple addresses; Can specify CC, and up to 2 attachments.

Tested in Windows Excel 2013 / Outlook 2013 and Windows Excel 2016 / Outlook

# Options
- **Browse** Select the .msg template you want to use to send the email
- **Generate Drafts** Generates drafts in Outlook
- **Schedule** Generates drafts then schedules them to be sent. Must send a test draft first. 

# Installation
1. Download the .bas file and import to existing Excel workbook OR use the outlook-scheduler-template
2. Add the Microsoft Outlook 2016 reference through VBA --> Tools --> References
