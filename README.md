# Gmail Auto-forwarder<br>
**What is Gmail Auto-forwarder?**<br>
Automatically filter and forward targeted received emails accordingly to their recipients and keep Inbox, Sent, and Trash clean after forwarded.<br>
<br>
**What use-case to use Gmail Auto-forwarder?**<br>
Set a single email address or alias as a recipient—gateway email—from many sources to send their notifications instead of multiple individual or group emails as their recipients.<br>
<br>
**Is there something to keep in mind?**<br>
- This script is only executable on [Google Apps Script](https://www.google.com/script/start)
- Please review the Google services [quotas](https://developers.google.com/apps-script/guides/services/quotas)
## Main features
- Use Gmail [search operators](https://support.google.com/mail/answer/7190) to match targeted emails
- Forward the matched emails to multiple recipients
- Remove permanently the forwarded emails from Inbox, Sent, and Trash
- Automation based on [schedule](https://developers.google.com/apps-script/guides/triggers/installable#time-driven_triggers)
- Log the activities during execution
- Configurable to avoid Google services [limitations](https://developers.google.com/apps-script/reference/gmail/gmail-app#searchquery) (Optional)<br>
  **Note:** Changes to the configuration are effective on the upcoming scheduled execution.
## Getting started
**Prerequisites:**
```
1. Create a project.
2. Create a file or use the pre-existed one.
   [Optional] If using the pre-existed one, clear its content.
3. Add a new service and choose 'Gmail API'.
```
**Installation:**
```
1. Copy the script from 'gmail_auto-forwarder.js' and paste it on the file.
2. Choose 'setTrigger' to install the script in scheduled automation.
3. Authorize this project to access the Google account.
```
**Uninstall:**
```
1. Choose 'unsetTrigger' to uninstall the script.
2. Access the Google account page and navigate to:
   > Data & Privacy
     > Data from apps and services you use
       > Third-party apps & services
         > The project name
3. Choose 'Delete all connections you have with the project' and confirm.
```
---
**Ade Destrianto**<br/>
https://destrianetworks.id
