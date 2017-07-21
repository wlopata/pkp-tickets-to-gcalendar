# pkp-tickets-to-gcalendar

This script processes train tickets bought on [intercity.pl](http://www.intercity.pl) that are sent to your Gmail inbox as PDF attachments. Script saves the tickets in Google Drive and creates a Google Calendar event for your trip with a link to the ticket for easy access.

## Installation
 1. Log in to your Google account and go to https://script.google.com/
 2. Copy and paste [the script](https://github.com/lopekpl/pkp-tickets-to-gcalendar/blob/master/Code.gs) into the editor, and save it
 3. Go to *Resources > Advanced Google services...* and enable Drive and Calendar APIs:
 
![advanced google services dialog](https://raw.githubusercontent.com/lopekpl/pkp-tickets-to-gcalendar/master/advanced_google_services.png)
 
 4. Go to *Edit > Current project's triggers* and configure the script to run periodically:

![project's triggers config dialog](https://raw.githubusercontent.com/lopekpl/pkp-tickets-to-gcalendar/master/project_triggers_config.png)

 5. Click *Run > exportTicketsToCalendar* to verify that the script works
