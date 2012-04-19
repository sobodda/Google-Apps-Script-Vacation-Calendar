##Background

This script collects all events on one calendar and summarizes them in one daily event on another calendar and daily announcement on a Google Site. It is intended for use for vacation / sick reporting for organizations using Google Apps for Education or Business.  More info and screenshots in the [Google Apps Developer blog post](http://googleappsdeveloper.blogspot.com/2011/05/brown-universitys-it-dept-uses-apps.html)



###tl;dr
Once the script is set up, everyone in your target scope will invite a specific email address to their vacation / sick events (configured as `sourceCalName` variable). To make this really easy for your co-workers, create a service account in your domain to be invited to events (e.g., `i-am-out@mydomain.com`).

The script will scan the calendar that everyone is inviting. It will take the events on that calendar for each day and consolidate them into one daily event on a second calendar (identified in the `destinationCalName variable` in the script). It will also post an announcement to a Google Site announcement page (configured as the `outPage` variable).


###How to install

This assumes a generic service account is used - so that the events and posts are not created by a real user.

- Create a generic service account (e.g., `i-am-out@mydomain.com`) 

- Sign on with that account.

The service account primary calendar will be the one that people invite to their vacation events. 

Create a second calendar on which the condensed daily event will be created,  *Daily Out of Office Report*. The name of this other calendar will be stored in the `destinationCalName` variable in the script.

- Create a Google spreadsheet.

- In the Google Spreadsheet, go into the script editor and paste the .js script. 

- Update the following variables 

>	`sourcecalname`

>	`destinationcalname`

> 	`siteupdatenotify`

>	 `outpage`

- Create a new trigger for `automateGone` 
	- time-driven / hour timer / every hour

- Create a new trigger for `automateGoneDaily` 
	- time-driven / day timer / 9-10am (or your desired time)

When you set up these triggers, googleAppsScript will **prompt to confirm** the script access to calendar / mail / sites / docs

####Bonus Configuration
You can customize the getEventType function for the types of out events in your environment.  

Out of the box, pattern matching for `sick`,  `vacation`,  `conference/training`, or just 'out' (the default)

To test, invite the email address to an event and then run automateGone - you should see an event created on the second calendar