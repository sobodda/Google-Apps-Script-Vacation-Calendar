/*
Functions:
 automateGone: creates/updates the next 7 days Who is Out events on calendar
 automateGoneDaily: updates the website, creates/updates next 31 days Who is Out 
    events on the calendar
 whoIsOut

Triggers:
 automateGone: run hourly
 automateGoneDaily: run once a day between 9 and 10am
*/

var sourceCalName = "my-calendar-email-address@mydomain.com"; 
    /*calendar which is invited to gone events*/
var destinationCalName = "Sick and Vacation Calendar"; 
    /*calendar that people overlay to view who is out*/
var siteUpdateNotify = "youraddress@mydomain.com"; 
    //email this address when out list posted to google site
var d = new Date();
var timezone = "GMT-" + d.getTimezoneOffset()/60; 
    /*uses Timezone detail from gDoc config, validate this is accurate...*/
var outPage = "https://sites.google.com/my-site-name/who-is-out/out/"; 
    /*the site to which the list is posted*/


function automateGone()
{
	/*this function will be called every hour to update the one event on the 
        official calendar*/
    
	var myDate = new Date(); 
	myDate.setHours(0);
	myDate.setMinutes(0);
	myDate.setSeconds(0);
	myDate.setDate(myDate.getDate()-1);

	for( i=0; i<7; i++)  //update next 7 days
	{
		myDate.setDate(myDate.getDate() + 1); 

		if(myDate.getDay() !== 0 && myDate.getDay() != 6)	//not a weekend
		{ 
			var goneListArray = whoIsOut(myDate);
			//Logger.log(myDate + " " + goneListArray);
			createGoneEvent(goneListArray , myDate);
		}
	}
}

function automateGoneDaily()
{
	/*this function will be called once a day to update the one event on the 
        official calendar*/
    
	var myDate = new Date(); 
	myDate.setHours(0);
	myDate.setMinutes(0);
	myDate.setSeconds(0);
  
	/*update site - will only get updated if there isn't already a post for     
        the date */
    
	var goneListArray = whoIsOut(myDate); 
	updateInternalSite(goneListArray, myDate);
	myDate.setDate(myDate.getDate()-1);
	//update calendar for next 90 days
	for( i=0; i<91; i++)
	{
		myDate.setDate(myDate.getDate() + 1);  
		
		if(myDate.getDay() !== 0 && myDate.getDay() != 6)
		{
			var goneListArray = whoIsOut(myDate); 
                /*goneListArray also declared in func(automateGone), required
                    in this function as well*/
			createGoneEvent(goneListArray , myDate);
		}
	}
}

function whoIsOut(theDate)
{
	var peopleOut = new Array();
        /*suggesting the array literal notation [], need to test*/
	var Calendar = CalendarApp.getCalendarsByName(sourceCalName)[0];
	var events = Calendar.getEventsForDay(theDate);
	    /*cycle through events and add info to peopleOut*/
	var eventLines = ""; //unused var?

	for (var e in events) 
	{
		var event = events[e];
		var duration = (event.getEndTime() - event.getStartTime())/86400000;
			
		if( (theDate >= event.getStartTime() ) || duration<1 ) 	
            /*weed out future all-day events mistakenly pulled by broken 
            getEventsForDay function*/
		{ 

			if(duration>1)	//multi-day event - delete one day from end time
			{ 
				var endDate = event.getEndTime();
				endDate.setDate(endDate.getDate() - 1);  
				var durDisplay =  Utilities.formatDate(event.getStartTime(), 
                    (event.isAllDayEvent() === true ? "GMT" : timezone),
                    "M/d") + " - " + Utilities.formatDate(endDate, 
                    (event.isAllDayEvent() === true ? "GMT" : timezone), "M/d");
			}
			
			else if(duration<1)  //part day
			{
				var durDisplay =  Utilities.formatDate(event.getStartTime(), 
                    (event.isAllDayEvent() === true ? "GMT" : timezone), 
                    "h:mm a") + " - " + Utilities.formatDate(event.getEndTime(),
                    (event.isAllDayEvent() === true ? "GMT" : timezone),
                    "h:mm a");
			}
			
			else  //full day
			{
				var durDisplay = "full day"; 
			}
  
			/*the following will not be necessary because of the new 
            event.getOriginalCalendarId() method but in case someone needs to 
            report that someone else is out and doesn't have rights to their 
            calendar, they can put their email address in the 
            event description*/

			var onBehalf = "";
			//var regexEmail = /^[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$/;

			var regexEmail = /[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}/;

			//Logger.log(event.getDescription() + " " + event.getDescription().search(regexEmail));
			
			if(event.getDescription().search(regexEmail)>-1) 	
                /*workaround for inviting on behalf of - looks for email address 
                in description*/
			{ 
				var descWordArray = event.getDescription().split(" ");

				for (x in descWordArray)
				{

					if(descWordArray[x].search(regexEmail)>-1)
					{
						onBehalf = descWordArray[x]; 
					}
				}
			}
			
			var personEmail = (onBehalf==="") ? 
                event.getOriginalCalendarId() : onBehalf;
			peopleOut.push(personEmail.toLowerCase() + " " +  
                getEventType(event.getTitle()) + " " + durDisplay);
		}
	}
	
	peopleOut.sort();
	return peopleOut;
}

function getEventType(theTitle)
{
	/*keywords to decide whether to list person as sick, vacaction, conference,
    or just out (default)*/
    
	eventType = "";
	
	if(theTitle.search(/sick|doctor|dr\.|dentist|medical/i)>-1)
	{
		eventType = "sick/doctor";
	}
	else if(theTitle.search(/vacation|vaca|vac|holiday/i)>-1)
	{
		eventType = "vacation";   
	}
	else if(theTitle.search(/conference|training|seminar|session|meeting/i)>-1)
	{
		eventType = "conference/training";
	}
	else
	{
		eventType = "out";
	}
	return eventType;
}

function createGoneEvent(goneListArray, theDate)
{
	//this function deletes existing 'who is out' event and adds updated one
	if(goneListArray.length===0)	//do not post if nobody is out
	{
		return false;
	}
	
	var goneList = goneListArray.join("\n");
	theDate.setDate(theDate.getDate());
	if(goneList.length===0)
	{
		return false;
	}

	var official = CalendarApp.openByName(destinationCalName);
	var events = official.getEventsForDay(theDate);
  
	for (var e in events) 
	{
		var event = events[e];
		var startDate = Utilities.formatDate(event.getStartTime(),
            (event.isAllDayEvent() === true ? "GMT" : timezone),"M/d/yyyy");
		var theDateshort = Utilities.formatDate(theDate,"GMT","M/d/yyyy");

		if(event.getTitle()=="Who is Out" && (theDateshort>=startDate))	 
            /*need to compare dates because of all day event bug*/
		{
			event.deleteEvent(); //delete existing event before recreating
		}
    }
	
	goneList += "\nLast updated: " + Utilities.formatDate(new Date(), 
        timezone, "EEEE M/d h:mm a");
	var advancedArgs = {description:goneList};
	official.createAllDayEvent("Who is Out", theDate, advancedArgs);  
}

function triggerInternalSiteUpdate()
{
	var myDate = new Date();
	myDate.setHours(0);
	myDate.setMinutes(0);
	myDate.setSeconds(0);
	var goneListArray = whoIsOut(myDate); 
	updateInternalSite(goneListArray , myDate);
}

function updateInternalSite(goneListArray, theDate)
{
	if(goneListArray.length===0||postExists(theDate))
	{
		//if nobody is out, or if the out list has already been posted, do not post
		return false;
	}
  
	var goneList = goneListArray.join("<br/>");
	var theDateshort = Utilities.formatDate(theDate,"GMT","M/d/yyyy");
	var page = SitesApp.getPageByUrl(outPage);
	
	if (page === null)
	{ 
		return false;
	}
	else
	{ 
		var announcement = page.createAnnouncement(
            "Who is Out: " + theDateshort, goneList); 
            /*cannot find 'announcement' in use anywhere else ?depricate? */
		MailApp.sendEmail(
            siteUpdateNotify, "Updated Who is Out Website", "Who is Out: " + 
            theDateshort + " " + goneList);
	}
}

function postExists(theDate)
{
	var post = SitesApp.getPageByUrl(
        outPage + "who-is-out-" + Utilities.formatDate(theDate,
        "GMT","M-d-yyyy"));
        
	if (post === null)  //does not exist
	{ 
        return false;
	}
	    else	//exists
	{ 
        return true;
	}
}
