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
    //calendar which is invited to gone events
var destinationCalName = "Sick and Vacation Calendar"; 
    //calendar that people overlay to view who is out
var siteUpdateNotify = "youraddress@mydomain.com"; 
    //email this address when out list posted to google site
var d = new Date();
var timezone = "GMT-" + d.getTimezoneOffset()/60; 
    //will be GMT-4 in spring, GMT-5 in fall
    //automatic method of DST change? Probably...
var outpage = "https://sites.google.com/my-site-name/who-is-out/out/"; 
    //the site to which the list is posted


function automateGone()
{
	/*this function will be called every hour to update the one event on the 
        official calendar*/
    
	var mydate = new Date(); 
	mydate.setHours(0);
	mydate.setMinutes(0);
	mydate.setSeconds(0);
	mydate.setDate(mydate.getDate()-1);

	for(i=0; i<7; i++) 	 //update next 7 days
	{
		mydate.setDate(mydate.getDate() + 1); 

		if(mydate.getDay() !== 0 && mydate.getDay() != 6)	//not a weekend
		{ 
			var gonelistarray = whoIsOut(mydate);
			//Logger.log(mydate + " " + gonelistarray);
			createGoneEvent(gonelistarray , mydate);
		}
	}
}

function automateGoneDaily()
{
	/*this function will be called once a day to update the one event on the 
        official calendar*/
    
	var mydate = new Date(); 
	mydate.setHours(0);
	mydate.setMinutes(0);
	mydate.setSeconds(0);
  
	/*update site - will only get updated if there isn't already a post for     
        the date */
    
	var gonelistarray = whoIsOut(mydate); 
	updateInternalSite(gonelistarray, mydate);
	mydate.setDate(mydate.getDate()-1);
	//update calendar for next 90 days
	for(i=0; i<91; i++)
	{
		mydate.setDate(mydate.getDate() + 1);  
		
		if(mydate.getDay() !== 0 && mydate.getDay() != 6)
		{
			var gonelistarray = whoIsOut(mydate); //prev declared?
			createGoneEvent(gonelistarray , mydate);
		}
	}
}

function whoIsOut(thedate)
{
	var pplout = new Array();
	var Calendar = CalendarApp.getCalendarsByName(sourceCalName)[0];
	var events = Calendar.getEventsForDay(thedate);
	// cycle through events and add info to pplout 
	var eventLines = ""; //unused var?

	for (var e in events) 
	{
		var event = events[e];
		var duration = (event.getEndTime() - event.getStartTime())/86400000;
			
		if((thedate>=event.getStartTime())||duration<1) 	
            /*weed out future all-day events mistakenly pulled by broken 
            getEventsForDay function*/
		{ 

			if(duration>1)	//multi-day event - delete one day from end time
			{ 
				var enddate = event.getEndTime();
				enddate.setDate(enddate.getDate() - 1);  
				var durdisplay =  Utilities.formatDate(event.getStartTime(), 
                    (event.isAllDayEvent() === true ? "GMT" : timezone),
                    "M/d") + " - " + Utilities.formatDate(enddate, 
                    (event.isAllDayEvent() === true ? "GMT" : timezone), "M/d");
			}
			
			else if(duration<1)  //part day
			{
				var durdisplay =  Utilities.formatDate(event.getStartTime(), 
                    (event.isAllDayEvent() === true ? "GMT" : timezone), 
                    "h:mm a") + " - " + Utilities.formatDate(event.getEndTime(),
                    (event.isAllDayEvent() === true ? "GMT" : timezone),
                    "h:mm a");
			}
			
			else  //full day
			{
				var durdisplay = "full day"; 
			}
  
			/*the following will not be necessary because of the new 
            event.getOriginalCalendarId() method but in case someone needs to 
            report that someone else is out and doesn't have rights to their 
            calendar, they can put their email address in the 
            event description*/

			var onbehalf = "";
			//var regexemail = /^[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$/;

			var regexemail = /[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}/;

			//Logger.log(event.getDescription() + " " + event.getDescription().search(regexemail));
			
			if(event.getDescription().search(regexemail)>-1) 	
                /*workaround for inviting on behalf of - looks for email address 
                in description*/
			{ 
				var descwordarray = event.getDescription().split(" ");

				for (x in descwordarray)
				{

					if(descwordarray[x].search(regexemail)>-1)
					{
						onbehalf = descwordarray[x]; 
					}
				}
			}
			
			var personemail = (onbehalf==="") ? 
                event.getOriginalCalendarId() : onbehalf;
			pplout.push(personemail.toLowerCase() + " " +  
                getEventType(event.getTitle()) + " " + durdisplay);
		}
	}
	
	pplout.sort();
	return pplout;
}

function getEventType(thetitle)
{
	/*keywords to decide whether to list person as sick, vacaction, conference,
    or just out (default)*/
    
	eventType = "";
	
	if(thetitle.search(/sick|doctor|dr\.|dentist|medical/i)>-1)
	{
		eventType = "sick/doctor";
	}
	else if(thetitle.search(/vacation|vaca/i)>-1)
	{
		eventType = "vacation";   
	}
	else if(thetitle.search(/conference|training|seminar|session|meeting/i)>-1)
	{
		eventType = "conference/training";
	}
	else
	{
		eventType = "out";
	}
	return eventType;
}

function createGoneEvent(gonelistarray, thedate)
{
	//this function deletes existing 'who is out' event and adds updated one
	if(gonelistarray.length===0)	//do not post if nobody is out
	{
		return false;
	}
	
	gonelist = gonelistarray.join("\n");
	thedate.setDate(thedate.getDate());
	if(gonelist.length===0)
	{
		return false;
	}

	var official = CalendarApp.openByName(destinationCalName);
	var events = official.getEventsForDay(thedate);
  
	for (var e in events) 
	{
		var event = events[e];
		var startdate = Utilities.formatDate(event.getStartTime(),
            (event.isAllDayEvent() === true ? "GMT" : timezone),"M/d/yyyy");
		var thedateshort = Utilities.formatDate(thedate,"GMT","M/d/yyyy");

		if(event.getTitle()=="Who is Out" && (thedateshort>=startdate))	 
            /*need to compare dates because of all day event bug*/
		{
			event.deleteEvent(); //delete existing event before recreating
		}
    }
	
	gonelist += "\nLast updated: " + Utilities.formatDate(new Date(), 
        timezone, "EEEE M/d h:mm a");
	var advancedArgs = {description:gonelist};
	official.createAllDayEvent("Who is Out", thedate, advancedArgs);  
}

function triggerInternalSiteUpdate()
{
	var mydate = new Date();
	mydate.setHours(0);
	mydate.setMinutes(0);
	mydate.setSeconds(0);
	var gonelistarray = whoIsOut(mydate); 
	updateInternalSite(gonelistarray , mydate);
}

function updateInternalSite(gonelistarray, thedate)
{
	if(gonelistarray.length===0||postExists(thedate))
	{
		//if nobody is out, or if the out list has already been posted, do not post
		return false;
	}
  
	gonelist = gonelistarray.join("<br/>");
	var thedateshort = Utilities.formatDate(thedate,"GMT","M/d/yyyy");
	var page = SitesApp.getPageByUrl(outpage);
	
	if (page === null)
	{ 
		return false;
	}
	else
	{ 
		var announcement = page.createAnnouncement(
            "Who is Out: " + thedateshort, gonelist); 
		MailApp.sendEmail(
            siteUpdateNotify, "Updated Who is Out Website", "Who is Out: " + 
            thedateshort + " " + gonelist);
	}
}

function postExists(thedate)
{
	var post = SitesApp.getPageByUrl(
        outpage + "who-is-out-" + Utilities.formatDate(thedate,
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
