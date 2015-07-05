<%@language="JScript" CODEPAGE="65001"%>
<!--#include file="../strings.asp" -->
<!--#include file="datefuncs.asp" -->
<!--#include file="tourfuncs.asp" -->
<!--#include file="eventfuncs.asp" -->
<%
// Now for any variables local to this page
var newone = false;   // default to update rather than create
var m_title, m_start, m_end, m_finalsday, m_who;
var m_contact, m_starttime, m_endtime, m_finalsstart, m_finalsend;
var m_cost, m_url, m_category
var m_id;   // tournamentid of the tournament being updated/created
var m_newone;  // is this a new tournament or an update of an existing one
var tourObj = new Object();
var resultObj = new Object();
var eventObj = new Object();
var RS, Conn, dbconnect;
var SQL1, SQLText;
var SQL2 = new String("").toString();
var updateSQL = new String("").toString();
var updateSQL2 = new String("").toString();
var m_email, m_event;
var checkboxoff = new String("off").toString();
var checkboxon = new String("on").toString();
var SQLresult = new String("").toString();
var newEventID = -1;
var editURL = new String("").toString();
var dateArray;
var resultObject = new Object();

// var debugging=current_debug_status();
var updating=true;  // if true we will update the database
debugging = false;   // if true we will display debug info

// Process form/querystring parameters
m_id = Trim(new String(Request.Form("tourid")));
if (m_id == "" || m_id =="null" || m_id == "undefined" || m_id == "-1")
{
	m_id = Trim(new String(Request.QueryString("tourid")));
	if (m_id == "" || m_id =="null" || m_id == "undefined" || m_id == "-1")
	{
		m_id = new String("0").toString();
	}
}
// Now get other form/querystring variables
m_title = Trim(new String(Request.Form("tournote")));
if (m_title == "" || m_title =="null" || m_title == "undefined")
{
	m_title = Trim(new String(Request.QueryString("tournote")));
	if (m_title == "" || m_title =="null" || m_title == "undefined")
	{
		m_title = new String("").toString();
	}
}
m_start = Trim(new String(Request.Form("tourstart")));
if (m_start == "" || m_start =="null" || m_start == "undefined")
{
	m_start = Trim(new String(Request.QueryString("tourstart")));
	if (m_start == "" || m_start =="null" || m_start == "undefined")
	{
		m_start = new String("").toString();
	}
}
m_end = Trim(new String(Request.Form("tourend")));
if (m_end == "" || m_end =="null" || m_end == "undefined")
{
	m_end = Trim(new String(Request.QueryString("tourend")));
	if (m_end == "" || m_end =="null" || m_end == "undefined")
	{
		m_end = new String("").toString();
	}
}
m_finalsday = Trim(new String(Request.Form("tourfinalsday")));
if (m_finalsday == "" || m_finalsday =="null" || m_finalsday == "undefined")
{
	m_finalsday = Trim(new String(Request.QueryString("tourfinalsday")));
	if (m_finalsday == "" || m_finalsday =="null" || m_finalsday == "undefined")
	{
		m_finalsday = new String("null").toString();
	}
}
m_who = Trim(new String(Request.Form("tourwho")));
if (m_who == "" || m_who =="null" || m_who == "undefined")
{
	m_who = Trim(new String(Request.QueryString("tourwho")));
	if (m_who == "" || m_who =="null" || m_who == "undefined")
	{
		m_who = new String("null").toString();
	}
}
m_contact = Trim(new String(Request.Form("tourcontact")));
if (m_contact == "" || m_contact =="null" || m_contact == "undefined")
{
	m_contact = Trim(new String(Request.QueryString("tourcontact")));
	if (m_contact == "" || m_contact =="null" || m_contact == "undefined")
	{
		m_contact = new String("Please contact one of our coaches for further information.").toString();
	}
}
m_starttime = Trim(new String(Request.Form("tourstarttime")));
if (m_starttime == "" || m_starttime =="null" || m_starttime == "undefined")
{
	m_starttime = Trim(new String(Request.QueryString("tourstarttime")));
	if (m_starttime == "" || m_starttime =="null" || m_starttime == "undefined")
	{
		m_starttime = new String("").toString();
	}
}
m_endtime = Trim(new String(Request.Form("tourendtime")));
if (m_endtime == "" || m_endtime =="null" || m_endtime == "undefined")
{
	m_endtime = Trim(new String(Request.QueryString("tourendtime")));
	if (m_endtime == "" || m_endtime =="null" || m_endtime == "undefined")
	{
		m_endtime = new String("").toString();
	}
}
m_finalsstart = Trim(new String(Request.Form("finalsstarttime")));
if (debugging) {
	Response.Write("Finals start time as originally received ["+m_finalsstart+"]<br /><br />");
}
if (m_finalsstart == "" || m_finalsstart =="null" || m_finalsstart == "undefined")
{
	m_finalsstart = Trim(new String(Request.QueryString("finalsstarttime")));
	if (m_finalsstart == "" || m_finalsstart =="null" || m_finalsstart == "undefined")
	{
		m_finalsstart = new String("").toString();
	}
}
m_finalsend = Trim(new String(Request.Form("finalsendtime")));
if (debugging) {
	Response.Write("Finals end time as originally received ["+m_finalsend+"]<br /><br />");
}
if (m_finalsend == "" || m_finalsend =="null" || m_finalsend == "undefined")
{
	m_finalsend = Trim(new String(Request.QueryString("finalsendtime")));
	if (m_finalsend == "" || m_finalsend =="null" || m_finalsend == "undefined")
	{
		m_finalsend = new String("").toString();
	}
}
m_cost = Trim(new String(Request.Form("tourcost")));
if (m_cost == "" || m_cost =="null" || m_cost == "undefined")
{
	m_cost = Trim(new String(Request.QueryString("tourcost")));
	if (m_cost == "" || m_cost =="null" || m_cost == "undefined")
	{
		m_cost = new String("").toString();
	}
}
m_url = Trim(new String(Request.Form("toururl")));
if (m_url == "" || m_url =="null" || m_url == "undefined")
{
	m_url = Trim(new String(Request.QueryString("toururl")));
	if (m_url == "" || m_url =="null" || m_url == "undefined")
	{
		m_url = new String("tournaments.html").toString();
	}
}
m_category = Trim(new String(Request.Form("tourcategory")));
if (m_category == "" || m_category =="null" || m_category == "undefined")
{
	m_category = Trim(new String(Request.QueryString("tourcategory")));
	if (m_category == "" || m_category =="null" || m_category == "undefined")
	{
		m_category = new String("JUNIOR").toString();
	}
}

// OK, now pick up the two checkbox fields so we know what workflow 
// we need to do afterwards

m_email = Trim(new String(Request.Form("emailcheckbox")));
if (m_email == "" || m_email =="null" || m_email == "undefined")
{
	m_email = Trim(new String(Request.QueryString("emailcheckbox")));
	if (m_email == "" || m_email =="null" || m_email == "undefined")
	{
		m_email = new String("off").toString();
	}
}

m_event = Trim(new String(Request.Form("eventcheckbox")));
if (m_event == "" || m_event =="null" || m_event == "undefined")
{
	m_event = Trim(new String(Request.QueryString("eventcheckbox")));
	if (m_event == "" || m_event =="null" || m_event == "undefined")
	{
		m_event = new String("off").toString();
	}
}

// ================================================================
// Now do cross field checks
//
// 1.  If end date hasnt been supplied, set to start date
//
if (m_end == "") {
	m_end = new String(m_start);  // One-day tournament assumed
}
//
// 2.  If m_start == m_end then we must have at least a start time
//
if (m_end == m_start) {
	if (m_starttime == "") {
		if (m_endtime == "") {
			// No times specified - defaults to mid-day
			m_starttime = new String("12:00:00");
		} else {
			m_starttime = new String(m_endtime);
		}
	} else {   // We have a start time
		if (m_endtime == "") {
			// No end time specified - defaults to 5pm
			m_endtime = new String(m_starttime);
		} 
	}
}
//
//  3.  If we have a finals day, then we must allocate start end end times
//
if (m_finalsday == "") {
	// No finals day, make sure and set start and end times on finals day to null
	m_finalsstart = new String("");
	m_finalsend = new String("");
} else {
	// We have a finals day
	if (m_finalsstart == "") {
		// No start time - defauls to mid-day
		if (m_finalsend == "") {
			// No times specified - defaults to mid-day
			m_finalsstart = new String("12:00:00");
		} else {
			m_finalsstart = new String(m_finalsend);
		}
	} else {   // We have a start time on finals day
		if (m_finalsend == "") {
			// No end time specified - defaults to 5pm
			m_finalsend = new String("17:00:00");
		} 
	}
}


// ==============================================================================
// if this is a new one, add skeleton record and get its unique id back into m_id
// or now retrieve existing tournament details

if (! (m_id == "0"))  {
	tourObj = getTour(m_id);  // Retrieve the existing tournament record
}  
else {
	// New tournament 
	m_id = newTour();
	tourObj = getTour(m_id);
	if (debugging) {
		Response.Write("We have created a new dummy tournament and the ID = "+m_id+".<br />");
	}
}

// Update tournament Object with data from submitting form

tourObj.tournamentid = m_id;
tourObj.tourtitle = m_title;
tourObj.tourstart = m_start;
tourObj.tourend = m_end;
tourObj.tourfinalsday = m_finalsday;
tourObj.tourwho = m_who;
tourObj.tourcontact = m_contact;
tourObj.tourstarttime = m_starttime;
tourObj.tourendtime = m_endtime;
tourObj.finalsstarttime = m_finalsstart;
tourObj.finalsendtime = m_finalsend;
tourObj.tourcontact = m_contact;
tourObj.tourcost = m_cost;
tourObj.toururl = m_url;
tourObj.tourcategory = m_category;

if (debugging)
{
	Response.Write(printTour(tourObj));
	Response.Write("<br /><br />");
	Response.Write("Email checkbox value: ["+m_email+"]<br />")
	Response.Write("Event checkbox value: ["+m_event+"]<br />")
}

// Now update the tournament record with an id of m_id

resultObject = setTour(tourObj);

if ( ! (resultObject.result)) {
	// Update failed in some way
	// Send out debug info to screen
	Response.Write("<h4>Update failed</h4>");
	Response.Write("<p>Error number:  "+resultObject.errno+"</p>");
	Response.Write("<p>Description:  "+resultObject.description+"</p>");
	Response.Write("<p>SQL used:  "+resultObject.sql+"</p>");
	Response.Write("====================================================");
	Response.End();
}

// Workflow implications - see if we need to send a quick message and/or if 
// we need to update the events calendar (only on new tournament definitions?)

if (debugging)
{
	if (m_email == checkboxon)
		Response.Write("We WILL be sending out quick message re this tournament.<br />");
	if (m_email == checkboxoff)
		Response.Write("We will NOT be sending out quick message re this tournament.<br />");
	if (m_event == checkboxon)
		Response.Write("We WILL be inserting/amending an entry in the event calendar for this tournament.<br />");
	if (m_event == checkboxoff)
		Response.Write("We will NOT be inserting/amending an entry in the event calendar for this tournament.<br />");
	Response.Write("<hr /><br />")
}


// If we need to update/create event from this ...

if (m_event == checkboxon) {

	// Can only do an event if this is a one-day tournament or
	// if it has a specified Finals day
	// Finals Day will take precedence

	var oneday = (tourObj.tourstart == tourObj.tourend) ? true : false;
	var separatefinals = (tourObj.tourfinalsday) ? true : false;
	var evstart = new String("").toString();
	var evend = new String("").toString();

	if (oneday || separatefinals)  // either one-day tournament or Finals Day
	{
		if (separatefinals) {
			// Finals day so use that start/end time
			evstart = new String(tourObj.finalsstarttime);
			evend = new String(tourObj.finalsendtime);
		} else {
			// One-day tournamwnt with no Finals Day
			evstart = new String(tourObj.tourstarttime);
			evend = new String(tourObj.tourendtime);
		}
	}

	//  Retrieve event detail relating to this tournament if it exists
	newEventID = null;
	eventObj = getEventForTournament(m_id);
	if (! eventObj) {
		if (oneday || separatefinals) {
			newEventID = newEvent();  // This will create a new row in the table ...
			eventid = newEventID;
			eventObj = getEventByID(eventid);
			// Update this new event with defaults from the tournament
			dummy = new String(m_start).toString();
			dateArray = m_start.split("/");  // the year should now be in dateArray(2)
			dummy = new String(dateArray[2]).toString();
			eventObj.eventyear = parseInt(dummy);
			eventObj.eventtype = new String("TOURNAMENT").toString();
			eventObj.eventdate = evstart;
			eventObj.eventtime = tourObj.tourstarttime;
			eventObj.eventnote = tourObj.tourtitle;
			eventObj.eventreport = null;
			eventObj.enddate = evend;
			eventObj.fixturelink = null;
			eventObj.tourlink = "tournaments.html";
			eventObj.holidaylink = "holidaycamps.html";
			eventObj.advert = null;
			updateSQL2 = setEvent(eventObj, true);
		}
		if (debugging) {
			Response.Write("<br />SQL used for setEvent is <br /<br />"+updateSQL2+"<br />");
		}
	}
	else
	{
		eventid = eventObj.eventid;
	}

	if (debugging) {
		if (newEventID) {
			Response.Write("Couldn&apos;t find existing event for this tournament - so have created a new one with an ID of "+newEventID+"<br />");
			// SQLresult = deleteEvent(eventid);   //  ... so we need to clear up by deleting it now.
		} 
		Response.Write(printEvent(eventObj));
		Response.Write("<br /><br />");
		Response.Write("SQL for the clearup is <br /><br />");
		Response.Write(SQLresult);
		Response.Write("<br /><br /><hr /><br />");
	}

	// Now edit the event data using existing admin screen

	editURL = "http://hamptontennis.org.uk/admin/#/events/"+eventid;
	if (debugging) {
		Response.Write("<br /><br />");
		Response.Write("URL to use for edit = ["+editURL+"]<br />");		
	} else {
		Response.Redirect(editURL);
	}
	

}  // end of section for checkboxon for m_event


// Send me a confirmatory email about what has been done.

if (debugging) {
	Response.Write("This is where the confirmatory email will be sent from<br /><br />");
}

// Processing finished - now return to the list of tournaments

if (debugging)
	Response.End();

Response.Redirect("http://hamptontennis.org.uk/admin/#/tournaments");

%>



