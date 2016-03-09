<%@language="JScript" CODEPAGE="65001"%>
<%
Response.AddHeader("Access-Control-Allow-Origin", "http://www.hamptontennis.org.uk");
%>
<!--#include file="functions.asp" -->
<!--#include file="strings.asp" -->
<!--#include file="emailfuncs.asp" -->
<!--#include file="eventfuncs.asp" -->
<%

//  Need to rework to use utility functions newEvent, getEventById and setEvent
//  from eventfuncs.asp

var strtime, strdate, title2;
var clubname = new String("Hampton-In-Arden Sports Club");
var pagetitle = new String("Updating Details of Event");
// Now for any variables local to this page
var m_eventdate, m_eventtime, m_eventnote, m_eventtype, m_eventid, m_eventreport;
var vdate, m_eventyear;   // m_eventyear is always calculated from m_eventdate
var m_enddate, m_endtime;
var eventObj = new Object();
var m_debug;
var ConnObj;
var RS,RS2,RS3;
var SQLStmt, SQL2, SQL3, updateSQL;
var kount;
var memberknt;
var dbconnect=Application("hamptonsportsdb");
var debugging=true;
var updating=true;
// Set up default greeting strings
strdate = datestring();
strtime = timestring();
// End of page start up coding
// Initialise update variables
m_eventid = Trim(new String(Request.QueryString("eventid")));
m_eventdate = Trim(new String(Request.QueryString("eventdate")));
m_eventtime = Trim(new String(Request.QueryString("eventtime")));
m_eventnote = Trim(new String(Request.QueryString("eventnote")));
m_eventtype = Trim(new String(Request.QueryString("eventtype")));
m_eventreport = Trim(new String(Request.QueryString("eventreport")));
m_enddate = Trim(new String(Request.QueryString("enddate")));
m_endtime = Trim(new String(Request.QueryString("endtime")));
m_debug = Trim(new String(Request.QueryString("debug")));

// Set debugging dependent on querystring override
if (m_debug == "y" || m_debug == "Y") {
	debugging = true;
} 

// Flag if this is a new member insertion 
newone = (m_eventid == "-1");
// reset if null
if (m_eventdate=="undefined" || m_eventdate == "null" || m_eventdate == "")
{
	today = new Date();
	m_eventdate = new String(ddmmyyyy(today)).toString();
}
if (m_eventtime=="undefined" || m_eventtime == "null" || m_eventtime == "")
	m_eventtime = new String("00:00:00").toString();
else
	m_eventtime = new String(m_eventtime.substr(0,2)+":"+m_eventtime.substr(2)+":00").toString();
if (m_eventnote=="undefined" || m_eventnote == "null" || m_eventnote == "")
	m_eventnote = new String("No event title supplied").toString();
if (m_eventtype=="undefined" || m_eventtype == "null" || m_eventtype == "")
	m_eventtype = new String("EVENT").toString();
if (m_eventreport=="undefined" || m_eventreport == "null" || m_eventreport == "")
	m_eventreport = new String("").toString();
if (m_enddate=="undefined" || m_enddate == "null" || m_enddate == "")
	m_enddate = new String(m_eventdate).toString();
if (m_endtime=="undefined" || m_endtime == "null" || m_endtime == "")
	m_endtime = new String("00:00:00").toString();
else
	m_endtime = new String(m_endtime.substr(0,2)+":"+m_endtime.substr(2)+":00").toString();

// What sort of mtime value do we have
if (m_eventtime.length < 8) {
	// not in 8-character form so we need to split and reformat
	timearr = m_eventtime.split(":");
	m_eventtime = new String(Lpad(timearr[0],2,"0")+":"+Lpad(timearr[1],2,"0"));
}

// Calculate event year from event date
// What sort of inout date do we have
if (m_eventdate.length > 8) {
	// HTML5 yyyy-mm-dd format
	m_eventyear= new Number(m_eventdate.substr(0,4));

} else {
	// assume in dd/mm/yyyy format
	m_eventyear = new Number(m_eventdate.substr(6,4))
}
if (isNaN(m_eventyear)) {
	// we have an invalid date
}

// Do DB update
ConnObj=Server.CreateObject("ADODB.Connection");
RS=Server.CreateObject("ADODB.Recordset");
ConnObj.Open(dbconnect);
if (debugging) {
	Response.Write("m_eventid from initial submission is now "+m_eventid);
}
if (newone)
{
	// Create the new event
	if (debugging) {
		Response.Write("New event - need to generate new DB entry. m_eventid is now "+m_eventid);
	}
	m_eventid = newEvent();
}
if (debugging) {
	Response.Write("m_eventid is now "+m_eventid);
}
// Retrieve event by its ID
eventObj = getEventByID(m_eventid);

// Update local event object
eventObj.eventdate = new String(m_eventdate).toString();
eventObj.eventtime = new String(m_eventtime).toString();
eventObj.eventyear = new String(m_eventyear).toString();
eventObj.eventtype = new String(m_eventtype).toString();
eventObj.eventnote = new String(m_eventnote).toString();
eventObj.eventid = new Number(m_eventid).valueOf();
eventObj.eventreport = new String(m_eventreport).toString();
eventObj.enddate = new String(m_enddate).toString();
eventObj.endtime = new String(m_endtime).toString();
eventObj.fixturelink = new String("").toString();
eventObj.tourlink = new String("").toString();
eventObj.holidaylink = new String("").toString();
eventObj.advert = new String("").toString();

if (debugging)
{
	printEvent(eventObj);
}

// Update database from event object
if (updating) {
	updateSQL = setEvent(eventObj, false);
}

if (debugging)
{
	Response.Write("current_debug_status()<br />");
	Response.Write("m_eventid = "+m_eventid+"<br />");
	Response.Write("m_eventdate = "+m_eventdate+"<br />");
	Response.Write("m_eventtime = "+m_eventtime+"<br />");
	Response.Write("m_enddate = "+m_enddate+"<br />");
	Response.Write("m_endtime = "+m_endtime+"<br />");
	Response.Write("m_eventnote = "+m_eventnote+"<br />");
	Response.Write("m_eventreport = "+m_eventreport+"<br />");
	Response.Write("m_eventyear = "+m_eventyear+"<br />");
	Response.Write("<br />updateSQL = ["+updateSQL+"]<br />");
}

RS=null;
ConnObj.Close();
ConnObj=null;
if ((! current_debug_status()) && updating )
	Response.Redirect("/admin/#{/events");
%>

