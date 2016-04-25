
<%@language="JScript" CODEPAGE="65001"%>
<!--#include file="json2.js.asp" -->
<!--#include file="functions.asp" -->
<!--#include file="strings.asp" -->
<!--#include file="socialtennisobj.asp" -->
<!--#include file="emailfuncs.asp" -->
<!--#include file="datefuncs.asp" -->
<%

// Set up cache control on this page
Response.AddHeader("Cache-Control", "no-cache,no-store,must-revalidate");
Response.AddHeader("Pragma", "no-cache");
Response.AddHeader("Expires", 0);

var strtime, strdate, title2;
var clubname = new String("Hampton-In-Arden Sports Club");
var pagetitle = new String("Generate Events From Social Tennis Schedule");
// Now for any variables local to this page
var RS, Conn, SQL1, SQL2;
var dbconnect=Application("hamptonsportsdb"); 
var SQLstart, SQLmiddle, SQLend;
var resultObj = new Object();
var uniqueref;
var today = new Date();
var socialsessiondetail = new Object(); 
var start_time, end_time;
var socialsessions = new Array();
var weekdays = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];
var mydebug;
var destinationtable = "events";   // Production = events, testing = newevents
var m_note = new String("");
var m_date, m_start, m_end, m_year;
var inserting = false;

var debugging = false;  // Production = false
var updating = true;   // Production = true    Add records to the database

// Process QueryString for run options
mydebug = Request.QueryString("debug");
if (mydebug == "" || mydebug == "null" || mydebug == "undefined")
{
	mydebug = new String("N").toString();
}
if (mydebug == "Y") {
	destinationtable = new String("newevents").toString();
	debugging = true;
} else {
	destinationtable = new String("events").toString();
	debugging = false;
}

// debugging = true;

// Set up DB connections and unchanging bits of SQL
Conn = Server.CreateObject("ADODB.Connection");
RS = Server.CreateObject("ADODB.RecordSet");
RS2 = Server.CreateObject("ADODB.RecordSet");
Conn.Open(dbconnect);

// Firstly lets get the dates for the next coaching schedule.
// Initial implementation - hard-code start and end date
// new Date(year, month, day, hours, minutes, seconds, milliseconds)

start_date = new Date("March 5, 2016 01:15:00");
end_date = new Date("December 31, 2016 23:15:00");

if (debugging) {
	Response.Write("<h4>Period for Social Tennis Events</h4>");
	Response.Write("<table>");
	Response.Write("<tr><td>Start Date</td><td>End Date</td><td>Table</td></tr>");
	Response.Write("<tr><td>"+ddmmyyyy(start_date)+"</td><td>"+ddmmyyyy(end_date)+"</td><td>"+destinationtable+"</td></tr>");
	Response.Write("</table>");
}

// OK start loop from start_date 
mydate = start_date;
real_end = end_date;
real_end = real_end.setDate(real_end.getDate() + 1);

while (mydate < real_end ) 
{

	// Social tennis runs on Tues and Fri evenings, Saturday afternoons and Sunday mornings.

	inserting = false;   // default to not inserting - could be wrong day
	curdate = new Date(mydate);

	if (debugging) {
		Response.Write("<h5>"+ddmmyyyy(curdate)+"</h5>");
	}

	// 1.  Get the day of curdate

	curday = new String(weekdays[curdate.getDay()]).toString();
	theday = (curday.substring(0,3)).toUpperCase();

	switch (theday) {
		case "SUN":
			start_time = "10:00:00";
			end_time = "12:00:00";
			m_note = "Sunday morning social tennis session";
			inserting = true;
			break;
		case "TUE":
			start_time = "20:00:00";
			end_time = "22:00:00";
			m_note = "Tuesday evening social tennis session"
			inserting = true;
			break;
		case "FRI":
			start_time = "20:00:00";
			end_time = "22:00:00";
			m_note = "Friday evening social tennis session"
			inserting = true;
			break;
		case "SAT":
			start_time = "15:00:00";
			end_time = "17:00:00";
			m_note = "Saturday afternoon family social tennis session"
			inserting = true;
			break;
	}

	if ( inserting ) {

		// 3.  Set up SQL insertion based on this data

		m_date = ddmmyyyy(curdate);
		m_start = new String(start_time).toString();
		m_end = new String(end_time).toString();
		m_year = curdate.getFullYear();

		eventSQL = new String("INSERT INTO "+destinationtable+"([eventdate],[eventtime],[eventyear],[eventtype],[eventnote],[enddate],[endtime]) VALUES ('"+m_date+"','"+m_start+"',"+m_year+",'SOCIALTENNIS','"+m_note+"','"+m_date+"','"+m_end+"')").toString();
		if (debugging) {
			Response.Write(eventSQL+"<br />");
		}
		if (updating) {
			RS2 = Conn.Execute(eventSQL);
		}

	}   /// end of inserting check

	mydate.setDate(mydate.getDate() + 1);  // Move on to the next day

}

Response.End();

%>
