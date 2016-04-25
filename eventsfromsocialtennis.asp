
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

// Set up DB connections and unchanging bits of SQL
Conn = Server.CreateObject("ADODB.Connection");
RS = Server.CreateObject("ADODB.RecordSet");
RS2 = Server.CreateObject("ADODB.RecordSet");
Conn.Open(dbconnect);

// Firstly lets get the dates for the next coaching schedule.
// Initial implementation - hard-code start and end date
/*
SQL1 = new String("SELECT start_date, end_date, break_start, break_end FROM coaching_schedules WHERE schedule_role = 'NEXT'").toString();

RS=Conn.Execute(SQL1);// Only one row or absent so no loop
if (! RS.EOF) {
	start_date = new Date(RS("start_date"));
	end_date = new Date(RS("end_date"));
	break_start = new Date(RS("break_start"));
	break_end = new Date(RS("break_end"));
}
RS.Close();
*/

start_date = new Date("10/03/2016");
end_date = new Date("31/12/2016");

if (debugging) {
	Response.Write("<h4>After schedule retrieval</h4>");
	Response.Write("<table>");
	Response.Write("<tr><td>Start Date</td><td>End Date</td></tr>");
	Response.Write("<tr><td>"+ddmmyyyy(start_date)+"</td><td>"+ddmmyyyy(end_date)+"</td></tr>");
	Response.Write("</table>");
}

// OK start loop from start_date 
mydate = new Date(start_date);
real_end = new Date(end_date);
real_end = real_end.setDate(real_end.getDate() + 1);

while (mydate < real_end ) 
{

	// Social tennis runs on Tues and Fri evenings, Saturday afternoons and Sunday morning.

	curdate = new Date(mydate);

	if (debugging) {
		Response.Write("<h5>"+ddmmyyyy(curdate)+"</h5>");
	}
		// 1.  Get the day of curdate
		// 2.  Set up SQL insertions if Tues, Fri, Sat or Sungg

		// 1.  Get the day of curdate

		curday = new String(weekdays[curdate.getDay()]).toString();
		theday = (curday.substring(0,3)).toUpperCase();

		switch (theday) {
			case "SUN":
				start_time = "10:00:00";
				end_time = "12:00:00";
				break;
			case "TUE":
				start_time = "20:00:00";
				end_time = "22:00:00";
				break;
			case "FRI":
				start_time = "20:00:00";
				end_time = "22:00:00";
				break;
			case "SAT":
				start_time = "15:00:00";
				end_time = "17:00:00";
				break;
		}

		// 3.  Set up SQL insertion based on this data

		m_date = ddmmyyyy(curdate);
		m_start = new String(start_time).toString();
		m_end = new String(end_time).toString();
		m_year = curdate.getFullYear();

		eventSQL = new String("INSERT INTO "+destinationtable+"([eventdate],[eventtime],[eventyear],[eventtype],[eventnote],[enddate],[endtime]) VALUES ('"+m_date+"','"+m_start+"',"+m_year+",'SOCIAL','','"+m_date+"','"+m_end+"')").toString();
		if (debugging) {
			Response.Write(eventSQL+"<br />");
		}
		if (updating) {
			RS2 = Conn.Execute(eventSQL);
		}

	mydate.setDate(mydate.getDate() + 1);  // Move on to the next day

}

Response.End();

%>
