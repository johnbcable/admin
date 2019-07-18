<%@language="JScript" CODEPAGE="65001"%>
<%
Response.AddHeader("Access-Control-Allow-Origin", "*");
Response.AddHeader("Cache-Control", "no-cache,no-store,must-revalidate");
Response.AddHeader("Pragma", "no-cache");
Response.AddHeader("Expires", 0);
%>
<!--#include file="json2.js.asp" -->
<!--#include file="functions.asp" -->
<!--#include file="strings.asp" -->
<!--#include file="coachinglessonobj.asp" -->
<!--#include file="emailfuncs.asp" -->
<!--#include file="datefuncs.asp" -->
<%
var strtime, strdate, title2;
var clubname = new String("Hampton-In-Arden Sports Club");
var pagetitle = new String("Generate Events From Coaching Lesson Schedule");
// Now for any variables local to this page
var RS, Conn, SQL1, SQL2;
var dbconnect=Application("hamptonsportsdb"); 
var SQLstart, SQLmiddle, SQLend;
var resultObj = new Object();
var uniqueref;
var today = new Date();
var lessondetail = new Object(); 
var start_time, end_time, lesson_text, group_code;
var start_date, end_date, break_start, break_end;
var date_start, date_end;
var schedule_role, lesson_day, lesson_id;
var lessons = new Array();
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

SQL1 = new String("SELECT start_date, end_date, break_start, break_end FROM coaching_schedules WHERE schedule_role = 'NEXT'").toString();

RS=Conn.Execute(SQL1);// Only one row or absent so no loop
if (! RS.EOF) {
	start_date = new Date(RS("start_date"));
	end_date = new Date(RS("end_date"));
	break_start = new Date(RS("break_start"));
	break_end = new Date(RS("break_end"));
}
RS.Close();

if (debugging) {
	Response.Write("<h4>After schedule retrieval</h4>");
	Response.Write("<table>");
	Response.Write("<tr><td>Start Date</td><td>End Date</td><td>Start of break</td><td>End of Break</td></tr>");
	Response.Write("<tr><td>"+ddmmyyyy(start_date)+"</td><td>"+ddmmyyyy(end_date)+"</td><td>"+ddmmyyyy(break_start)+"</td><td>"+ddmmyyyy(break_end)+"</td></tr>");
	Response.Write("</table>");
}

// OK start loop from start_date 
mydate = new Date(start_date);
start_break = new Date(break_start);
end_break = new Date(break_end);
real_end = new Date(end_date);

start_break = start_break.setDate(start_break.getDate() - 1);
end_break = end_break.setDate(end_break.getDate() + 1);
real_end = real_end.setDate(real_end.getDate() + 1);

while (mydate < real_end ) {
// for (var i=0; i < 60; i++) {

	curdate = new Date(mydate);

	if (debugging) {
		Response.Write("<h5>"+ddmmyyyy(curdate)+"</h5>");
	}
	// if this date is in the break period, ignore it

	if (curdate > start_break && curdate < end_break) {
		if (debugging) {
			Response.Write(curdate+" is in the break period ("+break_start+" to "+break_end+")");
		}
	}
	else 
	{
		// We are not in the break period so 
		// 1.  Get the day of curdate
		// 2.  Retrieve all lessons for that day
		// 3.  Set up SQL insertions based on these

		// 1.  Get the day of curdate

		curday = new String(weekdays[curdate.getDay()]).toString();
		theday = (curday.substring(0,3)).toUpperCase();

		if (! (theday == "SUN") ) {

			// exclude Sundays

			// 2.  Retrieve all lessons for that day

			lessonSQL = new String("SELECT Format(start_time,'HH:mm:ss') AS start_time, Format(end_time,'HH:mm:ss') AS end_time, lesson_text FROM coaching_lessons WHERE lesson_day = '"+theday+"' AND start_time NOT LIKE '00:00:00' AND schedule_role = 'NEXT' ORDER BY lesson_id, group_display_order, start_time, end_time").toString();

			RS = Conn.Execute(lessonSQL);
			while (! RS.EOF) {

				m_date = ddmmyyyy(curdate);
				m_year = curdate.getFullYear();
				m_start = new String(RS("start_time")).toString();
				m_end = new String(RS("end_time")).toString();
				m_text = new String(RS("lesson_text")).toString();


				// 3.  Set up SQL insertion based on this data

				eventSQL = new String("INSERT INTO "+destinationtable+"([eventdate],[eventtime],[eventyear],[eventtype],[eventnote],[enddate],[endtime]) VALUES ('"+m_date+"','"+m_start+"',"+m_year+",'COACHING','"+m_text+"','"+m_date+"','"+m_end+"')").toString();
				if (debugging) {
					Response.Write(eventSQL+"<br />");
				}
				if (updating) {
					RS2 = Conn.Execute(eventSQL);
				}

				RS.MoveNext();
			}
			RS.Close();

		}


	}

	mydate.setDate(mydate.getDate() + 1);  // Move on to the next day

}


Response.End();


%>


