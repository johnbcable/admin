<%@language="JScript" CODEPAGE="65001"%>
<!--#include file="functions.asp" -->
<!--#include file="strings.asp" -->
<!--#include file="emailfuncs.asp" -->
<!--#include file="datefuncs.asp" -->
<%

// Set up cache control on this page
Response.AddHeader("Cache-Control", "no-cache,no-store,must-revalidate");
Response.AddHeader("Pragma", "no-cache");
Response.AddHeader("Expires", 0);

var strtime, strdate, title2;
var clubname = new String("Hampton-In-Arden Sports Club");
var pagetitle = new String("Updating Coaching Dates and Fees");
// Now for any variables local to this page
var RS, Conn, SQL1, SQL2;
var dbconnect=Application("hamptonsportsdb"); 
var SQLstart, SQLmiddle, SQLend;
var resultObj = new Object();
var uniqueref;
var coachingobject = new Object();
var today = new Date();

var debugging = false;  // Production = false

// defaults for variables
coachingobject.startdate = new String(ddmmyyyy(today));
coachingobject.enddate = new String(ddmmyyyy(today));
coachingobject.excludestart = new String(ddmmyyyy(today));
coachingobject.excludeend = new String(ddmmyyyy(today));
coachingobject.fees60 = "0.00";
coachingobject.fees45 = "0.00";
coachingobject.fees30 = "0.00";

// Retrieve POST'ed data

coachingobject.startdate = Trim(new String(Request.Form("start_date")));
if (coachingobject.startdate == "" || coachingobject.startdate =="null" || coachingobject.startdate == "undefined")
{
	coachingobject.startdate = new String("").toString();
} 

coachingobject.enddate = Trim(new String(Request.Form("end_date")));
if (coachingobject.enddate == "" || coachingobject.enddate =="null" || coachingobject.enddate == "undefined")
{
	coachingobject.enddate = new String("").toString();
} 

coachingobject.excludestart = Trim(new String(Request.Form("break_start")));
if (coachingobject.excludestart == "" || coachingobject.excludestart =="null" || coachingobject.excludestart == "undefined")
{
	coachingobject.excludestart = new String("").toString();
} 

coachingobject.excludeend = Trim(new String(Request.Form("break_end")));
if (coachingobject.excludeend == "" || coachingobject.excludeend =="null" || coachingobject.excludeend == "undefined")
{
	coachingobject.excludeend = new String("").toString();
} 

coachingobject.fees60 = Trim(new String(Request.Form("fees_hour")));
if (coachingobject.fees60 == "" || coachingobject.fees60 =="null" || coachingobject.fees60 == "undefined")
{
	coachingobject.fees60 = new String("0.00").toString();
} 

coachingobject.fees45 = Trim(new String(Request.Form("fees_45mins")));
if (coachingobject.fees45 == "" || coachingobject.fees45 =="null" || coachingobject.fees45 == "undefined")
{
	coachingobject.fees45 = new String("0.00").toString();
} 

coachingobject.fees30 = Trim(new String(Request.Form("fees_30mins")));
if (coachingobject.fees30 == "" || coachingobject.fees30 =="null" || coachingobject.fees30 == "undefined")
{
	coachingobject.fees30 = new String("0.00").toString();
} 


// Update coaching dates details from POST'ed data

Conn = Server.CreateObject("ADODB.Connection");
RS = Server.CreateObject("ADODB.RecordSet");
Conn.Open(dbconnect);
SQLstart = new String("UPDATE coaching_dates ")
SQLend = new String(" WHERE schedule_role = 'NEXT'").toString();
SQLmiddle = new String("SET ").toString();
SQLmiddle += " fees_hour='"+coachingobject.fees60+"',";
SQLmiddle += " fees_45mins='"+coachingobject.fees45+"',";
SQLmiddle += " fees_30mins='"+coachingobject.fees30+"',";

// Now do date fields. If null dont insert them as part of the update clause
//  Access doesnt like setting date fields to ''
// N.B.  If dob is the last column, dont put trailing comma on
if (! (coachingobject.startdate == "")) {
	SQLmiddle += " start_date='"+coachingobject.startdate+"', ";
} else {
	SQLmiddle += " start_date=null, ";
}
if (! (coachingobject.enddate == "")) {
	SQLmiddle += " end_date='"+coachingobject.enddate+"', ";
} else {
	SQLmiddle += " end_date=null, ";
}
if (! (coachingobject.excludestart == "")) {
	SQLmiddle += " break_start='"+coachingobject.excludestart+"', ";
} else {
	SQLmiddle += " break_start=null, ";
}
if (! (coachingobject.excludeend == "")) {
	SQLmiddle += " break_end='"+coachingobject.excludeend+"' ";
} else {
	SQLmiddle += " break_end=null ";
}

// Default values into result object
resultObj.result = true;
resultObj.errno = 0;
resultObj.description = new String("").toString();

SQL1 = new String(SQLstart+SQLmiddle+SQLend).toString();
resultObj.sql = new String(SQL1).toString();

if (! debugging) {
	try {
		RS = Conn.Execute(SQL1);
	}
	catch(e) {
		resultObj.result = false;
		resultObj.errno = (e.number & 0xFFFF);
		resultObj.description += e.description;
		resultObj.sql = new String(SQLstart+SQLmiddle+SQLend).toString();
	}
}

if (debugging) {
	var theresult = resultObj.result ? "TRUE" : "FALSE";
	Response.Write("Information from the DB update:<br />");
	Response.Write("Result: "+ theresult+"<br />");
	Response.Write("errno: "+resultObj.errno+"<br />");
	Response.Write("Description: "+resultObj.description+"<br />");
	Response.Write("SQL used: <br />"+resultObj.sql+"<br /><br />");

	Response.Write("<h3>Values POST-ed</h3>");
	Response.Write("<table>");
	Response.Write("<tr><td>Start Date</td><td>"+coachingobject.startdate+"</td></tr>");
	Response.Write("<tr><td>End Date</td><td>"+coachingobject.enddate+"</td></tr>");
	Response.Write("<tr><td>Start of period to exclude</td><td>"+coachingobject.excludestart+"</td></tr>");
	Response.Write("<tr><td>End of period being excluded</td><td>"+coachingobject.excludeend+"</td></tr>");
	Response.Write("<tr><td>Fee for 1-hour lessons</td><td>"+coachingobject.fees60+"</td></tr>");
	Response.Write("<tr><td>Fee for 45-minute lessons</td><td>"+coachingobject.fees45+"</td></tr>");
	Response.Write("<tr><td>Fee for 30-minute lessons</td><td>"+coachingobject.fees30+"</td></tr>");
	Response.Write("</table>");

	Response.End();

}

// On completion, redirect to coachingmondays.html

if (! debugging) {
	Response.Redirect("./coachingmondays.html");
}

Response.End();


%>


