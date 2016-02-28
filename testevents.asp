<%@language="JScript"%>
<!--#include file="datefuncs.asp" -->
<!--#include file="eventfuncs.asp" -->
<%
// First variables that need to be set for each page
var strtime, strdate;
var clubname = new String("Hampton-In-Arden Sports Club");
var pagetitle = new String("Event Testing Page");
// Now for any variables local to this page
var dbconnect=Application("hamptonsportsdb");
var ConnObj, RstObj, StnObj;
var ConnObj2, RstObj2, StnObj2;
var RS, RS2, RS3;
var SQL1, SQL2;
var d, thisyear;
var i, kount, dummy, dummy1;
var eventids = new Array (786,787,788);
var curteam, curseason;
var teamcaptain, teamname;
var venuetext, teamnote;
var thefixturedate, strfixturedate;
var isOdd = true;
var stripeText = new String("").toString();
var eventObj = new Object();
// Set up default greeting strings
strdate = "";
strtime = "";
// End of page start up coding
displaydate = "";
// var debugging=current_debug_status();
var debugging = false;


// End of page start up coding
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>Hampton-In-Arden Tennis Club Home Page</title>
	<link rel="stylesheet" media="screen" type="text/css" href="layout.css" />
	<link rel="stylesheet" media="screen" type="text/css" href="colours.css" />
	<link rel="stylesheet" media="screen" type="text/css" href="typography.css" /> 
	<link rel="stylesheet" media="print" type="text/css" href="print3col.css" /> 
	<!-- Comment out the next style sheet if runniung in production -->
<%
if (debugging)
{
%>
<link rel="stylesheet" media="screen" type="text/css" href="borders.css" />
<%
}
%>
</head>
<body>
<!--   1.  Branding   -->
<div id="branding">
	<a href="index.asp" id="homelink"><img id="clublogo" src="images/logo.gif" alt="Hampton-In-Arden Sports Club logo" /></a>
	<h1>Hampton-In-Arden Sports Club</h1>
	<h2>Tennis Section</h2>

<!--   2.  Navigation   -->
	<div id="nav_main">
		<ul id="topmenu">
			<li id="nav_ourclub"><a href="aboutus.asp">Our Club</a></li>
			<li id="nav_coaching"><a href="juniors/coaching.html">Coaching</a></li>
			<li id="nav_playing"><a href="playing.asp">Playing</a></li>
			<li id="nav_links"><a href="juniors/index.html">Juniors</a></li>
			<li id="nav_contact"><a href="juniors/contact.html">Contact</a></li>
			<li id="nav_members"><a href="members.asp">Members</a></li>
		</ul>
		<p id="today">
			<%= displaydate %>&nbsp;<%= strtime %>
		</p>
	</div>
	
</div>

<div id="wrapper">

<!--   3. Content    -->
	<div id="content">
			<h1>Administration:<b>&nbsp;Testing Event Functions</b></h1>
<%

ConnObj = Server.CreateObject("ADODB.Connection");
RS = Server.CreateObject("ADODB.RecordSet");
ConnObj.Open(dbconnect);
SQL1 = "SELECT MAX(eventid) AS lastevent FROM events";

Response.Write("<h1>Firstly, fetching event data from last created event</h1>");
var latestevent = -1;
RS = ConnObj.Execute(SQL1);
while (! RS.EOF) {
	Response.Write("<p>Inside RS loop from "+SQL1+"<br /></p>");
	Response.Write("<p>RS(0) = "+RS(0)+"<br /></p>");
	latestevent = new Number(RS(0));
	Response.Write("<p>latestevent after assignment inside loop = "+latestevent+"<br /></p>");
	RS.MoveNext();
}
RS.Close();
Response.Write("<p>Last created event = ["+latestevent+"]<br /></p>");

eventObj = getEventByID(latestevent);
Response.Write("<h2>Latest Event: "+eventObj.eventnote+" - by ID ("+eventObj.eventid+")</h2>")
Response.Write(printEvent(eventObj));
Response.Write("<br /><hr /><br />");

Response.Write("<h1>Secondly, fetching event data from list of events </h1>");
for (i=0; i<3; i++) {
	eventObj = getEventByID(eventids[i]);
	Response.Write("<h2>Event: "+eventObj.eventnote+" - by ID ("+eventObj.eventid+")</h2>")
	Response.Write(printEvent(eventObj));
	Response.Write("<br />");
}	
Response.Write("<hr /><br />");
Response.Write("Create a new event ...<br />")
var newEventID = newEvent();
Response.Write("Created a new event with the ID of "+newEventID+"<br />");
Response.Write("Now get the event object back from this ID ... <br />");
eventObj = getEventByID(newEventID);
Response.Write("<h2>Tournament: "+eventObj.eventnote+" - by ID ("+eventObj.eventid+")</h2>");
Response.Write(printEvent(eventObj));
Response.Write("<br />");
Response.Write("Update the event object with new title etc and save back to DB <br />");
eventObj.eventnote = "Update title for new event";
eventObj.eventdate = "01/09/2014";
eventObj.eventtype = "ADULT";
eventObj.eventid = newEventID;
Response.Write("<br />Event object after in-memory updates:<br /><br />");
Response.Write("<h2>Event: "+eventObj.eventnote+" - by ID ("+eventObj.eventid+") is now</h2>");
Response.Write(printEvent(eventObj));
var updateSQL = setEvent(eventObj, false);
Response.Write("The update SQL was:<br /><br />"+updateSQL+"<br /><br /><hr />");



Response.Write("Now get the updated event object back from this ID ... "+newEventID+"<br />");
neweventObj = new Object();
neweventObj = getEventByID(newEventID);
Response.Write("<br />Event object after retrieval from DB:<br /><br />");
Response.Write("<h2>Event: "+neweventObj.eventnote+" - by ID ("+neweventObj.eventid+")</h2>");
Response.Write(printEvent(neweventObj));
Response.Write("<br />");

// Now clear up - delete this event

Response.Write("Now delete event "+newEventID+" to clear up after ourselves<br />");
var deleteSQL = deleteEvent(newEventID, false);
Response.Write("The deletion SQL was:<br /><br />"+deleteSQL+"<br /><br /><hr />");

%>		
	</div>
	
<!--     4.      Supplementary navigation    -->
	<div id="leftcolumn">
	</div>

<!--    5.   Supplementary content     -->	
	<div id="rightcolumn">

	</div>
</div>

<!--     6.    Site info     -->

</body>
</html>
<%
%>


