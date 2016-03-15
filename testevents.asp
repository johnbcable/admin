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
<!doctype html>
<html class="no-js" lang="en">
  <head>
    <meta charset="utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <title>Hampton-in-Arden Tennis Club - Adding New Event to Calendar</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta http-equiv="Pragma" content="no-cache">
    <meta http-equiv="Expires" content="-1">
    <meta name="Description" lang="en" content="Hampton in Arden Tennis Club web site">
    <!-- Schema.org markup for Google+ -->
    <meta itemprop="name" content="Hampton in Arden Tennis Club web site">
    <meta itemprop="description" content="Hampton in Arden tennis club is a family-friendly tennis club in the green belt area around Solihull ">
    <meta itemprop="image" content="http://hamptontennis.org.uk/img/logos/logo.gif">
    <!-- Twitter Card data -->
    <meta name="twitter:card" content="summary_large_image">
    <meta name="twitter:site" content="@hamptontennis">
    <meta name="twitter:title" content="Hampton in Arden Tennis Club web site">
    <meta name="twitter:description" content="Your family-friendly tennis club in Solihull">
    <meta name="twitter:creator" content="@author_handle">
    <meta name="twitter:image:src" content="http://hamptontennis.org.uk/img/logos/logo.gif">
    <!-- Open Graph meta information -->
    <meta property="og:title" content="Hampton in Arden Tennis Club">
    <meta property="og:type" content="website">
    <meta property="og:url" content="http://hamptontennis.org.uk/fullindex.html">
    <meta property="og:image" content="http://hamptontennis.org.uk/img/logos/logo.gif">
    <meta property="og:description" content="Your family-friendly tennis club in Solihull">
    <!-- Apple touch icon links -->
    <link rel="icon" sizes="192x192" href="/img/logos/icon192.png">
    <link rel="apple-touch-icon-precomposed" sizes="180x180" href="/img/logos/icon180.png">
    <link rel="apple-touch-icon-precomposed" sizes="152x152" href="/img/logos/icon152.png">
    <link rel="apple-touch-icon-precomposed" sizes="144x144" href="/img/logos/icon144.png">
    <link rel="apple-touch-icon-precomposed" sizes="120x120" href="/img/logos/icon120.png">
    <link rel="apple-touch-icon-precomposed" sizes="114x114" href="/img/logos/icon114.png">
    <link rel="apple-touch-icon-precomposed" sizes="76x76" href="/img/logos/icon76.png">
    <link rel="apple-touch-icon-precomposed" sizes="72x72" href="/img/logos/icon72.png">
    <link rel="apple-touch-icon-precomposed" href="/img/logos/apple-touch-icon-precomposed.png">
    <!-- Favicon link -->
    <link rel="shortcut icon" href="/favicon.ico">
    <!-- IE tile icon links -->
    <meta name="msapplication-TileColor" content="#FFFFFF">
    <meta name="msapplication-TileImage" content="/img/logos/icon144.png">
    <meta name="msapplication-square310x310logo" content="/img/logos/icon310.png">
    <meta name="msapplication-wide310x150logo" content="/img/logos/tile-wide.png">
    <meta name="msapplication-square150x150logo" content="/img/logos/icon150.png">
    <meta name="msapplication-square70x70logo" content="/img/logos/icon70.png">
    <!-- CSS links -->
    <link rel="stylesheet" href="bower_components/foundation/css/normalize.css" />
    <!-- <link rel="stylesheet" href="css/base.css" />  -->
    <link rel="stylesheet" href="/css/main.css" />
    <script src="/bower_components/modernizr/modernizr.js"></script>
    <style type="text/css">
    li.current a {
      background-color: white;
      font-weight: bold;
    }
    </style>
  </head>
<body>

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
eventObj.eventnote = "Updated title for new event";
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


</body>
</html>
<%
%>


