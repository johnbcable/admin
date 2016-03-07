<%@language="JScript"%>
<!--#include file="datefuncs.asp" -->
<!--#include file="tourfuncs.asp" -->
<%
// First variables that need to be set for each page
var strtime, strdate;
var clubname = new String("Hampton-In-Arden Sports Club");
var pagetitle = new String("Tournament Testing Page");
// Now for any variables local to this page
var dbconnect=Application("hamptonsportsdb");
var ConnObj, RstObj, StnObj;
var ConnObj2, RstObj2, StnObj2;
var RS, RS2, RS3;
var d, thisyear;
var i, kount, dummy, dummy1;
var tourids = new Array (16,17,18);
var curteam, curseason;
var teamcaptain, teamname;
var venuetext, teamnote;
var thefixturedate, strfixturedate;
var isOdd = true;
var stripeText = new String("").toString();
var tourObj = new Object();
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
			<h1>Administration:<b>&nbsp;Testing Tournament Functions</b></h1>
<%
Response.Write("<h1>Firstly, fetching older tournamrnt data</h1>");
for (i=0; i<3; i++) {
	tourObj = getTour(tourids[i]);
	Response.Write("<h2>Tournament: "+tourObj.tourtitle+" - by ID ("+tourObj.tournamentid+")</h2>")
	Response.Write(printTour(tourObj));
	Response.Write("<br />");
}	
Response.Write("<hr /><br />");
Response.Write("Create a new tournament ...<br />")
var newTourID = newTour();
Response.Write("Created a new tournament with the ID of "+newTourID+"<br />");
Response.Write("Now get the tournament object back from this ID ... <br />");
tourObj = getTour(newTourID);
Response.Write("<h2>Tournament: "+tourObj.tourtitle+" - by ID ("+tourObj.tournamentid+")</h2>");
Response.Write(printTour(tourObj));
Response.Write("<br />");
Response.Write("Update the tournament object with new title etc and save back to DB <br />");
tourObj.tourtitle = "Update title for new tournament";
tourObj.tourstart = "01/09/2014";
tourObj.tourend = "27/09/2014";
tourObj.tourcategory = "ADULT";
tourObj.tournamentid = newTourID;
Response.Write("<br />Tournament object after in-memory updates:<br /><br />");
Response.Write("<h2>Tournament: "+tourObj.tourtitle+" - by ID ("+tourObj.tournamentid+") is now</h2>");
Response.Write(printTour(tourObj));
var updateSQL = setTour(tourObj, false);
Response.Write("The update SQL was:<br /><br />"+updateSQL+"<br /><br /><hr />");



Response.Write("Now get the updated tournament object back from this ID ... "+newTourID+"<br />");
newtourObj = new Object();
newtourObj = getTour(newTourID);
Response.Write("<br />Tournament object after retrieval from DB:<br /><br />");
Response.Write("<h2>Tournament: "+newtourObj.tourtitle+" - by ID ("+newtourObj.tournamentid+")</h2>");
Response.Write(printTour(newtourObj));
Response.Write("<br />");

// Now clear up - delete this event

Response.Write("Now delete tournament "+newTourID+" to clear up after ourselves<br />");
var deleteSQL = deleteTour(newTourID, false);
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


