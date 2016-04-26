<%@language="JScript" CODEPAGE="65001"%>
<!--#include file="functions.asp" -->
<!--#include file="strings.asp" -->
<!--#include file="emailfuncs.asp" -->
<%
// First variables that need to be set for each page
var strtime, strdate;
var clubname = new String("Hampton-In-Arden Sports Club");
var pagetitle = new String("Fixture List Administration");
// Now for any variables local to this page
var dbconnect=Application("hamptonsportsdb");
var theseason;
var theid, theteam, thevenue;
var thenote, thefixturedate, myDate;
var venuetext, teamnote, teamtext;
var curteam, prevteam;
var ConnObj, RstObj, StnObj;
var ConnObj2, RstObj2, StnObj2;
var RS, RS2, RS3;
var fixtureknt;
var teamnames = new Array(10);
var teamclause, venueclause;
var debugging = false;
var d, dummy, i, kount, thisyear;
var dayparts;
var isOdd = true;
var teamObj = new Object();

var ourscore, theirscore;
var totalthem, totalus, unknown, uswin, uslose, usdraw;

// Set up default greeting strings
strdate = datestring();
strtime = timestring();
// start of page processing

totalthem = 0;
totalus = 0;
uswin = 0;
uslose = 0;
usdraw = 0;
unknown = 0;

if (!signedin())
	Response.Redirect("login.asp");
if (!isFixtureEditor())
	Response.Redirect("fixturelist.asp");
thisyear = new String("").toString();;
d = new Date();
thisyear += d.getYear();
fixtureknt=0;
theseason = new String(Request.Form("season")).toString();
if (theseason == "" || theseason == "null" || theseason == "undefined")
{
	theseason = new String(Request.QueryString("season")).toString();
	if (theseason == "" || theseason == "null" || theseason == "undefined")
		theseason = new String(thisyear).toString();
}
theseason = theseason.toUpperCase();
thevenue = new String(Request.Form("venue")).toString();
if (thevenue == "" || thevenue == "null" || thevenue == "undefined")
{
	thevenue = new String(Request.QueryString("venue")).toString();
	if (thevenue == "" || thevenue == "null" || thevenue == "undefined")
		thevenue = new String("BOTH").toString();
}
thevenue = thevenue.toUpperCase();
if (thevenue == "BOTH")
	venuetext = new String("Home and Away");
if (thevenue == "A")
	venuetext = new String("Away");
if (thevenue == "H")
	venuetext = new String("Home");
theteam = new String(Request.Form("team")).toString();
if (theteam == "" || theteam == "null" || theteam == "undefined")
{
	theteam = new String(Request.QueryString("team")).toString();
	if (theteam == "" || theteam == "null" || theteam == "undefined" || theteam == "All")
		theteam  = new String("All").toString();
}
teamtext = new String(theteam).toString();
teamclause = new String("");
if (theteam != "All") {
	teamclause = "and teamname like '"+theteam+"'";
	teamObj = getTeamByName(theteam);
	// teamObj = getTeam(1);
}
else {
	teamObj = getTeam(1);
}
// End of defaults - now construct SQL	
ConnObj=Server.CreateObject("ADODB.Connection");
RstObj=Server.CreateObject("ADODB.Recordset");
RS=Server.CreateObject("ADODB.Recordset");
ConnObj.Open(dbconnect);
if (thevenue == "BOTH")
	SQLStmt = new String("SELECT * from tennisfixtures where fixtureyear = "+theseason+" "+teamclause+" order by teamname,fixturedate");
else
	SQLStmt = new String("SELECT * from tennisfixtures where fixtureyear = "+theseason+" and homeoraway in ('"+thevenue+"') "+teamclause+" order by teamname,fixturedate");
if (debugging)
{
	Response.Write(SQLStmt+"<BR>");
	Response.End();
}
prevteam = new String("kshgkefshgblkesfj").toString();
curteam = new String("dbkjfdbkjrn").toString();
// Now fill up the league team names list
SQL5 = new String("SELECT * from tennisteams order by teamname");
RS = ConnObj.Execute(SQL5);
kount=1; // team IDs start at 1
teamnames[0] = new String("All").toString();
while (! RS.EOF)
{
	teamnames[kount] = new String(RS("teamname")).toString();
	kount=kount+1;
   	RS.MoveNext();
}
kount = kount -1 ;
RS=null;
displaydate = dateasstring(Date());
var debugging=current_debug_status();
// var debugging = true;
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
	<!-- Comment out the next style sheet if running in production -->
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
		<div id="resultsform">
			<form name="FixtureForm" action="lstfixtures.asp" method="post">
				<fieldset>
					<legend>Change selection</legend>
					<div>
						<b>Season:&nbsp;</b>
						<select name="season">
<%
for (i=thisyear; i > 2000; i--)
{
	dummy = new String(i).toString();
	Response.Write("<option>"+dummy);
}
%>
						</select>&nbsp;&nbsp;						
						<b>Team&nbsp;</b>
						<select name="team" id="team">
							<option>-- Select a Team --
<%
i=0;
while (i <= kount)
{
		dummy=new String(teamnames[i]).toString();
		if (dummy==theteam)
			Response.Write("<option selected>" + teamnames[i]+"\r\n");
		else
			Response.Write("<option>" + teamnames[i]+"\r\n");
		i = i + 1;
}
%>
					</select>&nbsp;&nbsp;
					<b>Venue:&nbsp;</b>
					<select name="venue">
						<option selected value="Both">Both
						<option value="H">Home
						<option value="A">Away
					</select>&nbsp;						
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					<input type="submit" value="Fetch"><br /><br />
					</div>
				</fieldset>
			</form>
		</div>	
<!--			
			<h2><%= teamtext %>&nbsp;<%= venuetext %>&nbsp;fixture list for the&nbsp;<%= theseason %>&nbsp;season</h2>
			 Choose what you wish to do with a fixture from the links in the <b>Action</b>&nbsp;column.<br />
			 Just click on an opponents team name to find out what Hampton's league record has been against them.
		</p>
-->
		<table id="fixturetable">
			<caption><span class="blueheading"><%= teamtext %></span>&nbsp;[<%= teamObj.division %>]&nbsp;fixture list for the&nbsp;<%= theseason %>&nbsp;season</caption>
			<thead>
				<tr>
					<th>Fixture<br />Date</th>
					<th style="text-align:center;">Venue</th>
					<th>Opponents</th>
					<th>First Pair<br />Second Pair</th>
					<th style="text-align:center;">Our<br />Score</th>
					<th style="text-align:center;">Their<br />Score</th>
					<th class="notprint" style="text-align:center;">Action</th>
				</tr>
			</thead>
			<tbody>
<%
	RstObj = ConnObj.Execute(SQLStmt);
	while (! RstObj.EOF)
	{
		curteam = new String(RstObj("teamname")).toString();
		// Now, put out heading line as appropriate
		if (curteam != prevteam)
		{
			if (theteam == "All")
			{
%>
				<tr>
					<td colspan="7" style="background-color:#DCE6CF; font-weight:bold; color:red;"><h4><%= curteam %></h4></td>
				</tr>
<%
			}
			prevteam = new String(curteam).toString();
		}							
		theid=new String(RstObj("fixtureid"));
		thenote = new String(RstObj("fixturenote"));
		if (thenote == "null" || thenote == "undefined")
				thenote = "";
		venuetext = "Home"
		if (RstObj("homeoraway")=="A")
			venuetext="Away";
		myDate = new Date(RstObj("fixturedate"));
		// thefixturedate = ddmmyyyy(myDate);
		thefixturedate = dateasstring(myDate);
		dayparts = thefixturedate.split(" ");
		stripeText=(isOdd)?'class="odd"':'';

		theirscore = new String(RstObj("opponentresult")).toString();
		ourscore = new String(RstObj("hamptonresult")).toString();
		if (ourscore == "" || ourscore == "null" || ourscore == "undefined")
		        ourscore = new String("0").toString();
		if (theirscore == "" || theirscore == "null" || theirscore == "undefined")
		        theirscore = new String("0").toString();
		dummy1 = parseInt(ourscore)
		dummy2 = parseInt(theirscore)
		totalus += dummy1;
		totalthem += dummy2;
		if (dummy1 == dummy2)
		{
		    if (dummy1 < 1)
		       	unknown += 1;
			else
				usdraw += 1;
		}
		if (dummy1 > dummy2)
		    uswin += 1;
		if (dummy1 < dummy2)
		    uslose += 1;

%>
				<tr <%= stripeText %>>
					<td><%= dayparts[0] %><br /><%= dayparts[1] %></td>
					<td style="text-align:center;"><%= venuetext %></td>
					<td><a href="lstteamresults.asp?theopposition=<%= RstObj("opponents") %>&layout=Full"><%= RstObj("opponents") %></a><br><b><%= thenote %></b></td>
					<td><%= RstObj("pair1") %><br /><%= RstObj("pair2") %></td>
					<td style="text-align:center;"><%= RstObj("hamptonresult") %></td>
					<td style="text-align:center;"><%= RstObj("opponentresult") %></td>
					<td class="notprint" style="text-align:center;">
<%
if (isFixtureEditor())
{
%>
					<a href="edit_fixture.asp?id=<%= theid %>">Result</a><br>
<%
}
if (isSupport())
{
%>
							<a href="delete_fixture.asp?id=<%= theid %>">Delete</a>
<%
}
%>
					</td>
				</tr>
<%
		isOdd = (! isOdd);
	RstObj.MoveNext();
	}
%>
			</tbody>
			<tfoot>
				<tr>
					<td colspan="7">
						&nbsp;
					</td>
				</tr>
			</tfoot>
		</table>		
		<br />
		<div>
		<p style="text-align: center;" style="width:50%;">
		<table>
			<caption>Summary results for the&nbsp;<%= theseason %>&nbsp;season</caption>
			<thead>
				<tr>
					<th>&nbsp;</th>
					<th>&nbsp;</th>
				</tr>
			</thead>
			<tbody>
				<tr>
					<td><b>Matches won by <%= teamtext %></b></td>
					<td style="text-align: center;"><%= uswin %></td>
				</tr>
				<tr>
					<td><b>Matches lost by <%= teamtext %></b></td>
					<td style="text-align: center;"><%= uslose %></td>
				</tr>
				<tr>
					<td><b>Matches drawn by <%= teamtext %></b></td>
					<td style="text-align: center;"><%= usdraw %></td>
				</tr>
				<tr>
					<td><b>Rubbers won by <%= teamtext %></b></td>
					<td style="text-align: center;"><%= totalus %></td>
				</tr>
				<tr>
					<td><b>Rubbers lost by <%= teamtext %></b></td>
					<td style="text-align: center;"><%= totalthem %></td>
				</tr>
				<tr>
					<td><b>Unknown results</b></td>
					<td style="text-align: center;"><%= unknown %></td>
				</tr>
			</tbody>
			<tfoot>
				<tr>
					<td colspan="2">
						&nbsp;
					</td>
				</tr>
			</tfoot>
		</table>		
		</p>
		</div>

<%
if ( isFixtureEditor() )	
{
%>
		<div id="fixtureadminarea" class="admin" style="clear:both;">
			<h3>Fixture Administration:</h3>
			<p>Set up brand new fixtures:<br />
			<ul>
				<li style="padding-left: 20px;"><a href="new_fixture.asp">for another team</a></li>
				<li style="padding-left: 20px;"><a href="new_fixture.asp?teamname=<%= teamtext %>&season=<%= theseason %>">for <%= teamtext %> in <%= theseason %></a></li>
			</ul>
			</p>
		</div>		
		
<%
}
%>		
	</div>
	
<!--     4.      Supplementary navigation    -->
	<div id="leftcolumn">
		<!--#include file="fixturepanel.asp" -->
		<!--#include file="memb_nav.asp" -->
	</div>

<!--    5.   Supplementary content     -->	
	<div id="rightcolumn">
		<!--#include file="searchpanel.asp" -->
		<!--#include file="addresspanel.asp" -->
		<!--#include file="newspanel.asp" -->
		<!--#include file="emailpanel.asp" -->
	</div>
</div>

<!--     6.    Site info     -->
<!--#include file="footer.asp" -->

</body>
</html>
<%
%>
