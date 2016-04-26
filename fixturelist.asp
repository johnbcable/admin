<%@language="JScript" CODEPAGE="65001"%>
<!--#include file="functions.asp" -->
<!--#include file="strings.asp" -->
<!--#include file="emailfuncs.asp" -->
<%
// First variables that need to be set for each page
var strtime, strdate;
var clubname = new String("Hampton-In-Arden Sports Club");
var pagetitle = new String("League Fixtures");
// Now for any variables local to this page
var dbconnect=Application("hamptonsportsdb");
var ConnObj, RstObj, StnObj;
var ConnObj2, RstObj2, StnObj2;
var RS, RS2, RS3;
var d, thisyear;
var i, kount, dummy, dummy1;
var teamnames = new Array(10);
var teamids = new Array(10);
var curteam, curseason;
var teamcaptain, teamname;
var venuetext, teamnote;
var thefixturedate, strfixturedate;
var isOdd = true;
var stripeText = new String("").toString();
var teamObj = new Object();
// Set up default greeting strings
strdate = datestring();
strtime = timestring();
thisyear = "";
d = new Date();
thisyear += d.getYear();
d = null;
curteam = new String(Request.QueryString("team")).toString();
if (curteam == "null" || curteam == "undefined")
	curteam = new String("FRIENDLY").toString();
curseason = new String(Request.QueryString("season")).toString();
if (curseason == "null" || curseason == "undefined")
	curseason = new String(thisyear).toString();
// Check if signed-in and can administer fixtures
if (signedin())
{
	if (isAdministrator())
		Response.Redirect("lstfixtures.asp?season="+curseason+"&team="+curteam);
}
// End of page start up coding
ConnObj=Server.CreateObject("ADODB.Connection");
RstObj=Server.CreateObject("ADODB.Recordset");
ConnObj.Open(dbconnect);
SQLStmt = new String("SELECT * from tennisteams where teamcategory = 'ADULT' order by teamname");
RstObj = ConnObj.Execute(SQLStmt);
kount=1; // team IDs start at 1
teamcaptain="";
teamname="";
teamnote="";
while (! RstObj.EOF)
{
		thisteam = new String(RstObj("teamname")).toString();
		teamnames[kount] = new String(RstObj("teamname")).toString();
		if (curteam==thisteam)
		{
			teamcaptain = new String(RstObj("teamcaptain")).toString();
			teamname=new String(teamnames[kount]).toString();
			teamnote = new String(RstObj("teamnote")).toString();
		}
		kount=kount+1;
    	RstObj.MoveNext();
}
if (teamnote == "null" || teamnote == "undefined")
	teamnote = "";
kount = kount -1 ;
RstObj=null;
pagetitle += " for "+teamname+": "+curseason;
displaydate = dateasstring(Date());
var debugging=current_debug_status();
// teamObj = getTeamByName(thisteam);
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
<style type="text/css">
#fixturetable tbody tr.odd, #fixturetable tbody tr.odd td {
  background-color:#ccc;
}
#fixturetable tbody td {
	border: 0;
	padding: 10px 10px 10px 0;
	border-bottom: 1px solid black;
}
</style>
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
			<h1>Playing:<b>&nbsp;League Fixtures</b></h1>
			<div id="fixturelistblurb">
				<p>
					Every effort is made to ensure that the following information is up-to-date.
					However, it cannot reflect last minute changes or cancellations. Please check
					the details with your captain before any match.
				</p>
			</div>
			
			<form name="frmJump" action="fixturelist.asp" method="GET">
				<fieldset>
					<legend>Fixturelist search parameters</legend>
					<div>
						<label for="team">Fixture list for which team?&nbsp;</label>
						<select name="team">
							<option>-- Select a Team --
<%
i=1;
while (i <= kount)
{
		dummy=new String(i).toString();
		if (dummy==curteam)
			Response.Write("<option selected>" + teamnames[i]+"\r\n");
		else
			Response.Write("<option>" + teamnames[i]+"\r\n");
		i = i + 1;
}
%>
						</select>
					</div>
					<div>
						<label for="season">For which season?</label>
						<select name="season">
<%
for (i=thisyear; i > 2000; i--)
{
	dummy = new String(i).toString();
	Response.Write("<option>"+dummy);
}
%>
						</select>
					</div>
					<div class="boxright" style="margin-right:40px;">
						<input type="submit" value="Fetch">
					</div>
				</fieldset>
			</form>
			
			<!-- Now output the fixture list -->
			
			<div id="fixturesearchresults" class="borderboth">
				<table id="fixturetable">
					<caption>
						<span class="blueheading"><%= curteam %></b>&nbsp;fixtures for the <b><%= curseason %></b> season.</span><br />
						Current Captain:<b>&nbsp;<%= teamcaptain %></b>&nbsp;&nbsp;&nbsp;<i><b><%= teamnote %></b></i>
					</caption>
					<thead>
						<tr>
							<th>Fixture<br />Date</th>
							<th>Venue</th>
							<th>Opponents</th>
							<th>First Pair<br />Second Pair</th>
							<th>Our<br />Score</th>
							<th>Their<br />Score</th>
						</tr>
					</thead>
					<tbody>
<%
	SQLStmt = new String("SELECT * from tennisfixtures where fixtureyear = "+curseason+" and teamname = '"+teamname+"' order by fixturedate");
	if (debugging)
		Response.Write(SQLStmt+"<BR>");
	RstObj = ConnObj.Execute(SQLStmt);
	while (! RstObj.EOF)
	{
		venuetext = "Home"
		if (RstObj("homeoraway")=="A")
			venuetext="Away";
		thefixturedate=new Date(RstObj("fixturedate"));
		strfixturedate = dateasstring(thefixturedate);
		dummy1 = new String(RstObj("fixturenote")).toString();
		if (dummy1 == "null" || dummy1 == "undefined")
				dummy1 = "";
		dummy2 = new String(RstObj("matchreport")).toString();
		if (dummy2 == "null" || dummy2 == "undefined")
			dummy2 = "";
		stripeText=(isOdd)?'class="odd"':'';
%>
	<tr <%= stripeText %>>
		<td>
			<%= strfixturedate %>
		</td>
		<td style="text-align:center;">
			<%= venuetext %>
		</td>
		<td>
			<a href="lstteamresults.asp?theopposition=<%= RstObj("opponents") %>&layout=Full"><%= RstObj("opponents") %></a><br><b><%= dummy1 %></b>
<%
if (! (dummy2 == ""))
{
%>
			<br>
			<span align="right" style="margin-right: 4px"><a href="showfixturestory.asp?matchreport=<%= dummy2 %>">More ..</a>
<%
}
%>
		</td>
		<td>
			<%= RstObj("pair1") %><br><%= RstObj("pair2") %>
		</td>
		<td style="text-align:center;">
			<%= RstObj("hamptonresult") %> &nbsp;
		</td>
		<td style="text-align:center;">
			<%= RstObj("opponentresult") %> &nbsp;
		</td>
	</tr>
<%
		isOdd = (! isOdd);
		RstObj.MoveNext();
	}
%>
			</table>
		</div>
		<div class="seealso">
			<h3 class="blueheading">See Also:</h3>
			<p class="leftindent">
				Like to check up on current league positions - click <a href="http://www.coventrytennisleague.co.uk/">here</a>.
			</p>
			<p class="leftindent">
				Need help with directions to clubs? Click <A HREF="directions.asp">here</A> to view our club address list.
			</p>		
		</div>
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


