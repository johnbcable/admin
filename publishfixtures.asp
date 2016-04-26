<%@language="JScript" CODEPAGE="65001"%>
<%
Response.AddHeader("Cache-Control", "no-cache,no-store,must-revalidate");
Response.AddHeader("Pragma", "no-cache");
Response.AddHeader("Expires", 0);
Response.AddHeader("Access-Control-Allow-Origin", "*");
%>
<!--#include file="json2.js.asp" -->
<!--#include file="functions.asp" -->
<!--#include file="strings.asp" -->
<!--#include file="fixtureobj.asp" -->
<!--#include file="emailfuncs.asp" -->
<!--#include file="datefuncs.asp" -->
<%

var strtime, strdate, title2;
var clubname = new String("Hampton-In-Arden Sports Club");
var pagetitle = new String("Updating Fixture Schedule");
// Now for any variables local to this page
var RS, RS2, Conn, SQL1, SQL2;
var dbconnect=Application("hamptonsportsdb"); 
var SQLstart, SQLmiddle, SQLend;
var resultObj = new Object();
var today = new Date();
var fixturedetail = new Object(); 
var fixtureyear, homeoraway, fixturenote, teamname;
var fixturedate, fixtureid, opponents;
var fixtures = new Array();
var defaultyear = currentYear();
var destinationtable;

var debugging = false;  // Production = false
var updating = true;   // Production = true

function debugWrite(message) {
	if (debugging) {
		Response.Write(message);
		Response.Flush();
	}
}

// Set up DB connections and unchanging bits of SQL
Conn = Server.CreateObject("ADODB.Connection");
RS = Server.CreateObject("ADODB.RecordSet");
RS2 = Server.CreateObject("ADODB.RecordSet");
Conn.Open(dbconnect);

// Process QueryString for run options
mydebug = Request.QueryString("debug");
if (mydebug == "" || mydebug == "null" || mydebug == "undefined")
{
	mydebug = new String("N").toString();
}
if (mydebug == "Y") {
	destinationtable = new String("newfixtures").toString();
	debugging = true;
} else {
	destinationtable = new String("tennisfixtures").toString();
	debugging = false;
}
SQLstart = new String("INSERT INTO "+destinationtable+"([fixturedate],[homeoraway],[opponents],[fixtureyear],[teamname]) ")

fixtureyear = Request.QueryString("year");
if (fixtureyear == "" || fixtureyear == "null" || fixtureyear == "undefined")
{
	fixtureyear = Request.Form("year");
	if (fixtureyear == "" || fixtureyear == "null" || fixtureyear == "undefined")
	{
		fixtureyear = currentSeason();   // default to the current currentSeason
	}
}

//      End of the parameter processing

// Make sure and delete any prior fixtures for this year from tennisfixtures

SQL = new String("DELETE FROM tennisfixtures WHERE fixtureyear = "+fixtureyear);
if (debugging) 
{
	Response.Write("Deletion SQL: "+SQL)
} 
else 
{
	if ( updating ) 
	{
		RS=Conn.Execute(SQL);
	}
}

// Now set up SQL to retrieve from fixturesetup

SQL = new String("SELECT * FROM fixturesetup ORDER BY fixtureid");
RS = Conn.Execute(SQL)

while (! RS.EOF)
{
	// fixtureyear set up outside this loop
	fixturedate = new String(RS("fixturedate")).toString();
	homeoraway = new String(RS("homeoraway")).toString();
	fixturenote = new String("").toString();
	teamname = new String(RS("teamname")).toString();
	opponents = new String(RS("opponents")).toString();

	debugWrite("fixturedate="+fixturedate+", homeoraway="+homeoraway+", teamname="+teamname+", opponents="+opponents+"<br/>");

	if (! (opponents == "NONE")) {   // ignore if no opponents

		SQLend = new String(")").toString();

		SQLmiddle = new String("VALUES (").toString();
		if (! fixturedate == "" )
			SQLmiddle += " '"+fixturedate+"',";
		else
			SQLmiddle += " null,";

		SQLmiddle += " '"+homeoraway+"',";
		SQLmiddle += " '"+opponents+"',";
		SQLmiddle += " "+fixtureyear+", ";
		SQLmiddle += " "+teamname+"' ";

		// Default values into result object
		resultObj.result = true;
		resultObj.errno = 0;
		resultObj.description = new String("").toString();

		SQL1 = new String(SQLstart+SQLmiddle+SQLend).toString();
		resultObj.sql = new String(SQL1).toString();

		if ( updating ) {
			try {
				RS2 = Conn.Execute(SQL1);
			}
			catch(e) {
				resultObj.result = false;
				resultObj.errno = (e.number & 0xFFFF);
				resultObj.description += e.description;
				resultObj.sql = new String(SQL1).toString();
			}
		}

		debugWrite("SQL used for update: <br />"+resultObj.sql+"<br /><br />");

	}

	RS.MoveNext();
	
}

// On completion, redirect appropriately - back to admin home page area

if (! debugging) {
	Response.Redirect("/admin");
}

Response.End();


%>


