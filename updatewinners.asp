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
<!--#include file="winnerobj.asp" -->
<!--#include file="emailfuncs.asp" -->
<!--#include file="datefuncs.asp" -->
<%

var strtime, strdate, title2;
var clubname = new String("Hampton-In-Arden Sports Club");
var pagetitle = new String("Updating Tournament Titles");
// Now for any variables local to this page
var RS, Conn, SQL1, SQL2;
var dbconnect=Application("hamptonsportsdb"); 
var SQLstart, SQLmiddle, SQLend;
var resultObj = new Object();
var uniqueref;
var today = new Date();
var winnerdetail = new Object(); 
var winneryear, homeoraway, winnernote, teamname;
var winnerdate, winnerid, opponents;
var winners = new Array();
var defaultyear = currentYear();

var debugging = false;  // Production = false

function debugWrite(message) {
	if (debugging) {
		Response.Write(message);
	}
}

// Set up DB connections and unchanging bits of SQL
Conn = Server.CreateObject("ADODB.Connection");
RS = Server.CreateObject("ADODB.RecordSet");
Conn.Open(dbconnect);
SQLstart = new String("UPDATE winnersetup ")

// Retrieve POST'ed data

// One-offs first

debugWrite("thisteamfield = ["+Request.Form("teamname_0")+"]<br />");

teamname = Trim(new String(Request.Form("teamname_0")));
if (teamname == "" || teamname == "null" || teamname == "undefined")
{
	teamname = new String("").toString();
} 

winneryear = Trim(new String(Request.Form("winneryear_0")));
if (winneryear == "" || winneryear =="null" || winneryear == "undefined")
{
	winneryear = new String("").toString();
} 

debugWrite("teamname = ["+teamname+"], winneryear = ["+winneryear+"]<br />");


// Then a line at a time
// Starts at line 0 - ignore lines with no group code

// Line 0

winnerid = Trim(new String(Request.Form("winnerid_0")));
if (winnerid == "" || winnerid =="null" || winnerid == "undefined")
{
	winnerid = new String("").toString();
} 

winnerdate = Trim(new String(Request.Form("winnerdate_0")));
if (winnerdate == "" || winnerdate =="null" || winnerdate == "undefined")
{
	winnerdate = new String("").toString();
} 

homeoraway = Trim(new String(Request.Form("homeoraway_0")));
if (homeoraway == "" || homeoraway =="null" || homeoraway == "undefined")
{
	homeoraway = new String("H").toString();
} 

winnernote = Trim(new String(Request.Form("winnernote_0")));
if (winnernote == "" || winnernote =="null" || winnernote == "undefined")
{
	winnernote = new String("").toString();
} 

opponents = Trim(new String(Request.Form("opponents_0")));
if (opponents == "" || opponents =="null" || opponents == "undefined")
{
	opponents = new String("NONE").toString();
} 

debugWrite("winner Id = ["+winnerid+"], winnerdate = ["+winnerdate+"], venue = ["+homeoraway+"], opponents = ["+opponents+"]<br />");

winnerdetail = new winnerObject(winnerid,winnerdate,teamname,winneryear);
winnerdetail.setOpponents(opponents);
winnerdetail.setVenue(homeoraway);
winners.push(winnerdetail);

// Line 1


winnerid = Trim(new String(Request.Form("winnerid_1")));
if (winnerid == "" || winnerid =="null" || winnerid == "undefined")
{
	winnerid = new String("").toString();
} 

winnerdate = Trim(new String(Request.Form("winnerdate_1")));
if (winnerdate == "" || winnerdate =="null" || winnerdate == "undefined")
{
	winnerdate = new String("").toString();
} 

homeoraway = Trim(new String(Request.Form("homeoraway_1")));
if (homeoraway == "" || homeoraway =="null" || homeoraway == "undefined")
{
	homeoraway = new String("H").toString();
} 

winnernote = Trim(new String(Request.Form("winnernote_1")));
if (winnernote == "" || winnernote =="null" || winnernote == "undefined")
{
	winnernote = new String("").toString();
} 

opponents = Trim(new String(Request.Form("opponents_1")));
if (opponents == "" || opponents =="null" || opponents == "undefined")
{
	opponents = new String("NONE").toString();
} 

debugWrite("winner Id = ["+winnerid+"], winnerdate = ["+winnerdate+"], venue = ["+homeoraway+"], opponents = ["+opponents+"]<br />");

winnerdetail = new winnerObject(winnerid,winnerdate,teamname,winneryear);
winnerdetail.setOpponents(opponents);
winnerdetail.setVenue(homeoraway);
winners.push(winnerdetail);

// Line 2

winnerid = Trim(new String(Request.Form("winnerid_2")));
if (winnerid == "" || winnerid =="null" || winnerid == "undefined")
{
	winnerid = new String("").toString();
} 

winnerdate = Trim(new String(Request.Form("winnerdate_2")));
if (winnerdate == "" || winnerdate =="null" || winnerdate == "undefined")
{
	winnerdate = new String("").toString();
} 

homeoraway = Trim(new String(Request.Form("homeoraway_2")));
if (homeoraway == "" || homeoraway =="null" || homeoraway == "undefined")
{
	homeoraway = new String("H").toString();
} 

winnernote = Trim(new String(Request.Form("winnernote_2")));
if (winnernote == "" || winnernote =="null" || winnernote == "undefined")
{
	winnernote = new String("").toString();
} 

opponents = Trim(new String(Request.Form("opponents_2")));
if (opponents == "" || opponents =="null" || opponents == "undefined")
{
	opponents = new String("NONE").toString();
} 

debugWrite("winner Id = ["+winnerid+"], winnerdate = ["+winnerdate+"], venue = ["+homeoraway+"], opponents = ["+opponents+"]<br />");

winnerdetail = new winnerObject(winnerid,winnerdate,teamname,winneryear);
winnerdetail.setOpponents(opponents);
winnerdetail.setVenue(homeoraway);
winners.push(winnerdetail);

// Line 3

winnerid = Trim(new String(Request.Form("winnerid_3")));
if (winnerid == "" || winnerid =="null" || winnerid == "undefined")
{
	winnerid = new String("").toString();
} 

winnerdate = Trim(new String(Request.Form("winnerdate_3")));
if (winnerdate == "" || winnerdate =="null" || winnerdate == "undefined")
{
	winnerdate = new String("").toString();
} 

homeoraway = Trim(new String(Request.Form("homeoraway_3")));
if (homeoraway == "" || homeoraway =="null" || homeoraway == "undefined")
{
	homeoraway = new String("H").toString();
} 

winnernote = Trim(new String(Request.Form("winnernote_3")));
if (winnernote == "" || winnernote =="null" || winnernote == "undefined")
{
	winnernote = new String("").toString();
} 

opponents = Trim(new String(Request.Form("opponents_3")));
if (opponents == "" || opponents =="null" || opponents == "undefined")
{
	opponents = new String("NONE").toString();
} 

debugWrite("winner Id = ["+winnerid+"], winnerdate = ["+winnerdate+"], venue = ["+homeoraway+"], opponents = ["+opponents+"]<br />");

winnerdetail = new winnerObject(winnerid,winnerdate,teamname,winneryear);
winnerdetail.setOpponents(opponents);
winnerdetail.setVenue(homeoraway);
winners.push(winnerdetail);

// Line 4

winnerid = Trim(new String(Request.Form("winnerid_4")));
if (winnerid == "" || winnerid =="null" || winnerid == "undefined")
{
	winnerid = new String("").toString();
} 

winnerdate = Trim(new String(Request.Form("winnerdate_4")));
if (winnerdate == "" || winnerdate =="null" || winnerdate == "undefined")
{
	winnerdate = new String("").toString();
} 

winneryear = Trim(new String(Request.Form("winneryear_4")));
if (winneryear == "" || winneryear =="null" || winneryear == "undefined")
{
	winneryear = new String("00:00:00").toString();
} 

homeoraway = Trim(new String(Request.Form("homeoraway_4")));
if (homeoraway == "" || homeoraway =="null" || homeoraway == "undefined")
{
	homeoraway = new String("H").toString();
} 

winnernote = Trim(new String(Request.Form("winnernote_4")));
if (winnernote == "" || winnernote =="null" || winnernote == "undefined")
{
	winnernote = new String("").toString();
} 

opponents = Trim(new String(Request.Form("opponents_4")));
if (opponents == "" || opponents =="null" || opponents == "undefined")
{
	opponents = new String("NONE").toString();
} 

debugWrite("winner Id = ["+winnerid+"], winnerdate = ["+winnerdate+"], venue = ["+homeoraway+"], opponents = ["+opponents+"]<br />");

winnerdetail = new winnerObject(winnerid,winnerdate,teamname,winneryear);
winnerdetail.setOpponents(opponents);
winnerdetail.setVenue(homeoraway);
winners.push(winnerdetail);

// Line 5

winnerid = Trim(new String(Request.Form("winnerid_5")));
if (winnerid == "" || winnerid =="null" || winnerid == "undefined")
{
	winnerid = new String("").toString();
} 

winnerdate = Trim(new String(Request.Form("winnerdate_5")));
if (winnerdate == "" || winnerdate =="null" || winnerdate == "undefined")
{
	winnerdate = new String("").toString();
} 

homeoraway = Trim(new String(Request.Form("homeoraway_5")));
if (homeoraway == "" || homeoraway =="null" || homeoraway == "undefined")
{
	homeoraway = new String("H").toString();
} 

winnernote = Trim(new String(Request.Form("winnernote_5")));
if (winnernote == "" || winnernote =="null" || winnernote == "undefined")
{
	winnernote = new String("").toString();
} 

opponents = Trim(new String(Request.Form("opponents_5")));
if (opponents == "" || opponents =="null" || opponents == "undefined")
{
	opponents = new String("NONE").toString();
} 

debugWrite("winner Id = ["+winnerid+"], winnerdate = ["+winnerdate+"], venue = ["+homeoraway+"], opponents = ["+opponents+"]<br />");

winnerdetail = new winnerObject(winnerid,winnerdate,teamname,winneryear);
winnerdetail.setOpponents(opponents);
winnerdetail.setVenue(homeoraway);
winners.push(winnerdetail);

// Line 6

winnerid = Trim(new String(Request.Form("winnerid_6")));
if (winnerid == "" || winnerid =="null" || winnerid == "undefined")
{
	winnerid = new String("").toString();
} 

winnerdate = Trim(new String(Request.Form("winnerdate_6")));
if (winnerdate == "" || winnerdate =="null" || winnerdate == "undefined")
{
	winnerdate = new String("").toString();
} 

homeoraway = Trim(new String(Request.Form("homeoraway_6")));
if (homeoraway == "" || homeoraway =="null" || homeoraway == "undefined")
{
	homeoraway = new String("H").toString();
} 

winnernote = Trim(new String(Request.Form("winnernote_6")));
if (winnernote == "" || winnernote =="null" || winnernote == "undefined")
{
	winnernote = new String("").toString();
} 

opponents = Trim(new String(Request.Form("opponents_6")));
if (opponents == "" || opponents =="null" || opponents == "undefined")
{
	opponents = new String("NONE").toString();
} 

debugWrite("winner Id = ["+winnerid+"], winnerdate = ["+winnerdate+"], venue = ["+homeoraway+"], opponents = ["+opponents+"]<br />");

winnerdetail = new winnerObject(winnerid,winnerdate,teamname,winneryear);
winnerdetail.setOpponents(opponents);
winnerdetail.setVenue(homeoraway);
winners.push(winnerdetail);

// Line 7

winnerid = Trim(new String(Request.Form("winnerid_7")));
if (winnerid == "" || winnerid =="null" || winnerid == "undefined")
{
	winnerid = new String("").toString();
} 

winnerdate = Trim(new String(Request.Form("winnerdate_7")));
if (winnerdate == "" || winnerdate =="null" || winnerdate == "undefined")
{
	winnerdate = new String("").toString();
} 

homeoraway = Trim(new String(Request.Form("homeoraway_7")));
if (homeoraway == "" || homeoraway =="null" || homeoraway == "undefined")
{
	homeoraway = new String("H").toString();
} 

winnernote = Trim(new String(Request.Form("winnernote_7")));
if (winnernote == "" || winnernote =="null" || winnernote == "undefined")
{
	winnernote = new String("").toString();
} 

opponents = Trim(new String(Request.Form("opponents_7")));
if (opponents == "" || opponents =="null" || opponents == "undefined")
{
	opponents = new String("NONE").toString();
} 

debugWrite("winner Id = ["+winnerid+"], winnerdate = ["+winnerdate+"], venue = ["+homeoraway+"], opponents = ["+opponents+"]<br />");

winnerdetail = new winnerObject(winnerid,winnerdate,teamname,winneryear);
winnerdetail.setOpponents(opponents);
winnerdetail.setVenue(homeoraway);
winners.push(winnerdetail);

// Line 8

winnerid = Trim(new String(Request.Form("winnerid_8")));
if (winnerid == "" || winnerid =="null" || winnerid == "undefined")
{
	winnerid = new String("").toString();
} 

winnerdate = Trim(new String(Request.Form("winnerdate_8")));
if (winnerdate == "" || winnerdate =="null" || winnerdate == "undefined")
{
	winnerdate = new String("").toString();
} 

homeoraway = Trim(new String(Request.Form("homeoraway_8")));
if (homeoraway == "" || homeoraway =="null" || homeoraway == "undefined")
{
	homeoraway = new String("H").toString();
} 

winnernote = Trim(new String(Request.Form("winnernote_8")));
if (winnernote == "" || winnernote =="null" || winnernote == "undefined")
{
	winnernote = new String("").toString();
} 

opponents = Trim(new String(Request.Form("opponents_8")));
if (opponents == "" || opponents =="null" || opponents == "undefined")
{
	opponents = new String("NONE").toString();
} 

debugWrite("winner Id = ["+winnerid+"], winnerdate = ["+winnerdate+"], venue = ["+homeoraway+"], opponents = ["+opponents+"]<br />");

winnerdetail = new winnerObject(winnerid,winnerdate,teamname,winneryear);
winnerdetail.setOpponents(opponents);
winnerdetail.setVenue(homeoraway);
winners.push(winnerdetail);

// Line 9

winnerid = Trim(new String(Request.Form("winnerid_9")));
if (winnerid == "" || winnerid =="null" || winnerid == "undefined")
{
	winnerid = new String("").toString();
} 

winnerdate = Trim(new String(Request.Form("winnerdate_9")));
if (winnerdate == "" || winnerdate =="null" || winnerdate == "undefined")
{
	winnerdate = new String("").toString();
} 

homeoraway = Trim(new String(Request.Form("homeoraway_9")));
if (homeoraway == "" || homeoraway =="null" || homeoraway == "undefined")
{
	homeoraway = new String("H").toString();
} 

winnernote = Trim(new String(Request.Form("winnernote_9")));
if (winnernote == "" || winnernote =="null" || winnernote == "undefined")
{
	winnernote = new String("").toString();
} 

opponents = Trim(new String(Request.Form("opponents_9")));
if (opponents == "" || opponents =="null" || opponents == "undefined")
{
	opponents = new String("NONE").toString();
} 

debugWrite("winner Id = ["+winnerid+"], winnerdate = ["+winnerdate+"], venue = ["+homeoraway+"], opponents = ["+opponents+"]<br />");

winnerdetail = new winnerObject(winnerid,winnerdate,teamname,winneryear);
winnerdetail.setOpponents(opponents);
winnerdetail.setVenue(homeoraway);
winners.push(winnerdetail);

// Line 10

winnerid = Trim(new String(Request.Form("winnerid_10")));
if (winnerid == "" || winnerid =="null" || winnerid == "undefined")
{
	winnerid = new String("").toString();
} 

winnerdate = Trim(new String(Request.Form("winnerdate_10")));
if (winnerdate == "" || winnerdate =="null" || winnerdate == "undefined")
{
	winnerdate = new String("").toString();
} 

homeoraway = Trim(new String(Request.Form("homeoraway_10")));
if (homeoraway == "" || homeoraway =="null" || homeoraway == "undefined")
{
	homeoraway = new String("H").toString();
} 

winnernote = Trim(new String(Request.Form("winnernote_10")));
if (winnernote == "" || winnernote =="null" || winnernote == "undefined")
{
	winnernote = new String("").toString();
} 

opponents = Trim(new String(Request.Form("opponents_10")));
if (opponents == "" || opponents =="null" || opponents == "undefined")
{
	opponents = new String("NONE").toString();
} 

debugWrite("winner Id = ["+winnerid+"], winnerdate = ["+winnerdate+"], venue = ["+homeoraway+"], opponents = ["+opponents+"]<br />");

winnerdetail = new winnerObject(winnerid,winnerdate,teamname,winneryear);
winnerdetail.setOpponents(opponents);
winnerdetail.setVenue(homeoraway);
winners.push(winnerdetail);

// Line 11

winnerid = Trim(new String(Request.Form("winnerid_11")));
if (winnerid == "" || winnerid =="null" || winnerid == "undefined")
{
	winnerid = new String("").toString();
} 

winnerdate = Trim(new String(Request.Form("winnerdate_11")));
if (winnerdate == "" || winnerdate =="null" || winnerdate == "undefined")
{
	winnerdate = new String("").toString();
} 

homeoraway = Trim(new String(Request.Form("homeoraway_11")));
if (homeoraway == "" || homeoraway =="null" || homeoraway == "undefined")
{
	homeoraway = new String("H").toString();
} 

winnernote = Trim(new String(Request.Form("winnernote_11")));
if (winnernote == "" || winnernote =="null" || winnernote == "undefined")
{
	winnernote = new String("").toString();
} 

opponents = Trim(new String(Request.Form("opponents_11")));
if (opponents == "" || opponents =="null" || opponents == "undefined")
{
	opponents = new String("NONE").toString();
} 

debugWrite("winner Id = ["+winnerid+"], winnerdate = ["+winnerdate+"], venue = ["+homeoraway+"], opponents = ["+opponents+"]<br />");

winnerdetail = new winnerObject(winnerid,winnerdate,teamname,winneryear);
winnerdetail.setOpponents(opponents);
winnerdetail.setVenue(homeoraway);
winners.push(winnerdetail);

// Line 12

winnerid = Trim(new String(Request.Form("winnerid_12")));
if (winnerid == "" || winnerid =="null" || winnerid == "undefined")
{
	winnerid = new String("").toString();
} 

winnerdate = Trim(new String(Request.Form("winnerdate_12")));
if (winnerdate == "" || winnerdate =="null" || winnerdate == "undefined")
{
	winnerdate = new String("").toString();
} 

homeoraway = Trim(new String(Request.Form("homeoraway_12")));
if (homeoraway == "" || homeoraway =="null" || homeoraway == "undefined")
{
	homeoraway = new String("H").toString();
} 

winnernote = Trim(new String(Request.Form("winnernote_12")));
if (winnernote == "" || winnernote =="null" || winnernote == "undefined")
{
	winnernote = new String("").toString();
} 

opponents = Trim(new String(Request.Form("opponents_12")));
if (opponents == "" || opponents =="null" || opponents == "undefined")
{
	opponents = new String("NONE").toString();
} 

debugWrite("winner Id = ["+winnerid+"], winnerdate = ["+winnerdate+"], venue = ["+homeoraway+"], opponents = ["+opponents+"]<br />");

winnerdetail = new winnerObject(winnerid,winnerdate,teamname,winneryear);
winnerdetail.setOpponents(opponents);
winnerdetail.setVenue(homeoraway);
winners.push(winnerdetail);

// Line 13

winnerid = Trim(new String(Request.Form("winnerid_13")));
if (winnerid == "" || winnerid =="null" || winnerid == "undefined")
{
	winnerid = new String("").toString();
} 

winnerdate = Trim(new String(Request.Form("winnerdate_13")));
if (winnerdate == "" || winnerdate =="null" || winnerdate == "undefined")
{
	winnerdate = new String("").toString();
} 

homeoraway = Trim(new String(Request.Form("homeoraway_13")));
if (homeoraway == "" || homeoraway =="null" || homeoraway == "undefined")
{
	homeoraway = new String("H").toString();
} 

winnernote = Trim(new String(Request.Form("winnernote_13")));
if (winnernote == "" || winnernote =="null" || winnernote == "undefined")
{
	winnernote = new String("").toString();
} 

opponents = Trim(new String(Request.Form("opponents_13")));
if (opponents == "" || opponents =="null" || opponents == "undefined")
{
	opponents = new String("NONE").toString();
} 

debugWrite("winner Id = ["+winnerid+"], winnerdate = ["+winnerdate+"], venue = ["+homeoraway+"], opponents = ["+opponents+"]<br />");

winnerdetail = new winnerObject(winnerid,winnerdate,teamname,winneryear);
winnerdetail.setOpponents(opponents);
winnerdetail.setVenue(homeoraway);
winners.push(winnerdetail);

// Line 14

winnerid = Trim(new String(Request.Form("winnerid_14")));
if (winnerid == "" || winnerid =="null" || winnerid == "undefined")
{
	winnerid = new String("").toString();
} 

winnerdate = Trim(new String(Request.Form("winnerdate_14")));
if (winnerdate == "" || winnerdate =="null" || winnerdate == "undefined")
{
	winnerdate = new String("").toString();
} 

homeoraway = Trim(new String(Request.Form("homeoraway_14")));
if (homeoraway == "" || homeoraway =="null" || homeoraway == "undefined")
{
	homeoraway = new String("H").toString();
} 

winnernote = Trim(new String(Request.Form("winnernote_14")));
if (winnernote == "" || winnernote =="null" || winnernote == "undefined")
{
	winnernote = new String("").toString();
} 

opponents = Trim(new String(Request.Form("opponents_14")));
if (opponents == "" || opponents =="null" || opponents == "undefined")
{
	opponents = new String("NONE").toString();
} 

debugWrite("winner Id = ["+winnerid+"], winnerdate = ["+winnerdate+"], venue = ["+homeoraway+"], opponents = ["+opponents+"]<br />");

winnerdetail = new winnerObject(winnerid,winnerdate,teamname,winneryear);
winnerdetail.setOpponents(opponents);
winnerdetail.setVenue(homeoraway);
winners.push(winnerdetail);

//      End of the potential winners

debugWrite("winners = "+JSON.stringify(winners)+"<br /><hr />");

// Update winner details from POST'ed data
// Loop through all the winners, update via winnerid as unique row identifier

for (var j=0; j<winners.length; j++) {

	winnerdetail = winners[j];

	if (! (winnerdetail.opponents == "NONE")) {   // ignore if no opponents

		SQLend = new String(" WHERE winnerid = "+winnerdetail.winnerid).toString();

		SQLmiddle = new String("SET ").toString();
		SQLmiddle += " winneryear="+winnerdetail.winneryear+",";
		SQLmiddle += " homeoraway='"+winnerdetail.homeoraway+"',";
		SQLmiddle += " winnernote='"+winnerdetail.winnernote+"',";
		SQLmiddle += " teamname='"+winnerdetail.teamname+"', ";
		SQLmiddle += " opponents='"+winnerdetail.opponents+"', ";

		// Now do date fields. If null dont insert them as part of the update clause
		//  Access doesnt like setting date fields to ''

		if (! (winnerdetail.winnerdate == ""))
			SQLmiddle += " winnerdate='"+winnerdetail.winnerdate+"' ";
		else
			SQLmiddle += " winnerdate=null ";

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
				resultObj.sql = new String(SQL1).toString();
			}
		}

		debugWrite("SQL used for update: <br />"+resultObj.sql+"<br /><br />");

	}

}

// On completion, redirect appropriately

if (! debugging) {
	Response.Redirect("./winnersetup.html#/");
}

Response.End();


%>


