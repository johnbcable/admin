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
var RS, Conn, SQL1, SQL2;
var dbconnect=Application("hamptonsportsdb"); 
var SQLstart, SQLmiddle, SQLend;
var resultObj = new Object();
var uniqueref;
var today = new Date();
var fixturedetail = new Object(); 
var fixtureyear, homeoraway, fixturenote, teamname;
var fixturedate, fixtureid, opponents;
var fixtures = new Array();
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
SQLstart = new String("UPDATE tennisfixtures ")

// Retrieve POST'ed data

// One-offs first

debugWrite("thisteamfield = ["+Request.Form("teamname_0")+"]<br />");

teamname = Trim(new String(Request.Form("teamname_0")));
if (teamname == "" || teamname == "null" || teamname == "undefined")
{
	teamname = new String("").toString();
} 

fixtureyear = Trim(new String(Request.Form("fixtureyear_0")));
if (fixtureyear == "" || fixtureyear =="null" || fixtureyear == "undefined")
{
	fixtureyear = new String("").toString();
} 

debugWrite("teamname = ["+teamname+"], fixtureyear = ["+fixtureyear+"]<br />");


// Then a line at a time
// Starts at line 0 - ignore lines with no group code

// Line 0

fixtureid = Trim(new String(Request.Form("fixtureid_0")));
if (fixtureid == "" || fixtureid =="null" || fixtureid == "undefined")
{
	fixtureid = new String("").toString();
} 

fixturedate = Trim(new String(Request.Form("fixturedate_0")));
if (fixturedate == "" || fixturedate =="null" || fixturedate == "undefined")
{
	fixturedate = new String("").toString();
} 

homeoraway = Trim(new String(Request.Form("homeoraway_0")));
if (homeoraway == "" || homeoraway =="null" || homeoraway == "undefined")
{
	homeoraway = new String("H").toString();
} 

fixturenote = Trim(new String(Request.Form("fixturenote_0")));
if (fixturenote == "" || fixturenote =="null" || fixturenote == "undefined")
{
	fixturenote = new String(" ").toString();
} 

opponents = Trim(new String(Request.Form("opponents_0")));
if (opponents == "" || opponents =="null" || opponents == "undefined")
{
	opponents = new String("NONE").toString();
} 

pair1 = Trim(new String(Request.Form("pair1_0")));
if (pair1 == "" || pair1 =="null" || pair1 == "undefined")
{
	pair1 = new String("").toString();
} 

pair2 = Trim(new String(Request.Form("pair2_0")));
if (pair2 == "" || pair2 =="null" || pair2 == "undefined")
{
	pair2 = new String("").toString();
} 

hamptonresult = new Number(Request.Form("hamptonresult_0"));
if (hamptonresult == "" || hamptonresult =="null" || hamptonresult == "undefined")
{
	hamptonresult = 0;
} 

opponentresult = new Number(Request.Form("opponentresult_0"));
if (opponentresult == "" || opponentresult =="null" || opponentresult == "undefined")
{
	opponentresult = 0;
} 

debugWrite("fixture Id = ["+fixtureid+"], fixturedate = ["+fixturedate+"], venue = ["+homeoraway+"], opponents = ["+opponents+"], pair1 = ["+pair1+"], pair2 = ["+pair2+"], hamptonresult = ["+hamptonresult+"], opponentresult = ["+opponentresult+"]<br />");

fixturedetail = new FixtureObject(fixtureid,fixturedate,teamname,fixtureyear);
fixturedetail.setOpponents(opponents);
fixturedetail.setVenue(homeoraway);
fixturedetail.setPairs(pair1, pair2);
fixturedetail.setResult(hamptonresult, opponentresult)
fixtures.push(fixturedetail);

// Line 1


fixtureid = Trim(new String(Request.Form("fixtureid_1")));
if (fixtureid == "" || fixtureid =="null" || fixtureid == "undefined")
{
	fixtureid = new String("").toString();
} 

fixturedate = Trim(new String(Request.Form("fixturedate_1")));
if (fixturedate == "" || fixturedate =="null" || fixturedate == "undefined")
{
	fixturedate = new String("").toString();
} 

homeoraway = Trim(new String(Request.Form("homeoraway_1")));
if (homeoraway == "" || homeoraway =="null" || homeoraway == "undefined")
{
	homeoraway = new String("H").toString();
} 

fixturenote = Trim(new String(Request.Form("fixturenote_1")));
if (fixturenote == "" || fixturenote =="null" || fixturenote == "undefined")
{
	fixturenote = new String(" ").toString();
} 

opponents = Trim(new String(Request.Form("opponents_1")));
if (opponents == "" || opponents =="null" || opponents == "undefined")
{
	opponents = new String("NONE").toString();
} 

pair1 = Trim(new String(Request.Form("pair1_1")));
if (pair1 == "" || pair1 =="null" || pair1 == "undefined")
{
	pair1 = new String("").toString();
} 

pair2 = Trim(new String(Request.Form("pair2_1")));
if (pair2 == "" || pair2 =="null" || pair2 == "undefined")
{
	pair2 = new String("").toString();
} 

hamptonresult = new Number(Request.Form("hamptonresult_1"));
if (hamptonresult == "" || hamptonresult =="null" || hamptonresult == "undefined")
{
	hamptonresult = 0;
} 

opponentresult = new Number(Request.Form("opponentresult_1"));
if (opponentresult == "" || opponentresult =="null" || opponentresult == "undefined")
{
	opponentresult = 0;
} 

debugWrite("fixture Id = ["+fixtureid+"], fixturedate = ["+fixturedate+"], venue = ["+homeoraway+"], opponents = ["+opponents+"], pair1 = ["+pair1+"], pair2 = ["+pair2+"], hamptonresult = ["+hamptonresult+"], opponentresult = ["+opponentresult+"]<br />");

fixturedetail = new FixtureObject(fixtureid,fixturedate,teamname,fixtureyear);
fixturedetail.setOpponents(opponents);
fixturedetail.setVenue(homeoraway);
fixturedetail.setPairs(pair1, pair2);
fixturedetail.setResult(hamptonresult, opponentresult)
fixtures.push(fixturedetail);

// Line 2

fixtureid = Trim(new String(Request.Form("fixtureid_2")));
if (fixtureid == "" || fixtureid =="null" || fixtureid == "undefined")
{
	fixtureid = new String("").toString();
} 

fixturedate = Trim(new String(Request.Form("fixturedate_2")));
if (fixturedate == "" || fixturedate =="null" || fixturedate == "undefined")
{
	fixturedate = new String("").toString();
} 

homeoraway = Trim(new String(Request.Form("homeoraway_2")));
if (homeoraway == "" || homeoraway =="null" || homeoraway == "undefined")
{
	homeoraway = new String("H").toString();
} 

fixturenote = Trim(new String(Request.Form("fixturenote_2")));
if (fixturenote == "" || fixturenote =="null" || fixturenote == "undefined")
{
	fixturenote = new String(" ").toString();
} 

opponents = Trim(new String(Request.Form("opponents_2")));
if (opponents == "" || opponents =="null" || opponents == "undefined")
{
	opponents = new String("NONE").toString();
} 

pair1 = Trim(new String(Request.Form("pair1_2")));
if (pair1 == "" || pair1 =="null" || pair1 == "undefined")
{
	pair1 = new String("").toString();
} 

pair2 = Trim(new String(Request.Form("pair2_2")));
if (pair2 == "" || pair2 =="null" || pair2 == "undefined")
{
	pair2 = new String("").toString();
} 

hamptonresult = new Number(Request.Form("hamptonresult_2"));
if (hamptonresult == "" || hamptonresult =="null" || hamptonresult == "undefined")
{
	hamptonresult = 0;
} 

opponentresult = new Number(Request.Form("opponentresult_2"));
if (opponentresult == "" || opponentresult =="null" || opponentresult == "undefined")
{
	opponentresult = 0;
} 

debugWrite("fixture Id = ["+fixtureid+"], fixturedate = ["+fixturedate+"], venue = ["+homeoraway+"], opponents = ["+opponents+"], pair1 = ["+pair1+"], pair2 = ["+pair2+"], hamptonresult = ["+hamptonresult+"], opponentresult = ["+opponentresult+"]<br />");

fixturedetail = new FixtureObject(fixtureid,fixturedate,teamname,fixtureyear);
fixturedetail.setOpponents(opponents);
fixturedetail.setVenue(homeoraway);
fixturedetail.setPairs(pair1, pair2);
fixturedetail.setResult(hamptonresult, opponentresult)
fixtures.push(fixturedetail);


// Line 3

fixtureid = Trim(new String(Request.Form("fixtureid_3")));
if (fixtureid == "" || fixtureid =="null" || fixtureid == "undefined")
{
	fixtureid = new String("").toString();
} 

fixturedate = Trim(new String(Request.Form("fixturedate_3")));
if (fixturedate == "" || fixturedate =="null" || fixturedate == "undefined")
{
	fixturedate = new String("").toString();
} 

homeoraway = Trim(new String(Request.Form("homeoraway_3")));
if (homeoraway == "" || homeoraway =="null" || homeoraway == "undefined")
{
	homeoraway = new String("H").toString();
} 

fixturenote = Trim(new String(Request.Form("fixturenote_3")));
if (fixturenote == "" || fixturenote =="null" || fixturenote == "undefined")
{
	fixturenote = new String(" ").toString();
} 

opponents = Trim(new String(Request.Form("opponents_3")));
if (opponents == "" || opponents =="null" || opponents == "undefined")
{
	opponents = new String("NONE").toString();
} 

pair1 = Trim(new String(Request.Form("pair1_3")));
if (pair1 == "" || pair1 =="null" || pair1 == "undefined")
{
	pair1 = new String("").toString();
} 

pair2 = Trim(new String(Request.Form("pair2_3")));
if (pair2 == "" || pair2 =="null" || pair2 == "undefined")
{
	pair2 = new String("").toString();
} 

hamptonresult = new Number(Request.Form("hamptonresult_3"));
if (hamptonresult == "" || hamptonresult =="null" || hamptonresult == "undefined")
{
	hamptonresult = 0;
} 

opponentresult = new Number(Request.Form("opponentresult_3"));
if (opponentresult == "" || opponentresult =="null" || opponentresult == "undefined")
{
	opponentresult = 0;
} 

debugWrite("fixture Id = ["+fixtureid+"], fixturedate = ["+fixturedate+"], venue = ["+homeoraway+"], opponents = ["+opponents+"], pair1 = ["+pair1+"], pair2 = ["+pair2+"], hamptonresult = ["+hamptonresult+"], opponentresult = ["+opponentresult+"]<br />");

fixturedetail = new FixtureObject(fixtureid,fixturedate,teamname,fixtureyear);
fixturedetail.setOpponents(opponents);
fixturedetail.setVenue(homeoraway);
fixturedetail.setPairs(pair1, pair2);
fixturedetail.setResult(hamptonresult, opponentresult)
fixtures.push(fixturedetail);


// Line 4

fixtureid = Trim(new String(Request.Form("fixtureid_4")));
if (fixtureid == "" || fixtureid =="null" || fixtureid == "undefined")
{
	fixtureid = new String("").toString();
} 

fixturedate = Trim(new String(Request.Form("fixturedate_4")));
if (fixturedate == "" || fixturedate =="null" || fixturedate == "undefined")
{
	fixturedate = new String("").toString();
} 

fixtureyear = Trim(new String(Request.Form("fixtureyear_4")));
if (fixtureyear == "" || fixtureyear =="null" || fixtureyear == "undefined")
{
	fixtureyear = new String("00:00:00").toString();
} 

homeoraway = Trim(new String(Request.Form("homeoraway_4")));
if (homeoraway == "" || homeoraway =="null" || homeoraway == "undefined")
{
	homeoraway = new String("H").toString();
} 

fixturenote = Trim(new String(Request.Form("fixturenote_4")));
if (fixturenote == "" || fixturenote =="null" || fixturenote == "undefined")
{
	fixturenote = new String(" ").toString();
} 

opponents = Trim(new String(Request.Form("opponents_4")));
if (opponents == "" || opponents =="null" || opponents == "undefined")
{
	opponents = new String("NONE").toString();
} 

pair1 = Trim(new String(Request.Form("pair1_4")));
if (pair1 == "" || pair1 =="null" || pair1 == "undefined")
{
	pair1 = new String("").toString();
} 

pair2 = Trim(new String(Request.Form("pair2_4")));
if (pair2 == "" || pair2 =="null" || pair2 == "undefined")
{
	pair2 = new String("").toString();
} 

hamptonresult = new Number(Request.Form("hamptonresult_4"));
if (hamptonresult == "" || hamptonresult =="null" || hamptonresult == "undefined")
{
	hamptonresult = 0;
} 

opponentresult = new Number(Request.Form("opponentresult_4"));
if (opponentresult == "" || opponentresult =="null" || opponentresult == "undefined")
{
	opponentresult = 0;
} 

debugWrite("fixture Id = ["+fixtureid+"], fixturedate = ["+fixturedate+"], venue = ["+homeoraway+"], opponents = ["+opponents+"], pair1 = ["+pair1+"], pair2 = ["+pair2+"], hamptonresult = ["+hamptonresult+"], opponentresult = ["+opponentresult+"]<br />");

fixturedetail = new FixtureObject(fixtureid,fixturedate,teamname,fixtureyear);
fixturedetail.setOpponents(opponents);
fixturedetail.setVenue(homeoraway);
fixturedetail.setPairs(pair1, pair2);
fixturedetail.setResult(hamptonresult, opponentresult)
fixtures.push(fixturedetail);


// Line 5

fixtureid = Trim(new String(Request.Form("fixtureid_5")));
if (fixtureid == "" || fixtureid =="null" || fixtureid == "undefined")
{
	fixtureid = new String("").toString();
} 

fixturedate = Trim(new String(Request.Form("fixturedate_5")));
if (fixturedate == "" || fixturedate =="null" || fixturedate == "undefined")
{
	fixturedate = new String("").toString();
} 

homeoraway = Trim(new String(Request.Form("homeoraway_5")));
if (homeoraway == "" || homeoraway =="null" || homeoraway == "undefined")
{
	homeoraway = new String("H").toString();
} 

fixturenote = Trim(new String(Request.Form("fixturenote_5")));
if (fixturenote == "" || fixturenote =="null" || fixturenote == "undefined")
{
	fixturenote = new String(" ").toString();
} 

opponents = Trim(new String(Request.Form("opponents_5")));
if (opponents == "" || opponents =="null" || opponents == "undefined")
{
	opponents = new String("NONE").toString();
} 

pair1 = Trim(new String(Request.Form("pair1_5")));
if (pair1 == "" || pair1 =="null" || pair1 == "undefined")
{
	pair1 = new String("").toString();
} 

pair2 = Trim(new String(Request.Form("pair2_5")));
if (pair2 == "" || pair2 =="null" || pair2 == "undefined")
{
	pair2 = new String("").toString();
} 

hamptonresult = new Number(Request.Form("hamptonresult_5"));
if (hamptonresult == "" || hamptonresult =="null" || hamptonresult == "undefined")
{
	hamptonresult = 0;
} 

opponentresult = new Number(Request.Form("opponentresult_5"));
if (opponentresult == "" || opponentresult =="null" || opponentresult == "undefined")
{
	opponentresult = 0;
} 

debugWrite("fixture Id = ["+fixtureid+"], fixturedate = ["+fixturedate+"], venue = ["+homeoraway+"], opponents = ["+opponents+"], pair1 = ["+pair1+"], pair2 = ["+pair2+"], hamptonresult = ["+hamptonresult+"], opponentresult = ["+opponentresult+"]<br />");

fixturedetail = new FixtureObject(fixtureid,fixturedate,teamname,fixtureyear);
fixturedetail.setOpponents(opponents);
fixturedetail.setVenue(homeoraway);
fixturedetail.setPairs(pair1, pair2);
fixturedetail.setResult(hamptonresult, opponentresult)
fixtures.push(fixturedetail);


// Line 6

fixtureid = Trim(new String(Request.Form("fixtureid_6")));
if (fixtureid == "" || fixtureid =="null" || fixtureid == "undefined")
{
	fixtureid = new String("").toString();
} 

fixturedate = Trim(new String(Request.Form("fixturedate_6")));
if (fixturedate == "" || fixturedate =="null" || fixturedate == "undefined")
{
	fixturedate = new String("").toString();
} 

homeoraway = Trim(new String(Request.Form("homeoraway_6")));
if (homeoraway == "" || homeoraway =="null" || homeoraway == "undefined")
{
	homeoraway = new String("H").toString();
} 

fixturenote = Trim(new String(Request.Form("fixturenote_6")));
if (fixturenote == "" || fixturenote =="null" || fixturenote == "undefined")
{
	fixturenote = new String(" ").toString();
} 

opponents = Trim(new String(Request.Form("opponents_6")));
if (opponents == "" || opponents =="null" || opponents == "undefined")
{
	opponents = new String("NONE").toString();
} 

pair1 = Trim(new String(Request.Form("pair1_6")));
if (pair1 == "" || pair1 =="null" || pair1 == "undefined")
{
	pair1 = new String("").toString();
} 

pair2 = Trim(new String(Request.Form("pair2_6")));
if (pair2 == "" || pair2 =="null" || pair2 == "undefined")
{
	pair2 = new String("").toString();
} 

hamptonresult = new Number(Request.Form("hamptonresult_6"));
if (hamptonresult == "" || hamptonresult =="null" || hamptonresult == "undefined")
{
	hamptonresult = 0;
} 

opponentresult = new Number(Request.Form("opponentresult_6"));
if (opponentresult == "" || opponentresult =="null" || opponentresult == "undefined")
{
	opponentresult = 0;
} 

debugWrite("fixture Id = ["+fixtureid+"], fixturedate = ["+fixturedate+"], venue = ["+homeoraway+"], opponents = ["+opponents+"], pair1 = ["+pair1+"], pair2 = ["+pair2+"], hamptonresult = ["+hamptonresult+"], opponentresult = ["+opponentresult+"]<br />");

fixturedetail = new FixtureObject(fixtureid,fixturedate,teamname,fixtureyear);
fixturedetail.setOpponents(opponents);
fixturedetail.setVenue(homeoraway);
fixturedetail.setPairs(pair1, pair2);
fixturedetail.setResult(hamptonresult, opponentresult)
fixtures.push(fixturedetail);


// Line 7

fixtureid = Trim(new String(Request.Form("fixtureid_7")));
if (fixtureid == "" || fixtureid =="null" || fixtureid == "undefined")
{
	fixtureid = new String("").toString();
} 

fixturedate = Trim(new String(Request.Form("fixturedate_7")));
if (fixturedate == "" || fixturedate =="null" || fixturedate == "undefined")
{
	fixturedate = new String("").toString();
} 

homeoraway = Trim(new String(Request.Form("homeoraway_7")));
if (homeoraway == "" || homeoraway =="null" || homeoraway == "undefined")
{
	homeoraway = new String("H").toString();
} 

fixturenote = Trim(new String(Request.Form("fixturenote_7")));
if (fixturenote == "" || fixturenote =="null" || fixturenote == "undefined")
{
	fixturenote = new String(" ").toString();
} 

opponents = Trim(new String(Request.Form("opponents_7")));
if (opponents == "" || opponents =="null" || opponents == "undefined")
{
	opponents = new String("NONE").toString();
} 

pair1 = Trim(new String(Request.Form("pair1_7")));
if (pair1 == "" || pair1 =="null" || pair1 == "undefined")
{
	pair1 = new String("").toString();
} 

pair2 = Trim(new String(Request.Form("pair2_7")));
if (pair2 == "" || pair2 =="null" || pair2 == "undefined")
{
	pair2 = new String("").toString();
} 

hamptonresult = new Number(Request.Form("hamptonresult_7"));
if (hamptonresult == "" || hamptonresult =="null" || hamptonresult == "undefined")
{
	hamptonresult = 0;
} 

opponentresult = new Number(Request.Form("opponentresult_7"));
if (opponentresult == "" || opponentresult =="null" || opponentresult == "undefined")
{
	opponentresult = 0;
} 

debugWrite("fixture Id = ["+fixtureid+"], fixturedate = ["+fixturedate+"], venue = ["+homeoraway+"], opponents = ["+opponents+"], pair1 = ["+pair1+"], pair2 = ["+pair2+"], hamptonresult = ["+hamptonresult+"], opponentresult = ["+opponentresult+"]<br />");

fixturedetail = new FixtureObject(fixtureid,fixturedate,teamname,fixtureyear);
fixturedetail.setOpponents(opponents);
fixturedetail.setVenue(homeoraway);
fixturedetail.setPairs(pair1, pair2);
fixturedetail.setResult(hamptonresult, opponentresult)
fixtures.push(fixturedetail);


// Line 8

fixtureid = Trim(new String(Request.Form("fixtureid_8")));
if (fixtureid == "" || fixtureid =="null" || fixtureid == "undefined")
{
	fixtureid = new String("").toString();
} 

fixturedate = Trim(new String(Request.Form("fixturedate_8")));
if (fixturedate == "" || fixturedate =="null" || fixturedate == "undefined")
{
	fixturedate = new String("").toString();
} 

homeoraway = Trim(new String(Request.Form("homeoraway_8")));
if (homeoraway == "" || homeoraway =="null" || homeoraway == "undefined")
{
	homeoraway = new String("H").toString();
} 

fixturenote = Trim(new String(Request.Form("fixturenote_8")));
if (fixturenote == "" || fixturenote =="null" || fixturenote == "undefined")
{
	fixturenote = new String(" ").toString();
} 

opponents = Trim(new String(Request.Form("opponents_8")));
if (opponents == "" || opponents =="null" || opponents == "undefined")
{
	opponents = new String("NONE").toString();
} 

pair1 = Trim(new String(Request.Form("pair1_8")));
if (pair1 == "" || pair1 =="null" || pair1 == "undefined")
{
	pair1 = new String("").toString();
} 

pair2 = Trim(new String(Request.Form("pair2_8")));
if (pair2 == "" || pair2 =="null" || pair2 == "undefined")
{
	pair2 = new String("").toString();
} 

hamptonresult = new Number(Request.Form("hamptonresult_8"));
if (hamptonresult == "" || hamptonresult =="null" || hamptonresult == "undefined")
{
	hamptonresult = 0;
} 

opponentresult = new Number(Request.Form("opponentresult_8"));
if (opponentresult == "" || opponentresult =="null" || opponentresult == "undefined")
{
	opponentresult = 0;
} 

debugWrite("fixture Id = ["+fixtureid+"], fixturedate = ["+fixturedate+"], venue = ["+homeoraway+"], opponents = ["+opponents+"], pair1 = ["+pair1+"], pair2 = ["+pair2+"], hamptonresult = ["+hamptonresult+"], opponentresult = ["+opponentresult+"]<br />");

fixturedetail = new FixtureObject(fixtureid,fixturedate,teamname,fixtureyear);
fixturedetail.setOpponents(opponents);
fixturedetail.setVenue(homeoraway);
fixturedetail.setPairs(pair1, pair2);
fixturedetail.setResult(hamptonresult, opponentresult)
fixtures.push(fixturedetail);


// Line 9

fixtureid = Trim(new String(Request.Form("fixtureid_9")));
if (fixtureid == "" || fixtureid =="null" || fixtureid == "undefined")
{
	fixtureid = new String("").toString();
} 

fixturedate = Trim(new String(Request.Form("fixturedate_9")));
if (fixturedate == "" || fixturedate =="null" || fixturedate == "undefined")
{
	fixturedate = new String("").toString();
} 

homeoraway = Trim(new String(Request.Form("homeoraway_9")));
if (homeoraway == "" || homeoraway =="null" || homeoraway == "undefined")
{
	homeoraway = new String("H").toString();
} 

fixturenote = Trim(new String(Request.Form("fixturenote_9")));
if (fixturenote == "" || fixturenote =="null" || fixturenote == "undefined")
{
	fixturenote = new String(" ").toString();
} 

opponents = Trim(new String(Request.Form("opponents_9")));
if (opponents == "" || opponents =="null" || opponents == "undefined")
{
	opponents = new String("NONE").toString();
} 

pair1 = Trim(new String(Request.Form("pair1_9")));
if (pair1 == "" || pair1 =="null" || pair1 == "undefined")
{
	pair1 = new String("").toString();
} 

pair2 = Trim(new String(Request.Form("pair2_9")));
if (pair2 == "" || pair2 =="null" || pair2 == "undefined")
{
	pair2 = new String("").toString();
} 

hamptonresult = new Number(Request.Form("hamptonresult_9"));
if (hamptonresult == "" || hamptonresult =="null" || hamptonresult == "undefined")
{
	hamptonresult = 0;
} 

opponentresult = new Number(Request.Form("opponentresult_9"));
if (opponentresult == "" || opponentresult =="null" || opponentresult == "undefined")
{
	opponentresult = 0;
} 

debugWrite("fixture Id = ["+fixtureid+"], fixturedate = ["+fixturedate+"], venue = ["+homeoraway+"], opponents = ["+opponents+"], pair1 = ["+pair1+"], pair2 = ["+pair2+"], hamptonresult = ["+hamptonresult+"], opponentresult = ["+opponentresult+"]<br />");

fixturedetail = new FixtureObject(fixtureid,fixturedate,teamname,fixtureyear);
fixturedetail.setOpponents(opponents);
fixturedetail.setVenue(homeoraway);
fixturedetail.setPairs(pair1, pair2);
fixturedetail.setResult(hamptonresult, opponentresult)
fixtures.push(fixturedetail);

// Line 10

fixtureid = Trim(new String(Request.Form("fixtureid_10")));
if (fixtureid == "" || fixtureid =="null" || fixtureid == "undefined")
{
	fixtureid = new String("").toString();
} 

fixturedate = Trim(new String(Request.Form("fixturedate_10")));
if (fixturedate == "" || fixturedate =="null" || fixturedate == "undefined")
{
	fixturedate = new String("").toString();
} 

homeoraway = Trim(new String(Request.Form("homeoraway_10")));
if (homeoraway == "" || homeoraway =="null" || homeoraway == "undefined")
{
	homeoraway = new String("H").toString();
} 

fixturenote = Trim(new String(Request.Form("fixturenote_10")));
if (fixturenote == "" || fixturenote =="null" || fixturenote == "undefined")
{
	fixturenote = new String(" ").toString();
} 

opponents = Trim(new String(Request.Form("opponents_10")));
if (opponents == "" || opponents =="null" || opponents == "undefined")
{
	opponents = new String("NONE").toString();
} 

pair1 = Trim(new String(Request.Form("pair1_10")));
if (pair1 == "" || pair1 =="null" || pair1 == "undefined")
{
	pair1 = new String("").toString();
} 

pair2 = Trim(new String(Request.Form("pair2_10")));
if (pair2 == "" || pair2 =="null" || pair2 == "undefined")
{
	pair2 = new String("").toString();
} 

hamptonresult = new Number(Request.Form("hamptonresult_10"));
if (hamptonresult == "" || hamptonresult =="null" || hamptonresult == "undefined")
{
	hamptonresult = 0;
} 

opponentresult = new Number(Request.Form("opponentresult_10"));
if (opponentresult == "" || opponentresult =="null" || opponentresult == "undefined")
{
	opponentresult = 0;
} 

debugWrite("fixture Id = ["+fixtureid+"], fixturedate = ["+fixturedate+"], venue = ["+homeoraway+"], opponents = ["+opponents+"], pair1 = ["+pair1+"], pair2 = ["+pair2+"], hamptonresult = ["+hamptonresult+"], opponentresult = ["+opponentresult+"]<br />");

fixturedetail = new FixtureObject(fixtureid,fixturedate,teamname,fixtureyear);
fixturedetail.setOpponents(opponents);
fixturedetail.setVenue(homeoraway);
fixturedetail.setPairs(pair1, pair2);
fixturedetail.setResult(hamptonresult, opponentresult)
fixtures.push(fixturedetail);


// Line 11

fixtureid = Trim(new String(Request.Form("fixtureid_11")));
if (fixtureid == "" || fixtureid =="null" || fixtureid == "undefined")
{
	fixtureid = new String("").toString();
} 

fixturedate = Trim(new String(Request.Form("fixturedate_11")));
if (fixturedate == "" || fixturedate =="null" || fixturedate == "undefined")
{
	fixturedate = new String("").toString();
} 

homeoraway = Trim(new String(Request.Form("homeoraway_11")));
if (homeoraway == "" || homeoraway =="null" || homeoraway == "undefined")
{
	homeoraway = new String("H").toString();
} 

fixturenote = Trim(new String(Request.Form("fixturenote_11")));
if (fixturenote == "" || fixturenote =="null" || fixturenote == "undefined")
{
	fixturenote = new String(" ").toString();
} 

opponents = Trim(new String(Request.Form("opponents_11")));
if (opponents == "" || opponents =="null" || opponents == "undefined")
{
	opponents = new String("NONE").toString();
} 

pair1 = Trim(new String(Request.Form("pair1_11")));
if (pair1 == "" || pair1 =="null" || pair1 == "undefined")
{
	pair1 = new String("").toString();
} 

pair2 = Trim(new String(Request.Form("pair2_11")));
if (pair2 == "" || pair2 =="null" || pair2 == "undefined")
{
	pair2 = new String("").toString();
} 

hamptonresult = new Number(Request.Form("hamptonresult_11"));
if (hamptonresult == "" || hamptonresult =="null" || hamptonresult == "undefined")
{
	hamptonresult = 0;
} 

opponentresult = new Number(Request.Form("opponentresult_11"));
if (opponentresult == "" || opponentresult =="null" || opponentresult == "undefined")
{
	opponentresult = 0;
} 

debugWrite("fixture Id = ["+fixtureid+"], fixturedate = ["+fixturedate+"], venue = ["+homeoraway+"], opponents = ["+opponents+"], pair1 = ["+pair1+"], pair2 = ["+pair2+"], hamptonresult = ["+hamptonresult+"], opponentresult = ["+opponentresult+"]<br />");

fixturedetail = new FixtureObject(fixtureid,fixturedate,teamname,fixtureyear);
fixturedetail.setOpponents(opponents);
fixturedetail.setVenue(homeoraway);
fixturedetail.setPairs(pair1, pair2);
fixturedetail.setResult(hamptonresult, opponentresult)
fixtures.push(fixturedetail);


// Line 12

fixtureid = Trim(new String(Request.Form("fixtureid_12")));
if (fixtureid == "" || fixtureid =="null" || fixtureid == "undefined")
{
	fixtureid = new String("").toString();
} 

fixturedate = Trim(new String(Request.Form("fixturedate_12")));
if (fixturedate == "" || fixturedate =="null" || fixturedate == "undefined")
{
	fixturedate = new String("").toString();
} 

homeoraway = Trim(new String(Request.Form("homeoraway_12")));
if (homeoraway == "" || homeoraway =="null" || homeoraway == "undefined")
{
	homeoraway = new String("H").toString();
} 

fixturenote = Trim(new String(Request.Form("fixturenote_12")));
if (fixturenote == "" || fixturenote =="null" || fixturenote == "undefined")
{
	fixturenote = new String(" ").toString();
} 

opponents = Trim(new String(Request.Form("opponents_12")));
if (opponents == "" || opponents =="null" || opponents == "undefined")
{
	opponents = new String("NONE").toString();
} 

pair1 = Trim(new String(Request.Form("pair1_12")));
if (pair1 == "" || pair1 =="null" || pair1 == "undefined")
{
	pair1 = new String("").toString();
} 

pair2 = Trim(new String(Request.Form("pair2_12")));
if (pair2 == "" || pair2 =="null" || pair2 == "undefined")
{
	pair2 = new String("").toString();
} 

hamptonresult = new Number(Request.Form("hamptonresult_12"));
if (hamptonresult == "" || hamptonresult =="null" || hamptonresult == "undefined")
{
	hamptonresult = 0;
} 

opponentresult = new Number(Request.Form("opponentresult_12"));
if (opponentresult == "" || opponentresult =="null" || opponentresult == "undefined")
{
	opponentresult = 0;
} 

debugWrite("fixture Id = ["+fixtureid+"], fixturedate = ["+fixturedate+"], venue = ["+homeoraway+"], opponents = ["+opponents+"], pair1 = ["+pair1+"], pair2 = ["+pair2+"], hamptonresult = ["+hamptonresult+"], opponentresult = ["+opponentresult+"]<br />");

fixturedetail = new FixtureObject(fixtureid,fixturedate,teamname,fixtureyear);
fixturedetail.setOpponents(opponents);
fixturedetail.setVenue(homeoraway);
fixturedetail.setPairs(pair1, pair2);
fixturedetail.setResult(hamptonresult, opponentresult)
fixtures.push(fixturedetail);


// Line 13

fixtureid = Trim(new String(Request.Form("fixtureid_13")));
if (fixtureid == "" || fixtureid =="null" || fixtureid == "undefined")
{
	fixtureid = new String("").toString();
} 

fixturedate = Trim(new String(Request.Form("fixturedate_13")));
if (fixturedate == "" || fixturedate =="null" || fixturedate == "undefined")
{
	fixturedate = new String("").toString();
} 

homeoraway = Trim(new String(Request.Form("homeoraway_13")));
if (homeoraway == "" || homeoraway =="null" || homeoraway == "undefined")
{
	homeoraway = new String("H").toString();
} 

fixturenote = Trim(new String(Request.Form("fixturenote_13")));
if (fixturenote == "" || fixturenote =="null" || fixturenote == "undefined")
{
	fixturenote = new String(" ").toString();
} 

opponents = Trim(new String(Request.Form("opponents_13")));
if (opponents == "" || opponents =="null" || opponents == "undefined")
{
	opponents = new String("NONE").toString();
} 

pair1 = Trim(new String(Request.Form("pair1_13")));
if (pair1 == "" || pair1 =="null" || pair1 == "undefined")
{
	pair1 = new String("").toString();
} 

pair2 = Trim(new String(Request.Form("pair2_13")));
if (pair2 == "" || pair2 =="null" || pair2 == "undefined")
{
	pair2 = new String("").toString();
} 

hamptonresult = new Number(Request.Form("hamptonresult_13"));
if (hamptonresult == "" || hamptonresult =="null" || hamptonresult == "undefined")
{
	hamptonresult = 0;
} 

opponentresult = new Number(Request.Form("opponentresult_13"));
if (opponentresult == "" || opponentresult =="null" || opponentresult == "undefined")
{
	opponentresult = 0;
} 

debugWrite("fixture Id = ["+fixtureid+"], fixturedate = ["+fixturedate+"], venue = ["+homeoraway+"], opponents = ["+opponents+"], pair1 = ["+pair1+"], pair2 = ["+pair2+"], hamptonresult = ["+hamptonresult+"], opponentresult = ["+opponentresult+"]<br />");

fixturedetail = new FixtureObject(fixtureid,fixturedate,teamname,fixtureyear);
fixturedetail.setOpponents(opponents);
fixturedetail.setVenue(homeoraway);
fixturedetail.setPairs(pair1, pair2);
fixturedetail.setResult(hamptonresult, opponentresult)
fixtures.push(fixturedetail);


// Line 14

fixtureid = Trim(new String(Request.Form("fixtureid_14")));
if (fixtureid == "" || fixtureid =="null" || fixtureid == "undefined")
{
	fixtureid = new String("").toString();
} 

fixturedate = Trim(new String(Request.Form("fixturedate_14")));
if (fixturedate == "" || fixturedate =="null" || fixturedate == "undefined")
{
	fixturedate = new String("").toString();
} 

homeoraway = Trim(new String(Request.Form("homeoraway_14")));
if (homeoraway == "" || homeoraway =="null" || homeoraway == "undefined")
{
	homeoraway = new String("H").toString();
} 

fixturenote = Trim(new String(Request.Form("fixturenote_14")));
if (fixturenote == "" || fixturenote =="null" || fixturenote == "undefined")
{
	fixturenote = new String(" ").toString();
} 

opponents = Trim(new String(Request.Form("opponents_14")));
if (opponents == "" || opponents =="null" || opponents == "undefined")
{
	opponents = new String("NONE").toString();
} 

pair1 = Trim(new String(Request.Form("pair1_14")));
if (pair1 == "" || pair1 =="null" || pair1 == "undefined")
{
	pair1 = new String("").toString();
} 

pair2 = Trim(new String(Request.Form("pair2_14")));
if (pair2 == "" || pair2 =="null" || pair2 == "undefined")
{
	pair2 = new String("").toString();
} 

hamptonresult = new Number(Request.Form("hamptonresult_14"));
if (hamptonresult == "" || hamptonresult =="null" || hamptonresult == "undefined")
{
	hamptonresult = 0;
} 

opponentresult = new Number(Request.Form("opponentresult_14"));
if (opponentresult == "" || opponentresult =="null" || opponentresult == "undefined")
{
	opponentresult = 0;
} 

debugWrite("fixture Id = ["+fixtureid+"], fixturedate = ["+fixturedate+"], venue = ["+homeoraway+"], opponents = ["+opponents+"], pair1 = ["+pair1+"], pair2 = ["+pair2+"], hamptonresult = ["+hamptonresult+"], opponentresult = ["+opponentresult+"]<br />");

fixturedetail = new FixtureObject(fixtureid,fixturedate,teamname,fixtureyear);
fixturedetail.setOpponents(opponents);
fixturedetail.setVenue(homeoraway);
fixturedetail.setPairs(pair1, pair2);
fixturedetail.setResult(hamptonresult, opponentresult)
fixtures.push(fixturedetail);


//      End of the potential fixtures

debugWrite("fixtures = "+JSON.stringify(fixtures)+"<br /><hr />");

// Update fixture results from POST'ed data
// Loop through all the fixtures, update via fixtureid as unique row identifier

for (var j=0; j<fixtures.length; j++) {

	fixturedetail = fixtures[j];

	if (! (fixturedetail.opponents == "NONE")) {   // ignore if no opponents

		SQLend = new String(" WHERE fixtureid = "+fixturedetail.fixtureid).toString();

		SQLmiddle = new String("SET ").toString();
		SQLmiddle += " fixtureyear="+fixturedetail.fixtureyear+",";
		SQLmiddle += " homeoraway='"+fixturedetail.homeoraway+"',";
		SQLmiddle += " fixturenote='"+fixturedetail.fixturenote+"',";
		SQLmiddle += " teamname='"+fixturedetail.teamname+"', ";
		SQLmiddle += " opponents='"+fixturedetail.opponents+"', ";
		SQLmiddle += " pair1='"+fixturedetail.pair1+"', ";
		SQLmiddle += " pair2='"+fixturedetail.pair2+"', ";
		SQLmiddle += " hamptonresult="+fixturedetail.hamptonresult+", ";
		SQLmiddle += " opponentresult="+fixturedetail.opponentresult+", ";

		// Now do date fields. If null dont insert them as part of the update clause
		//  Access doesnt like setting date fields to ''

		if (! (fixturedetail.fixturedate == ""))
			SQLmiddle += " fixturedate='"+fixturedetail.fixturedate+"' ";
		else
			SQLmiddle += " fixturedate=null ";

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
	Response.Redirect("./fixtureresults.html#/");
}

Response.End();


%>


