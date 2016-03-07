
<%@language="JScript" CODEPAGE="65001"%>
<!--#include file="functions.asp" -->
<!--#include file="strings.asp" -->
<!--#include file="emailfuncs.asp" -->
<%

// Set up cache control on this page
Response.AddHeader("Cache-Control", "no-cache,no-store,must-revalidate");
Response.AddHeader("Pragma", "no-cache");
Response.AddHeader("Expires", 0);

var strtime, strdate, title2;
var clubname = new String("Hampton-In-Arden Sports Club");
var pagetitle = new String("Updating Profile Details");
// Now for any variables local to this page
var RS, Conn, SQL1, SQL2;
var dbconnect=Application("hamptonsportsdb"); 
var memberObj = new Object();
var SQLstart, SQLmiddle, SQLend;
var resultObj = new Object();
var uniqueref;

// Set up default greeting strings
strdate = datestring();
strtime = timestring();
//
var debugging = false;  // Production = false

// defaults for variables
memberObj.surname = new String("").toString();
memberObj.forename1 = new String("").toString();
memberObj.dob = new String("").toString();
memberObj.gender = new String("").toString();
memberObj.britishtennisno = new String("UNKNOWN").toString();
memberObj.address1 = new String("").toString();
memberObj.address2 = new String("").toString();
memberObj.address3 = new String("").toString();
memberObj.address4 = new String("").toString();
memberObj.postcode = new String("").toString();
memberObj.homephone = new String("").toString();
memberObj.mobilephone = new String("").toString();
memberObj.email = new String("").toString();
memberObj.emergencyname = new String("").toString();
memberObj.emergencyphone = new String("").toString();
memberObj.medicalinfo = new String("").toString();

// Retrieve POST'ed data

uniqueref = Request.Form("uniqueref");

// memberObj = getMember(uniqueref);

memberObj.surname = Trim(new String(Request.Form("surname")));
if (memberObj.surname == "" || memberObj.surname =="null" || memberObj.surname == "undefined")
{
	memberObj.surname = new String("").toString();
} 

memberObj.forename1 = Trim(new String(Request.Form("forename1")));
if (memberObj.forename1 == "" || memberObj.forename1 =="null" || memberObj.forename1 == "undefined")
{
	memberObj.forename1 = new String("").toString();
} 

memberObj.dob = Trim(new String(Request.Form("dob")));
if (memberObj.dob == "" || memberObj.dob =="null" || memberObj.dob == "undefined")
{
	memberObj.dob = new String("").toString();
} 

memberObj.gender = Trim(new String(Request.Form("gender")));
if (memberObj.gender == "" || memberObj.gender =="null" || memberObj.gender == "undefined")
{
	memberObj.gender = new String("").toString();
} 

memberObj.britishtennisno = Trim(new String(Request.Form("britishtennisno")));
if (memberObj.britishtennisno == "" || memberObj.britishtennisno =="null" || memberObj.britishtennisno == "undefined")
{
	memberObj.britishtennisno = new String("UNKNOWN").toString();
} 

memberObj.address1 = Trim(new String(Request.Form("address1")));
if (memberObj.address1 == "" || memberObj.address1 =="null" || memberObj.address1 == "undefined")
{
	memberObj.address1 = new String("").toString();
} 

memberObj.address2 = Trim(new String(Request.Form("address2")));
if (memberObj.address2 == "" || memberObj.address2 =="null" || memberObj.address2 == "undefined")
{
	memberObj.address2 = new String("").toString();
} 

memberObj.address3 = Trim(new String(Request.Form("address3")));
if (memberObj.address3 == "" || memberObj.address3 =="null" || memberObj.address3 == "undefined")
{
	memberObj.address3 = new String("").toString();
} 

memberObj.address4 = Trim(new String(Request.Form("address4")));
if (memberObj.address4 == "" || memberObj.address4 =="null" || memberObj.address4 == "undefined")
{
	memberObj.address4 = new String("").toString();
} 

memberObj.postcode = Trim(new String(Request.Form("postcode")));
if (memberObj.postcode == "" || memberObj.postcode =="null" || memberObj.postcode == "undefined")
{
	memberObj.postcode = new String("").toString();
} 

memberObj.email = Trim(new String(Request.Form("email")));
if (memberObj.email == "" || memberObj.email =="null" || memberObj.email == "undefined")
{
	memberObj.email = new String("").toString();
} 

memberObj.homephone = Trim(new String(Request.Form("homephone")));
if (memberObj.homephone == "" || memberObj.homephone =="null" || memberObj.homephone == "undefined")
{
	memberObj.homephone = new String("").toString();
} 

memberObj.mobilephone = Trim(new String(Request.Form("mobilephone")));
if (memberObj.mobilephone == "" || memberObj.mobilephone =="null" || memberObj.mobilephone == "undefined")
{
	memberObj.mobilephone = new String("").toString();
} 

memberObj.emergencyname = Trim(new String(Request.Form("emergencyname")));
if (memberObj.emergencyname == "" || memberObj.emergencyname =="null" || memberObj.emergencyname == "undefined")
{
	memberObj.emergencyname = new String("").toString();
} 

memberObj.emergencyphone = Trim(new String(Request.Form("emergencyphone")));
if (memberObj.emergencyphone == "" || memberObj.emergencyphone =="null" || memberObj.emergencyphone == "undefined")
{
	memberObj.emergencyphone = new String("").toString();
} 

memberObj.medicalinfo = Trim(new String(Request.Form("medicalinfo")));
if (memberObj.medicalinfo == "" || memberObj.medicalinfo =="null" || memberObj.medicalinfo == "undefined")
{
	memberObj.medicalinfo = new String("").toString();
} 

// Update members details from POST'ed data

Conn = Server.CreateObject("ADODB.Connection");
RS = Server.CreateObject("ADODB.RecordSet");
Conn.Open(dbconnect);
SQLstart = new String("UPDATE members ")
SQLend = new String(" WHERE uniqueref="+uniqueref).toString();
SQLmiddle = new String("SET ").toString();
SQLmiddle += " surname='"+memberObj.surname+"',";
SQLmiddle += " forename1='"+memberObj.forename1+"',";
SQLmiddle += " gender='"+memberObj.gender+"', ";
SQLmiddle += " address1='"+memberObj.address1+"',";
SQLmiddle += " address2='"+memberObj.address2+"',";
SQLmiddle += " address3='"+memberObj.address3+"',";
SQLmiddle += " address4='"+memberObj.address4+"',";
SQLmiddle += " postcode='"+memberObj.postcode+"',";
SQLmiddle += " homephone='"+memberObj.homephone+"',";
SQLmiddle += " mobilephone='"+memberObj.mobilephone+"',";
SQLmiddle += " email='"+memberObj.email+"',";
SQLmiddle += " britishtennisno='"+memberObj.britishtennisno+"',";
SQLmiddle += " emergencyname='"+memberObj.emergencyname+"',";
SQLmiddle += " emergencyphone='"+memberObj.emergencyphone+"',";
SQLmiddle += " medicalinfo='"+memberObj.medicalinfo+"',";

// Now do date fields. If null dont insert them as part of the update clause
//  Access doesnt like setting date fields to ''
// N.B.  If dob is the last column, dont put trailing comma on
if (! (memberObj.dob == "")) {
	SQLmiddle += " dob='"+memberObj.dob+"' ";
} else {
	SQLmiddle += " dob=null ";
}

// Default values into result object
resultObj.result = true;
resultObj.errno = 0;
resultObj.description = new String("").toString();

SQL1 = new String(SQLstart+SQLmiddle+SQLend).toString();
resultObj.sql = new String(SQL1).toString();

try {
	RS = Conn.Execute(SQL1);
}
catch(e) {
	resultObj.result = false;
	resultObj.errno = (e.number & 0xFFFF);
	resultObj.description += e.description;
	resultObj.sql = new String(SQLstart+SQLmiddle+SQLend).toString();
}

if (debugging) {
	var theresult = resultObj.result ? "TRUE" : "FALSE";
	Response.Write("Information from the DB update:<br />");
	Response.Write("Result: "+ theresult+"<br />");
	Response.Write("errno: "+resultObj.errno+"<br />");
	Response.Write("Description: "+resultObj.description+"<br />");
	Response.Write("SQL used: <br />"+resultObj.sql+"<br /><br />");

	Response.Write("If not debugging will redirect to /renewpreferences.html<br /><br />");
	Response.Write("<h3>Values POST-ed</h3>");
	Response.Write("<table>");
	Response.Write("<tr><td>Surname</td><td>"+memberObj.surname+"</td></tr>");
	Response.Write("<tr><td>Forename</td><td>"+memberObj.forename1+"</td></tr>");
	Response.Write("<tr><td>DOB</td><td>"+memberObj.dob+"</td></tr>");
	Response.Write("<tr><td>Gender</td><td>"+memberObj.gender+"</td></tr>");
	Response.Write("<tr><td>British Tennis Number</td><td>"+memberObj.britishtennisno+"</td></tr>");
	Response.Write("<tr><td>Address</td><td>"+memberObj.address1+"</td></tr>");
	Response.Write("<tr><td>&nbsp;</td><td>"+memberObj.address2+"</td></tr>");
	Response.Write("<tr><td>&nbsp;</td><td>"+memberObj.address3+"</td></tr>");
	Response.Write("<tr><td>&nbsp;</td><td>"+memberObj.address4+"</td></tr>");
	Response.Write("<tr><td>Postcode</td><td>"+memberObj.postcode+"</td></tr>");
	Response.Write("<tr><td>Home phone</td><td>"+memberObj.homephone+"</td></tr>");
	Response.Write("<tr><td>Mobile phone</td><td>"+memberObj.mobilephone+"</td></tr>");
	Response.Write("<tr><td>Email</td><td>"+memberObj.email+"</td></tr>");
	Response.Write("<tr><td>Emergency Contact Name</td><td>"+memberObj.emergencyname+"</td></tr>");
	Response.Write("<tr><td>Emergency Contact Phone</td><td>"+memberObj.emergencyphone+"</td></tr>");
	Response.Write("<tr><td>Medical Info</td><td>"+memberObj.medicalinfo+"</td></tr>");

	Response.Write("</table>");

	Response.End();

}

// On completion, redirect to mypreferences.html

Response.Redirect("./renewpreferences.html")
Response.End();


%>

