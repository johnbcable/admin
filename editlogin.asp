<%@language="JScript" CODEPAGE="65001"%>
<!--#include file="functions.asp" -->
<!--#include file="strings.asp" -->
<!--#include file="resultobj.asp" -->
<%
// First variables that need to be set for each page
var strtime, strdate;
var debugging = false; // set to true if debug output reqd
var clubname = new String("Hampton-In-Arden Sports Club");
var pagetitle = new String("Membership Renewals Sign-In");
// Now for any variables local to this page
var v_membername = new String("");
var v_memberid;
var v_acclevel;
var v_uniqueref;
var v_onlinebookingid, v_onlinebookingpin;
var dbconnect=Application("hamptonsportsdb"); 
var ConnObj, RS, sqlString, SQL1, SQL2, SQL3, SQL4;
var badSignIn=true;
var errMess;
var myusername, myaddress;
var refurl;
var onwardurl;
var frmUserId, frmPassword;
var mySignIn = new Object();
var results = new ResultObject();

// var debugging = current_debug_status();
// Set up default greeting strings
strdate = datestring();
strtime = timestring();
v_memberid=new String("");
v_membername=new String("");
v_acclevel=0;
frmUserId = new String("");
frmPassword = new String("");

// End of page start up coding
refurl = new String(Request.Form("refurl")).toString();
if (refurl == "" || refurl=="null" || refurl == "undefined")
	refurl = new String("").toString();
onwardurl = new String(refurl).toString();
if (onwardurl == "")
	onwardurl = new String("members.html").toString();

frmUserId = new String(Request.Form("frmUserId")).toString();
if (frmUserId == "" || frmUserId=="null" || frmUserId == "undefined")
	frmUserId = new String("").toString();

frmPassword = new String(Request.Form("frmPassword")).toString();
if (frmPassword == "" || frmPassword=="null" || frmPassword == "undefined")
	frmPassword = new String("").toString();

myusername = new String(Request.ServerVariables("REMOTE_USER")).toString();
myaddress = new String(Request.ServerVariables("REMOTE_ADDR")).toString();
if (myusername == "" || myusername=="null" || myusername == "undefined")
	myusername = new String(myaddress).toString();
errMess = new String("")
displaydate = dateasstring(Date());
// End of page start up coding

// Now try and sign me in

ConnObj = Server.CreateObject("ADODB.connection");
ConnObj.Open(dbconnect);

// Clear out any prior sessions for this device
sqlString=new String("delete from user_sessions where remoteaddress = '"+myaddress+"' or remoteuser = '"+myusername+"'").toString();
RS=ConnObj.Execute(sqlString);

// Now try login check - only dojng onlinebookingid and onlinebookingpin
SQL1 = new String("Select uniqueref, onlinebookingid, onlinebookingpin, memberid, forename1, webaccess from members where onlinebookingid = " + frmUserId + " and onlinebookingpin = '" + frmPassword + "'");

RS = ConnObj.Execute(SQL1);
// Now check if sign in was OK 
if (! RS.EOF)
{
	badSignIn=false;
	mySignIn.uniqueref = new String(RS("uniqueref")).toString();
	mySignIn.onlinebookingid = new String("onlinebookingid").toString();
	mySignIn.onlinebookingpin = new String("onlinebookingpin").toString();
	mySignIn.memberid = new String(RS("memberid")).toString();
	mySignIn.forename = new String(RS("forename1")).toString();
	mySignIn.acclevel = new String(RS("webaccess")).toString();
}
RS.Close();

if (debugging)
{
%>
		<p>
			<br />
			frmUserId was     [<%= frmUserId %>]<br />
			frmPassword was   [<%= frmPassword %>]<br />
			SQL enquiry was   [<%= SQL1 %>]<br />
			Referring URL was [<%= refurl %>]<br />
			Onward URL was    [<%= onwardurl %>]<br />
			v_uniqueref was   [<%= mySignIn.uniqueref %>]<br />
			v_memberid was    [<%= mySignIn.memberid %>]<br />
			v_membername was  [<%= mySignIn.forename %>]<br />
			v_acclevel was    [<%= mySignIn.acclevel %>]<br />
			v_onlinebookingid was    [<%= mySignIn.onlinebookingid %>]<br />
			v_onlinebookingpin was    [<%= mySignIn.onlinebookingpin %>]<br />
		</p>
<%
}

// Make sure we have a default onward URl if one not come in from form
if (onwardurl == "")
	onwardurl = new String("/renewprofile.html").toString();

if (badSignIn)
{
	// Bad user ID / PIN combination detected
	// Add failure indicators to result object

	errMess+=new String("Incorrect user name and/or password supplied.");
	results.result = false;  // Failed login
	results.sql = new String(SQL1).toString();
	results.addErrorMessage(errmess);
} 
else 
{
	// Here is where we need to save the signin JS object to localstorage

	results.data.push(mySignIn);  // JS signin object

	// localStorage.setItem("tennissignin", mySignIn);

	setUser(v_uniqueref,v_memberid,v_membername,v_acclevel);  // Session record - not permanent

	RS=null;
	ConnObj.Close();
	ConnObj=null;

	// Go on to your onward URL if you signed-in OK
	if (! debugging)
		Response.Redirect(onwardurl); // redirect to member services area	
}

%>
