<%@language="JScript" CODEPAGE="65001"%>
<!--#include file="functions.asp" -->
<!--#include file="strings.asp" -->
<!--#include file="emailfuncs.asp" -->
<%
var strtime, strdate, title2;
var clubname = new String("Hampton-In-Arden Sports Club");
var pagetitle = new String("Updating Details of Event");
// Now for any variables local to this page
var m_eventdate, m_eventtime, m_eventnote, m_eventtype, m_eventid, m_eventreport;
var vdate, m_eventyear;   // m_eventyear is always calculated from m_eventdate
var ConnObj;
var RS,RS2,RS3;
var SQLStmt, SQL2, SQL3;
var kount;
var memberknt;
var dbconnect=Application("hamptonsportsdb");
var debugging=false;
var updating=true;
// Set up default greeting strings
strdate = datestring();
strtime = timestring();
// End of page start up coding
// Initialise update variables
m_eventid = Trim(new String(Request.QueryString("eventid")));
m_eventdate = Trim(new String(Request.QueryString("eventdate")));
m_eventtime = Trim(new String(Request.QueryString("eventtime")));
m_eventnote = Trim(new String(Request.QueryString("eventnote")));
m_eventtype = Trim(new String(Request.QueryString("eventtype")));
m_eventreport = Trim(new String(Request.QueryString("eventreport")));
// Flag if this is a new member insertion 
newone = (m_eventid == "-1");
// reset if null
if (m_eventdate=="undefined" || m_eventdate == "null" || m_eventdate == "")
{
	today = new Date();
	m_eventdate = new String(ddmmyyyy(today)).toString();
}
if (m_eventtime=="undefined" || m_eventtime == "null" || m_eventtime == "")
	m_eventtime = new String("00:00:00").toString();
else
	m_eventtime = new String(m_eventtime.substr(0,2)+":"+m_eventtime.substr(2)+":00").toString();
if (m_eventnote=="undefined" || m_eventnote == "null" || m_eventnote == "")
	m_eventnote = new String("No event title supplied").toString();
if (m_eventtype=="undefined" || m_eventtype == "null" || m_eventtype == "")
	m_eventtype = new String("EVENT").toString();
if (m_eventreport=="undefined" || m_eventreport == "null" || m_eventreport == "")
	m_eventreport = new String("").toString();

// Calculate event year from event date
vdate = new Date(m_eventdate.substr(6),m_eventdate.substr(3,2),m_eventdate.substr(0,2));
m_eventyear = new String(vdate.getFullYear()).toString();

// Do DB update
ConnObj=Server.CreateObject("ADODB.Connection");
RS=Server.CreateObject("ADODB.Recordset");
ConnObj.Open(dbconnect);
if (newone)
{
	// eventdate. eventtime, eventyear, eventtype, eventnote, eventid, eventreport
	// INSERT into EVENTS (eventdate,eventtime,eventyear,eventtype,eventnote,eventreport) values('13/01/2008','14:00:00',2008,'EVENT','Junior Tournament 2007 (Completion)','')
	SQL1 = new String("INSERT into EVENTS (eventdate,eventtime,eventyear,eventtype,eventnote,eventreport) VALUES('"+m_eventdate+"','"+m_eventtime+"',"+m_eventyear+",'"+m_eventtype+"','"+m_eventnote+"','"+m_eventreport+"')").toString();
}
else
{
	SQL1 = new String("UPDATE EVENTS set eventdate='"+m_eventdate+"',eventtime='"+m_eventtime+"',eventyear="+m_eventyear+",eventtype='"+m_eventtype+"',eventnote='"+m_eventnote+"',eventreport='"+m_eventreport+"' WHERE eventid = "+m_eventid).toString();
}
if (debugging)
{
	Response.Write("current_debug_status()<br />");
	Response.Write("m_eventid = "+m_eventid+"<br />");
	Response.Write("m_eventdate = "+m_eventdate+"<br />");
	Response.Write("m_eventtime = "+m_eventtime+"<br />");
	Response.Write("m_eventnote = "+m_eventnote+"<br />");
	Response.Write("m_eventreport = "+m_eventreport+"<br />");
	Response.Write("m_eventyear = "+m_eventyear+"<br />");
	Response.Write("<br />SQL1 = ["+SQL1+"]<br />");
}
if (updating)
	RS = ConnObj.Execute(SQL1);
// set real unique id
RS=null;
ConnObj.Close();
ConnObj=null;
if ((! current_debug_status()) && updating )
	Response.Redirect("/admin/#{/events");
%>

