<% If Session("project1_status") <> "login" Then Response.Redirect "login.asp" %>
<%
Response.expires = 0
Response.expiresabsolute = Now() - 1
Response.addHeader "pragma", "no-cache"
Response.addHeader "cache-control", "private"
Response.CacheControl = "no-cache"
%>
<!--#include file="db.asp"-->
<!--#include file="aspmkrfn.asp"-->
<%
Response.Buffer = True
key = Request.Querystring("key")
If key = "" Or IsNull(key) Then key = Request.Form("key")
If key = "" Or IsNull(key) Then Response.Redirect "eventslist.asp"

' Get action
a = Request.Form("a")
If a = "" Or IsNull(a) Then
	a = "I"	' Display with input box
End If

' Get fields from form
x_eventdate = Request.Form("x_eventdate")
x_eventtime = Request.Form("x_eventtime")
x_eventyear = Request.Form("x_eventyear")
x_eventtype = Request.Form("x_eventtype")
x_eventnote = Request.Form("x_eventnote")
x_eventid = Request.Form("x_eventid")
x_eventreport = Request.Form("x_eventreport")

' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Select Case a
	Case "I": ' Get a record to display
		tkey = "" & key & ""
		strsql = "SELECT * FROM [events] WHERE [eventid]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.Eof Then
			Response.Clear
			Response.Redirect "eventslist.asp"
		Else
			rs.MoveFirst
		End If

		' Get the field contents
		x_eventdate = rs("eventdate")
		x_eventtime = rs("eventtime")
		x_eventyear = rs("eventyear")
		x_eventtype = rs("eventtype")
		x_eventnote = rs("eventnote")
		x_eventid = rs("eventid")
		x_eventreport = rs("eventreport")
		rs.Close
		Set rs = Nothing
	Case "U": ' Update

		' Open record
		tkey = "" & key & ""
		strsql = "SELECT * FROM [events] WHERE [eventid]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn, 1, 2
		If rs.Eof Then
			Response.Clear
			Response.Redirect "eventslist.asp"
		End If
		tmpFld = EW_UnFormatDateTime(x_eventdate,7)
		If IsDate(tmpFld) Then
		    rs("eventdate") = CDate(tmpFld)
		Else
		    rs("eventdate") = Null
		End If
		tmpFld = EW_UnFormatDateTime(x_eventtime,7)
		If IsDate(tmpFld) Then
		    rs("eventtime") = CDate(tmpFld)
		Else
		    rs("eventtime") = Null
		End If
		tmpFld = x_eventyear
		If Not IsNumeric(tmpFld) Then tmpFld = 0
		rs("eventyear") = CLng(tmpFld)
		tmpFld = Trim(x_eventtype)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("eventtype") = tmpFld
		tmpFld = Trim(x_eventnote)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("eventnote") = tmpFld
		tmpFld = Trim(x_eventreport)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("eventreport") = tmpFld
		rs.Update
		rs.Close
		Set rs = Nothing
		conn.Close
		Set conn = Nothing
		Response.Clear
		Response.Redirect "eventslist.asp"
End Select
%>
<!--#include file="header.asp"-->
<p><span class="aspmaker">Edit TABLE: events<br><br><a href="eventslist.asp">Back to List</a></span></p>
<script language="JavaScript" src="ew.js"></script>
<script language="JavaScript">
<!-- start Javascript
function  EW_checkMyForm(EW_this) {
if (EW_this.x_eventdate && !EW_checkeurodate(EW_this.x_eventdate.value)) {
        if (!EW_onError(EW_this, EW_this.x_eventdate, "TEXT", "Incorrect date (dd/mm/yyyy) - eventdate"))
            return false; 
        }
if (EW_this.x_eventtime && !EW_checkeurodate(EW_this.x_eventtime.value)) {
        if (!EW_onError(EW_this, EW_this.x_eventtime, "TEXT", "Incorrect date (dd/mm/yyyy) - eventtime"))
            return false; 
        }
if (EW_this.x_eventyear && !EW_checkinteger(EW_this.x_eventyear.value)) {
        if (!EW_onError(EW_this, EW_this.x_eventyear, "TEXT", "Incorrect integer - eventyear"))
            return false; 
        }
return true;
}
// end JavaScript -->
</script>
<form onSubmit="return EW_checkMyForm(this);"  name="eventsedit" action="eventsedit.asp" method="post">
<p>
<input type="hidden" name="a" value="U">
<input type="hidden" name="key" value="<%= key %>">
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#CCCCCC">
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">eventdate</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_eventdate" value="<% If IsDate(x_eventdate) Then Response.Write EW_FormatDateTime(x_eventdate,7) Else Response.Write x_eventdate End If %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">eventtime</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_eventtime" value="<% If IsDate(x_eventtime) Then Response.Write EW_FormatDateTime(x_eventtime,7) Else Response.Write x_eventtime End If %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">eventyear</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_eventyear" size="30" value="<%= Server.HTMLEncode(x_eventyear&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">eventtype</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_eventtype" size="30" maxlength="30" value="<%= Server.HTMLEncode(x_eventtype&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">eventnote</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_eventnote" size="30" maxlength="100" value="<%= Server.HTMLEncode(x_eventnote&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">eventid</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_eventid %><input type="hidden" name="x_eventid" value="<%= x_eventid %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">eventreport</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_eventreport" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_eventreport&"") %>"></span>&nbsp;</td>
	</tr>
</table>
<p>
<input type="submit" name="Action" value="EDIT">
</form>
<!--#include file="footer.asp"-->
