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

' Single delete record
key = Request.querystring("key")
If key = "" Or IsNull(key) Then
	key = Request.Form("key")
End If
If key = "" Or IsNull(key) Then Response.Redirect "eventslist.asp"
sqlKey = sqlKey & "[eventid]=" & "" & key & ""

' Get action
a = Request.Form("a")
If a = "" Or IsNull(a) Then
	a = "I"	' Display with input box
End If

' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Select Case a
	Case "I": ' Display
		strsql = "SELECT * FROM [events] WHERE " & sqlKey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.Eof Then
			Response.Clear
			Response.Redirect "eventslist.asp"
		Else
			rs.MoveFirst
		End If
	Case "D": ' Delete
		strsql = "SELECT * FROM [events] WHERE " & sqlKey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn, 1, 2
		Do While Not rs.Eof
			rs.Delete
			rs.MoveNext
		Loop
		rs.Close
		Set rs = Nothing
		conn.Close
		Set conn = Nothing		
		Response.Clear
		Response.Redirect "eventslist.asp"
End Select
%>
<!--#include file="header.asp"-->
<p><span class="aspmaker">Delete from TABLE: events<br><br><a href="eventslist.asp">Back to List</a></span></p>
<form action="eventsdelete.asp" method="post">
<p>
<input type="hidden" name="a" value="D">
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#CCCCCC">
	<tr bgcolor="#0099CC">
		<td><span class="aspmaker" style="color: #FFFFFF;">eventdate</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">eventtime</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">eventyear</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">eventtype</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">eventnote</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">eventid</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">eventreport</span>&nbsp;</td>
	</tr>
<%
recCount = 0
Do While Not rs.Eof
	recCount = recCount + 1

	' Set row color
	bgcolor = "#FFFFFF"
%>
<%	

	' Display alternate color for rows
	If recCount Mod 2 <> 0 Then
		bgcolor = "#F5F5F5"
	End If
%>
<%
	x_eventdate = rs("eventdate")
	x_eventtime = rs("eventtime")
	x_eventyear = rs("eventyear")
	x_eventtype = rs("eventtype")
	x_eventnote = rs("eventnote")
	x_eventid = rs("eventid")
	x_eventreport = rs("eventreport")
%>
	<tr bgcolor="<%= bgcolor %>">
	<input type="hidden" name="key" value="<%= key %>">
		<td class="aspmaker"><% If IsDate(x_eventdate) Then Response.Write EW_FormatDateTime(x_eventdate,7) Else Response.Write x_eventdate End If %>&nbsp;</td>
		<td class="aspmaker"><% If IsDate(x_eventtime) Then Response.Write EW_FormatDateTime(x_eventtime,7) Else Response.Write x_eventtime End If %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_eventyear %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_eventtype %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_eventnote %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_eventid %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_eventreport %>&nbsp;</td>
  </tr>
<%
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%>
</table>
<p>
<input type="submit" name="Action" value="CONFIRM DELETE">
</form>
<!--#include file="footer.asp"-->
