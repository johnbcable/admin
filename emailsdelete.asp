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
If key = "" Or IsNull(key) Then Response.Redirect "emailslist.asp"
sqlKey = sqlKey & "[emailid]=" & "" & key & ""

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
		strsql = "SELECT * FROM [emails] WHERE " & sqlKey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.Eof Then
			Response.Clear
			Response.Redirect "emailslist.asp"
		Else
			rs.MoveFirst
		End If
	Case "D": ' Delete
		strsql = "SELECT * FROM [emails] WHERE " & sqlKey
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
		Response.Redirect "emailslist.asp"
End Select
%>
<!--#include file="header.asp"-->
<p><span class="aspmaker">Delete from TABLE: emails<br><br><a href="emailslist.asp">Back to List</a></span></p>
<form action="emailsdelete.asp" method="post">
<p>
<input type="hidden" name="a" value="D">
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#CCCCCC">
	<tr bgcolor="#0099CC">
		<td><span class="aspmaker" style="color: #FFFFFF;">emailid</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">subject</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">emailfile</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">attach 1</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">attach 2</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">attach 3</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">circulation</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">sent on</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">number sent</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">comment</span>&nbsp;</td>
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
	x_emailid = rs("emailid")
	x_subject = rs("subject")
	x_emailfile = rs("emailfile")
	x_attach1 = rs("attach1")
	x_attach2 = rs("attach2")
	x_attach3 = rs("attach3")
	x_circulation = rs("circulation")
	x_sent_on = rs("sent_on")
	x_number_sent = rs("number_sent")
	x_comment = rs("comment")
%>
	<tr bgcolor="<%= bgcolor %>">
	<input type="hidden" name="key" value="<%= key %>">
		<td class="aspmaker"><% Response.Write x_emailid %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_subject %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_emailfile %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_attach1 %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_attach2 %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_attach3 %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_circulation %>&nbsp;</td>
		<td class="aspmaker"><% If IsDate(x_sent_on) Then Response.Write EW_FormatDateTime(x_sent_on,7) Else Response.Write x_sent_on End If %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_number_sent %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_comment %>&nbsp;</td>
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
