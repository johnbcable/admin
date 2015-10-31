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
If key = "" Or IsNull(key) Then Response.Redirect "tennisteamslist.asp"
sqlKey = sqlKey & "[teamid]=" & "" & key & ""

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
		strsql = "SELECT * FROM [tennisteams] WHERE " & sqlKey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.Eof Then
			Response.Clear
			Response.Redirect "tennisteamslist.asp"
		Else
			rs.MoveFirst
		End If
	Case "D": ' Delete
		strsql = "SELECT * FROM [tennisteams] WHERE " & sqlKey
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
		Response.Redirect "tennisteamslist.asp"
End Select
%>
<!--#include file="header.asp"-->
<p><span class="aspmaker">Delete from TABLE: tennisteams<br><br><a href="tennisteamslist.asp">Back to List</a></span></p>
<form action="tennisteamsdelete.asp" method="post">
<p>
<input type="hidden" name="a" value="D">
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#CCCCCC">
	<tr bgcolor="#0099CC">
		<td><span class="aspmaker" style="color: #FFFFFF;">teamid</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">teamname</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">teamcategory</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">teamcaptain</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">teamnote</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">fixtureline</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">uniqueref</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">email</span>&nbsp;</td>
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
	x_teamid = rs("teamid")
	x_teamname = rs("teamname")
	x_teamcategory = rs("teamcategory")
	x_teamcaptain = rs("teamcaptain")
	x_teamnote = rs("teamnote")
	x_fixtureline = rs("fixtureline")
	x_uniqueref = rs("uniqueref")
	x_email = rs("email")
%>
	<tr bgcolor="<%= bgcolor %>">
	<input type="hidden" name="key" value="<%= key %>">
		<td class="aspmaker"><% Response.Write x_teamid %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_teamname %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_teamcategory %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_teamcaptain %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_teamnote %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_fixtureline %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_uniqueref %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_email %>&nbsp;</td>
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
