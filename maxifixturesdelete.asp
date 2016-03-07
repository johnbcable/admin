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
If key = "" Or IsNull(key) Then Response.Redirect "maxifixtureslist.asp"
sqlKey = sqlKey & "[fixtureid]=" & "" & key & ""

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
		strsql = "SELECT * FROM [maxifixtures] WHERE " & sqlKey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.Eof Then
			Response.Clear
			Response.Redirect "maxifixtureslist.asp"
		Else
			rs.MoveFirst
		End If
	Case "D": ' Delete
		strsql = "SELECT * FROM [maxifixtures] WHERE " & sqlKey
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
		Response.Redirect "maxifixtureslist.asp"
End Select
%>
<!--#include file="header.asp"-->
<p><span class="aspmaker">Delete from TABLE: maxifixtures<br><br><a href="maxifixtureslist.asp">Back to List</a></span></p>
<form action="maxifixturesdelete.asp" method="post">
<p>
<input type="hidden" name="a" value="D">
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#CCCCCC">
	<tr bgcolor="#0099CC">
		<td><span class="aspmaker" style="color: #FFFFFF;">fixturedate</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">hometeam</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">awayteam</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">homescore</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">awayscore</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">fixtureyear</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">fixturenote</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">fixtureid</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">matchreport</span>&nbsp;</td>
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
	x_fixturedate = rs("fixturedate")
	x_hometeam = rs("hometeam")
	x_awayteam = rs("awayteam")
	x_homescore = rs("homescore")
	x_awayscore = rs("awayscore")
	x_fixtureyear = rs("fixtureyear")
	x_fixturenote = rs("fixturenote")
	x_fixtureid = rs("fixtureid")
	x_matchreport = rs("matchreport")
%>
	<tr bgcolor="<%= bgcolor %>">
	<input type="hidden" name="key" value="<%= key %>">
		<td class="aspmaker"><% If IsDate(x_fixturedate) Then Response.Write EW_FormatDateTime(x_fixturedate,7) Else Response.Write x_fixturedate End If %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_hometeam %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_awayteam %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_homescore %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_awayscore %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_fixtureyear %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_fixturenote %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_fixtureid %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_matchreport %>&nbsp;</td>
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
