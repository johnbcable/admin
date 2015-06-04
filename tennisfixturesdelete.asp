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
If key = "" Or IsNull(key) Then Response.Redirect "tennisfixtureslist.asp"
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
		strsql = "SELECT * FROM [tennisfixtures] WHERE " & sqlKey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.Eof Then
			Response.Clear
			Response.Redirect "tennisfixtureslist.asp"
		Else
			rs.MoveFirst
		End If
	Case "D": ' Delete
		strsql = "SELECT * FROM [tennisfixtures] WHERE " & sqlKey
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
		Response.Redirect "tennisfixtureslist.asp"
End Select
%>
<!--#include file="header.asp"-->
<p><span class="aspmaker">Delete from TABLE: tennisfixtures<br><br><a href="tennisfixtureslist.asp">Back to List</a></span></p>
<form action="tennisfixturesdelete.asp" method="post">
<p>
<input type="hidden" name="a" value="D">
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#CCCCCC">
	<tr bgcolor="#0099CC">
		<td><span class="aspmaker" style="color: #FFFFFF;">fixturedate</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">homeoraway</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">opponents</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">hamptonresult</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">opponentresult</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">fixtureyear</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">teamname</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">fixturenote</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">fixtureid</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">matchreport</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">pair 1</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">pair 2</span>&nbsp;</td>
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
	x_homeoraway = rs("homeoraway")
	x_opponents = rs("opponents")
	x_hamptonresult = rs("hamptonresult")
	x_opponentresult = rs("opponentresult")
	x_fixtureyear = rs("fixtureyear")
	x_teamname = rs("teamname")
	x_fixturenote = rs("fixturenote")
	x_fixtureid = rs("fixtureid")
	x_matchreport = rs("matchreport")
	x_pair1 = rs("pair1")
	x_pair2 = rs("pair2")
%>
	<tr bgcolor="<%= bgcolor %>">
	<input type="hidden" name="key" value="<%= key %>">
		<td class="aspmaker"><% If IsDate(x_fixturedate) Then Response.Write EW_FormatDateTime(x_fixturedate,7) Else Response.Write x_fixturedate End If %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_homeoraway %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_opponents %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_hamptonresult %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_opponentresult %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_fixtureyear %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_teamname %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_fixturenote %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_fixtureid %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_matchreport %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_pair1 %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_pair2 %>&nbsp;</td>
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
