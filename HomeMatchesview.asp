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
If key = "" Or IsNull(key) Then Response.Redirect "HomeMatcheslist.asp"

' Get action
a = Request.Form("a")
If a = "" Or IsNull(a) Then
	a = "I"	' Display with input box
End If

' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Select Case a
	Case "I": ' Get a record to display
		tkey = "" & key & ""
		strsql = "SELECT * FROM [HomeMatches] WHERE []=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.Eof Then
			Response.Clear
			Response.Redirect "HomeMatcheslist.asp"
		Else
			rs.MoveFirst
		End If

		' Get the field contents
		x_fixturedate = rs("fixturedate")
		x_homeoraway = rs("homeoraway")
		x_opponents = rs("opponents")
		x_hamptonresult = rs("hamptonresult")
		x_opponentresult = rs("opponentresult")
		x_Fixtureyear = rs("Fixtureyear")
		x_teamname = rs("teamname")
		x_fixturenote = rs("fixturenote")
		x_fixtureid = rs("fixtureid")
		x_matchreport = rs("matchreport")
		rs.Close
		Set rs = Nothing
End Select
%>
<!--#include file="header.asp"-->
<p><span class="aspmaker">View VIEW: Home Matches<br><br><a href="HomeMatcheslist.asp">Back to List</a></span></p>
<p>
<form>
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#CCCCCC">
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">fixturedate</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% If IsDate(x_fixturedate) Then Response.Write EW_FormatDateTime(x_fixturedate,7) Else Response.Write x_fixturedate End If %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">homeoraway</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_homeoraway %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">opponents</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_opponents %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">hamptonresult</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_hamptonresult %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">opponentresult</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_opponentresult %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">Fixtureyear</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_Fixtureyear %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">teamname</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_teamname %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">fixturenote</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_fixturenote %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">fixtureid</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_fixtureid %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">matchreport</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_matchreport %></span>&nbsp;</td>
	</tr>
</table>
</form>
<p>
<!--#include file="footer.asp"-->
