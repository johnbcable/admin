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
If key = "" Or IsNull(key) Then Response.Redirect "otherteamslist.asp"

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
		strsql = "SELECT * FROM [otherteams] WHERE [uniqref]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.Eof Then
			Response.Clear
			Response.Redirect "otherteamslist.asp"
		Else
			rs.MoveFirst
		End If

		' Get the field contents
		x_uniqref = rs("uniqref")
		x_clubname = rs("clubname")
		x_teamname = rs("teamname")
		x_comments = rs("comments")
		rs.Close
		Set rs = Nothing
End Select
%>
<!--#include file="header.asp"-->
<p><span class="aspmaker">View TABLE: otherteams<br><br><a href="otherteamslist.asp">Back to List</a></span></p>
<p>
<form>
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#CCCCCC">
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">uniqref</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_uniqref %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">clubname</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_clubname %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">teamname</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_teamname %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">comments</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><%= Replace(x_comments & "", vbLf, "<br>") %></span>&nbsp;</td>
	</tr>
</table>
</form>
<p>
<!--#include file="footer.asp"-->
