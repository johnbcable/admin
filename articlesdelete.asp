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
If key = "" Or IsNull(key) Then Response.Redirect "articleslist.asp"
sqlKey = sqlKey & "[articleid]=" & "" & key & ""

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
		strsql = "SELECT * FROM [articles] WHERE " & sqlKey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.Eof Then
			Response.Clear
			Response.Redirect "articleslist.asp"
		Else
			rs.MoveFirst
		End If
	Case "D": ' Delete
		strsql = "SELECT * FROM [articles] WHERE " & sqlKey
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
		Response.Redirect "articleslist.asp"
End Select
%>
<!--#include file="header.asp"-->
<p><span class="aspmaker">Delete from TABLE: articles<br><br><a href="articleslist.asp">Back to List</a></span></p>
<form action="articlesdelete.asp" method="post">
<p>
<input type="hidden" name="a" value="D">
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#CCCCCC">
	<tr bgcolor="#0099CC">
		<td><span class="aspmaker" style="color: #FFFFFF;">articleid</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">title</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">author</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">courtcircularissueno</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">courtseq</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">courtcols</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">newsitem</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">newsfrom</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">newsuntil</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">newspriority</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">articlesection</span>&nbsp;</td>
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
	x_articleid = rs("articleid")
	x_title = rs("title")
	x_author = rs("author")
	x_courtcircularissueno = rs("courtcircularissueno")
	x_courtseq = rs("courtseq")
	x_courtcols = rs("courtcols")
	x_articletext = rs("articletext")
	x_newsitem = rs("newsitem")
	x_newsfrom = rs("newsfrom")
	x_newsuntil = rs("newsuntil")
	x_newspriority = rs("newspriority")
	x_articlesection = rs("articlesection")
%>
	<tr bgcolor="<%= bgcolor %>">
	<input type="hidden" name="key" value="<%= key %>">
		<td class="aspmaker"><% Response.Write x_articleid %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_title %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_author %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_courtcircularissueno %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_courtseq %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_courtcols %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_newsitem %>&nbsp;</td>
		<td class="aspmaker"><% If IsDate(x_newsfrom) Then Response.Write EW_FormatDateTime(x_newsfrom,7) Else Response.Write x_newsfrom End If %>&nbsp;</td>
		<td class="aspmaker"><% If IsDate(x_newsuntil) Then Response.Write EW_FormatDateTime(x_newsuntil,7) Else Response.Write x_newsuntil End If %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_newspriority %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_articlesection %>&nbsp;</td>
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
