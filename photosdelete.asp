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
If key = "" Or IsNull(key) Then Response.Redirect "photoslist.asp"
sqlKey = sqlKey & "[mainphoto]=" & "'" & key & "'"

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
		strsql = "SELECT * FROM [photos] WHERE " & sqlKey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.Eof Then
			Response.Clear
			Response.Redirect "photoslist.asp"
		Else
			rs.MoveFirst
		End If
	Case "D": ' Delete
		strsql = "SELECT * FROM [photos] WHERE " & sqlKey
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
		Response.Redirect "photoslist.asp"
End Select
%>
<!--#include file="header.asp"-->
<p><span class="aspmaker">Delete from TABLE: photos<br><br><a href="photoslist.asp">Back to List</a></span></p>
<form action="photosdelete.asp" method="post">
<p>
<input type="hidden" name="a" value="D">
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#CCCCCC">
	<tr bgcolor="#0099CC">
		<td><span class="aspmaker" style="color: #FFFFFF;">photoid</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">mainphoto</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">thumbnail</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">bottomcaption</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">topcaption</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">altcaption</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">takenon</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">submittedby</span>&nbsp;</td>
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
	x_photoid = rs("photoid")
	x_mainphoto = rs("mainphoto")
	x_thumbnail = rs("thumbnail")
	x_bottomcaption = rs("bottomcaption")
	x_topcaption = rs("topcaption")
	x_altcaption = rs("altcaption")
	x_takenon = rs("takenon")
	x_submittedby = rs("submittedby")
%>
	<tr bgcolor="<%= bgcolor %>">
	<input type="hidden" name="key" value="<%= key %>">
		<td class="aspmaker"><% Response.Write x_photoid %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_mainphoto %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_thumbnail %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_bottomcaption %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_topcaption %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_altcaption %>&nbsp;</td>
		<td class="aspmaker"><% If IsDate(x_takenon) Then Response.Write EW_FormatDateTime(x_takenon,7) Else Response.Write x_takenon End If %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_submittedby %>&nbsp;</td>
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
