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
If key = "" Or IsNull(key) Then Response.Redirect "otherclubslist.asp"
sqlKey = sqlKey & "[clubid]=" & "" & key & ""

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
		strsql = "SELECT * FROM [otherclubs] WHERE " & sqlKey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.Eof Then
			Response.Clear
			Response.Redirect "otherclubslist.asp"
		Else
			rs.MoveFirst
		End If
	Case "D": ' Delete
		strsql = "SELECT * FROM [otherclubs] WHERE " & sqlKey
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
		Response.Redirect "otherclubslist.asp"
End Select
%>
<!--#include file="header.asp"-->
<p><span class="aspmaker">Delete from TABLE: otherclubs<br><br><a href="otherclubslist.asp">Back to List</a></span></p>
<form action="otherclubsdelete.asp" method="post">
<p>
<input type="hidden" name="a" value="D">
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#CCCCCC">
	<tr bgcolor="#0099CC">
		<td><span class="aspmaker" style="color: #FFFFFF;">clubid</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">clubname</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">address 1</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">address 2</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">address 3</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">address 4</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">town</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">county</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">postcode</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">telephone</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">fax</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">email</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">contact</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">totaltennis</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">cluburl</span>&nbsp;</td>
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
	x_clubid = rs("clubid")
	x_clubname = rs("clubname")
	x_address1 = rs("address1")
	x_address2 = rs("address2")
	x_address3 = rs("address3")
	x_address4 = rs("address4")
	x_town = rs("town")
	x_county = rs("county")
	x_postcode = rs("postcode")
	x_telephone = rs("telephone")
	x_fax = rs("fax")
	x_email = rs("email")
	x_contact = rs("contact")
	x_totaltennis = rs("totaltennis")
	x_cluburl = rs("cluburl")
	x_comments = rs("comments")
%>
	<tr bgcolor="<%= bgcolor %>">
	<input type="hidden" name="key" value="<%= key %>">
		<td class="aspmaker"><% Response.Write x_clubid %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_clubname %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_address1 %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_address2 %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_address3 %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_address4 %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_town %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_county %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_postcode %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_telephone %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_fax %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_email %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_contact %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_totaltennis %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_cluburl %>&nbsp;</td>
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
