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
If key = "" Or IsNull(key) Then Response.Redirect "memberslist.asp"
sqlKey = sqlKey & "[memberid]=" & "'" & key & "'"

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
		strsql = "SELECT * FROM [members] WHERE " & sqlKey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.Eof Then
			Response.Clear
			Response.Redirect "memberslist.asp"
		Else
			rs.MoveFirst
		End If
	Case "D": ' Delete
		strsql = "SELECT * FROM [members] WHERE " & sqlKey
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
		Response.Redirect "memberslist.asp"
End Select
%>
<!--#include file="header.asp"-->
<p><span class="aspmaker">Delete from TABLE: members<br><br><a href="memberslist.asp">Back to List</a></span></p>
<form action="membersdelete.asp" method="post">
<p>
<input type="hidden" name="a" value="D">
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#CCCCCC">
	<tr bgcolor="#0099CC">
		<td><span class="aspmaker" style="color: #FFFFFF;">memberid</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">membergrade</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">surname</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">forename 1</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">initials</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">title</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">address 1</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">address 2</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">address 3</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">address 4</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">postcode</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">homephone</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">workphone</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">mobilephone</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">email</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">webpassword</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">webaccess</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">uniqueref</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">dob</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">mailing</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">joined</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">dateleft</span>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">gender</span>&nbsp;</td>
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
	x_memberid = rs("memberid")
	x_membergrade = rs("membergrade")
	x_surname = rs("surname")
	x_forename1 = rs("forename1")
	x_initials = rs("initials")
	x_title = rs("title")
	x_address1 = rs("address1")
	x_address2 = rs("address2")
	x_address3 = rs("address3")
	x_address4 = rs("address4")
	x_postcode = rs("postcode")
	x_homephone = rs("homephone")
	x_workphone = rs("workphone")
	x_mobilephone = rs("mobilephone")
	x_email = rs("email")
	x_webpassword = rs("webpassword")
	x_webaccess = rs("webaccess")
	x_uniqueref = rs("uniqueref")
	x_dob = rs("dob")
	x_mailing = rs("mailing")
	x_joined = rs("joined")
	x_dateleft = rs("dateleft")
	x_gender = rs("gender")
%>
	<tr bgcolor="<%= bgcolor %>">
	<input type="hidden" name="key" value="<%= key %>">
		<td class="aspmaker"><% Response.Write x_memberid %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_membergrade %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_surname %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_forename1 %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_initials %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_title %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_address1 %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_address2 %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_address3 %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_address4 %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_postcode %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_homephone %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_workphone %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_mobilephone %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_email %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_webpassword %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_webaccess %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_uniqueref %>&nbsp;</td>
		<td class="aspmaker"><% If IsDate(x_dob) Then Response.Write EW_FormatDateTime(x_dob,7) Else Response.Write x_dob End If %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_mailing %>&nbsp;</td>
		<td class="aspmaker"><% If IsDate(x_joined) Then Response.Write EW_FormatDateTime(x_joined,7) Else Response.Write x_joined End If %>&nbsp;</td>
		<td class="aspmaker"><% If IsDate(x_dateleft) Then Response.Write EW_FormatDateTime(x_dateleft,7) Else Response.Write x_dateleft End If %>&nbsp;</td>
		<td class="aspmaker"><% Response.Write x_gender %>&nbsp;</td>
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
