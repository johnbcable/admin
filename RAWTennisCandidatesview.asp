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
If key = "" Or IsNull(key) Then Response.Redirect "RAWTennisCandidateslist.asp"

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
		strsql = "SELECT * FROM [RAWTennisCandidates] WHERE []=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.Eof Then
			Response.Clear
			Response.Redirect "RAWTennisCandidateslist.asp"
		Else
			rs.MoveFirst
		End If

		' Get the field contents
		x_memberid = rs("memberid")
		x_membergrade = rs("membergrade")
		x_surname = rs("surname")
		x_forename1 = rs("forename1")
		x_initials = rs("initials")
		x_email = rs("email")
		x_uniqueref = rs("uniqueref")
		x_dob = rs("dob")
		x_mailing = rs("mailing")
		x_joined = rs("joined")
		x_dateleft = rs("dateleft")
		rs.Close
		Set rs = Nothing
End Select
%>
<!--#include file="header.asp"-->
<p><span class="aspmaker">View VIEW: RAWTennis Candidates<br><br><a href="RAWTennisCandidateslist.asp">Back to List</a></span></p>
<p>
<form>
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#CCCCCC">
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">memberid</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_memberid %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">membergrade</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_membergrade %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">surname</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_surname %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">forename 1</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_forename1 %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">initials</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_initials %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">email</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_email %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">uniqueref</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_uniqueref %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">dob</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% If IsDate(x_dob) Then Response.Write EW_FormatDateTime(x_dob,7) Else Response.Write x_dob End If %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">mailing</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_mailing %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">joined</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% If IsDate(x_joined) Then Response.Write EW_FormatDateTime(x_joined,7) Else Response.Write x_joined End If %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">dateleft</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% If IsDate(x_dateleft) Then Response.Write EW_FormatDateTime(x_dateleft,7) Else Response.Write x_dateleft End If %></span>&nbsp;</td>
	</tr>
</table>
</form>
<p>
<!--#include file="footer.asp"-->
