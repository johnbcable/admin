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

' Get fields from form
x_uniqref = Request.Form("x_uniqref")
x_clubname = Request.Form("x_clubname")
x_teamname = Request.Form("x_teamname")
x_comments = Request.Form("x_comments")

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
	Case "U": ' Update

		' Open record
		tkey = "" & key & ""
		strsql = "SELECT * FROM [otherteams] WHERE [uniqref]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn, 1, 2
		If rs.Eof Then
			Response.Clear
			Response.Redirect "otherteamslist.asp"
		End If
		tmpFld = Trim(x_clubname)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("clubname") = tmpFld
		tmpFld = Trim(x_teamname)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("teamname") = tmpFld
		tmpFld = Trim(x_comments)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("comments") = tmpFld
		rs.Update
		rs.Close
		Set rs = Nothing
		conn.Close
		Set conn = Nothing
		Response.Clear
		Response.Redirect "otherteamslist.asp"
End Select
%>
<!--#include file="header.asp"-->
<p><span class="aspmaker">Edit TABLE: otherteams<br><br><a href="otherteamslist.asp">Back to List</a></span></p>
<script language="JavaScript" src="ew.js"></script>
<script language="JavaScript">
<!-- start Javascript
function  EW_checkMyForm(EW_this) {
return true;
}
// end JavaScript -->
</script>
<form onSubmit="return EW_checkMyForm(this);"  name="otherteamsedit" action="otherteamsedit.asp" method="post">
<p>
<input type="hidden" name="a" value="U">
<input type="hidden" name="key" value="<%= key %>">
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#CCCCCC">
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">uniqref</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_uniqref %><input type="hidden" name="x_uniqref" value="<%= x_uniqref %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">clubname</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_clubname" size="30" maxlength="60" value="<%= Server.HTMLEncode(x_clubname&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">teamname</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_teamname" size="30" maxlength="60" value="<%= Server.HTMLEncode(x_teamname&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">comments</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><textarea cols=35 rows=4 name="x_comments"><%= x_comments %></textarea></span>&nbsp;</td>
	</tr>
</table>
<p>
<input type="submit" name="Action" value="EDIT">
</form>
<!--#include file="footer.asp"-->
