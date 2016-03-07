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
If key = "" Or IsNull(key) Then Response.Redirect "tennisteamslist.asp"

' Get action
a = Request.Form("a")
If a = "" Or IsNull(a) Then
	a = "I"	' Display with input box
End If

' Get fields from form
x_teamid = Request.Form("x_teamid")
x_teamname = Request.Form("x_teamname")
x_teamcategory = Request.Form("x_teamcategory")
x_teamcaptain = Request.Form("x_teamcaptain")
x_teamnote = Request.Form("x_teamnote")
x_fixtureline = Request.Form("x_fixtureline")
x_uniqueref = Request.Form("x_uniqueref")
x_email = Request.Form("x_email")

' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Select Case a
	Case "I": ' Get a record to display
		tkey = "" & key & ""
		strsql = "SELECT * FROM [tennisteams] WHERE [teamid]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.Eof Then
			Response.Clear
			Response.Redirect "tennisteamslist.asp"
		Else
			rs.MoveFirst
		End If

		' Get the field contents
		x_teamid = rs("teamid")
		x_teamname = rs("teamname")
		x_teamcategory = rs("teamcategory")
		x_teamcaptain = rs("teamcaptain")
		x_teamnote = rs("teamnote")
		x_fixtureline = rs("fixtureline")
		x_uniqueref = rs("uniqueref")
		x_email = rs("email")
		rs.Close
		Set rs = Nothing
	Case "U": ' Update

		' Open record
		tkey = "" & key & ""
		strsql = "SELECT * FROM [tennisteams] WHERE [teamid]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn, 1, 2
		If rs.Eof Then
			Response.Clear
			Response.Redirect "tennisteamslist.asp"
		End If
		tmpFld = Trim(x_teamname)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("teamname") = tmpFld
		tmpFld = Trim(x_teamcategory)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("teamcategory") = tmpFld
		tmpFld = Trim(x_teamcaptain)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("teamcaptain") = tmpFld
		tmpFld = Trim(x_teamnote)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("teamnote") = tmpFld
		tmpFld = Trim(x_fixtureline)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("fixtureline") = tmpFld
		tmpFld = x_uniqueref
		If Not IsNumeric(tmpFld) Then tmpFld = 0
		rs("uniqueref") = CLng(tmpFld)
		tmpFld = Trim(x_email)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("email") = tmpFld
		rs.Update
		rs.Close
		Set rs = Nothing
		conn.Close
		Set conn = Nothing
		Response.Clear
		Response.Redirect "tennisteamslist.asp"
End Select
%>
<!--#include file="header.asp"-->
<p><span class="aspmaker">Edit TABLE: tennisteams<br><br><a href="tennisteamslist.asp">Back to List</a></span></p>
<script language="JavaScript" src="ew.js"></script>
<script language="JavaScript">
<!-- start Javascript
function  EW_checkMyForm(EW_this) {
if (EW_this.x_uniqueref && !EW_checkinteger(EW_this.x_uniqueref.value)) {
        if (!EW_onError(EW_this, EW_this.x_uniqueref, "TEXT", "Incorrect integer - uniqueref"))
            return false; 
        }
return true;
}
// end JavaScript -->
</script>
<form onSubmit="return EW_checkMyForm(this);"  name="tennisteamsedit" action="tennisteamsedit.asp" method="post">
<p>
<input type="hidden" name="a" value="U">
<input type="hidden" name="key" value="<%= key %>">
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#CCCCCC">
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">teamid</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_teamid %><input type="hidden" name="x_teamid" value="<%= x_teamid %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">teamname</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_teamname" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_teamname&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">teamcategory</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_teamcategory" size="30" maxlength="20" value="<%= Server.HTMLEncode(x_teamcategory&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">teamcaptain</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_teamcaptain" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_teamcaptain&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">teamnote</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_teamnote" size="30" maxlength="255" value="<%= Server.HTMLEncode(x_teamnote&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">fixtureline</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_fixtureline" size="30" maxlength="150" value="<%= Server.HTMLEncode(x_fixtureline&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">uniqueref</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_uniqueref" size="30" value="<%= Server.HTMLEncode(x_uniqueref&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">email</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_email" size="30" maxlength="100" value="<%= Server.HTMLEncode(x_email&"") %>"></span>&nbsp;</td>
	</tr>
</table>
<p>
<input type="submit" name="Action" value="EDIT">
</form>
<!--#include file="footer.asp"-->
