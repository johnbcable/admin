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

' Get action
a = Request.Form("a")
If (a = "" Or IsNull(a)) Then
	key = Request.Querystring("key")
	If key <> "" Then
		a = "C" ' Copy record
	Else
		a = "I" ' Display blank record
	End If
End If

' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Select Case a
	Case "C": ' Get a record to display
		tkey = "" & key & ""
		strsql = "SELECT * FROM [tennisteams] WHERE [teamid]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.Eof Then
			Response.Clear
			Response.Redirect "tennisteamslist.asp"
		Else
			rs.MoveFirst

			' Get the field contents
			x_teamname = rs("teamname")
			x_teamcategory = rs("teamcategory")
			x_teamcaptain = rs("teamcaptain")
			x_teamnote = rs("teamnote")
			x_fixtureline = rs("fixtureline")
			x_uniqueref = rs("uniqueref")
			x_email = rs("email")
		End If
		rs.Close
		Set rs = Nothing
	Case "A": ' Add

		' Get fields from form
x_teamid = Request.Form("x_teamid")
x_teamname = Request.Form("x_teamname")
x_teamcategory = Request.Form("x_teamcategory")
x_teamcaptain = Request.Form("x_teamcaptain")
x_teamnote = Request.Form("x_teamnote")
x_fixtureline = Request.Form("x_fixtureline")
x_uniqueref = Request.Form("x_uniqueref")
x_email = Request.Form("x_email")

		' Open record
		strsql = "SELECT * FROM [tennisteams] WHERE 0 = 1"
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn, 1, 2
		rs.AddNew
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
<p><span class="aspmaker">Add to TABLE: tennisteams<br><br><a href="tennisteamslist.asp">Back to List</a></span></p>
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
<form onSubmit="return EW_checkMyForm(this);"  action="tennisteamsadd.asp" method="post">
<p>
<input type="hidden" name="a" value="A">
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#CCCCCC">
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
<input type="submit" name="Action" value="ADD">
</form>
<!--#include file="footer.asp"-->
