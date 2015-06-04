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
If key = "" Or IsNull(key) Then Response.Redirect "maxifixtureslist.asp"

' Get action
a = Request.Form("a")
If a = "" Or IsNull(a) Then
	a = "I"	' Display with input box
End If

' Get fields from form
x_fixturedate = Request.Form("x_fixturedate")
x_hometeam = Request.Form("x_hometeam")
x_awayteam = Request.Form("x_awayteam")
x_homescore = Request.Form("x_homescore")
x_awayscore = Request.Form("x_awayscore")
x_fixtureyear = Request.Form("x_fixtureyear")
x_fixturenote = Request.Form("x_fixturenote")
x_fixtureid = Request.Form("x_fixtureid")
x_matchreport = Request.Form("x_matchreport")

' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Select Case a
	Case "I": ' Get a record to display
		tkey = "" & key & ""
		strsql = "SELECT * FROM [maxifixtures] WHERE [fixtureid]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.Eof Then
			Response.Clear
			Response.Redirect "maxifixtureslist.asp"
		Else
			rs.MoveFirst
		End If

		' Get the field contents
		x_fixturedate = rs("fixturedate")
		x_hometeam = rs("hometeam")
		x_awayteam = rs("awayteam")
		x_homescore = rs("homescore")
		x_awayscore = rs("awayscore")
		x_fixtureyear = rs("fixtureyear")
		x_fixturenote = rs("fixturenote")
		x_fixtureid = rs("fixtureid")
		x_matchreport = rs("matchreport")
		rs.Close
		Set rs = Nothing
	Case "U": ' Update

		' Open record
		tkey = "" & key & ""
		strsql = "SELECT * FROM [maxifixtures] WHERE [fixtureid]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn, 1, 2
		If rs.Eof Then
			Response.Clear
			Response.Redirect "maxifixtureslist.asp"
		End If
		tmpFld = EW_UnFormatDateTime(x_fixturedate,7)
		If IsDate(tmpFld) Then
		    rs("fixturedate") = CDate(tmpFld)
		Else
		    rs("fixturedate") = Null
		End If
		tmpFld = Trim(x_hometeam)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("hometeam") = tmpFld
		tmpFld = Trim(x_awayteam)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("awayteam") = tmpFld
		tmpFld = x_homescore
		If Not IsNumeric(tmpFld) Then tmpFld = 0
		rs("homescore") = CLng(tmpFld)
		tmpFld = x_awayscore
		If Not IsNumeric(tmpFld) Then tmpFld = 0
		rs("awayscore") = CLng(tmpFld)
		tmpFld = x_fixtureyear
		If Not IsNumeric(tmpFld) Then tmpFld = 0
		rs("fixtureyear") = CLng(tmpFld)
		tmpFld = Trim(x_fixturenote)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("fixturenote") = tmpFld
		tmpFld = Trim(x_matchreport)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("matchreport") = tmpFld
		rs.Update
		rs.Close
		Set rs = Nothing
		conn.Close
		Set conn = Nothing
		Response.Clear
		Response.Redirect "maxifixtureslist.asp"
End Select
%>
<!--#include file="header.asp"-->
<p><span class="aspmaker">Edit TABLE: maxifixtures<br><br><a href="maxifixtureslist.asp">Back to List</a></span></p>
<script language="JavaScript" src="ew.js"></script>
<script language="JavaScript">
<!-- start Javascript
function  EW_checkMyForm(EW_this) {
if (EW_this.x_fixturedate && !EW_checkeurodate(EW_this.x_fixturedate.value)) {
        if (!EW_onError(EW_this, EW_this.x_fixturedate, "TEXT", "Incorrect date (dd/mm/yyyy) - fixturedate"))
            return false; 
        }
if (EW_this.x_homescore && !EW_checkinteger(EW_this.x_homescore.value)) {
        if (!EW_onError(EW_this, EW_this.x_homescore, "TEXT", "Incorrect integer - homescore"))
            return false; 
        }
if (EW_this.x_awayscore && !EW_checkinteger(EW_this.x_awayscore.value)) {
        if (!EW_onError(EW_this, EW_this.x_awayscore, "TEXT", "Incorrect integer - awayscore"))
            return false; 
        }
if (EW_this.x_fixtureyear && !EW_checkinteger(EW_this.x_fixtureyear.value)) {
        if (!EW_onError(EW_this, EW_this.x_fixtureyear, "TEXT", "Incorrect integer - fixtureyear"))
            return false; 
        }
return true;
}
// end JavaScript -->
</script>
<form onSubmit="return EW_checkMyForm(this);"  name="maxifixturesedit" action="maxifixturesedit.asp" method="post">
<p>
<input type="hidden" name="a" value="U">
<input type="hidden" name="key" value="<%= key %>">
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#CCCCCC">
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">fixturedate</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_fixturedate" value="<% If IsDate(x_fixturedate) Then Response.Write EW_FormatDateTime(x_fixturedate,7) Else Response.Write x_fixturedate End If %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">hometeam</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_hometeam" size="30" maxlength="30" value="<%= Server.HTMLEncode(x_hometeam&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">awayteam</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_awayteam" size="30" maxlength="60" value="<%= Server.HTMLEncode(x_awayteam&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">homescore</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_homescore" size="30" value="<%= Server.HTMLEncode(x_homescore&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">awayscore</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_awayscore" size="30" value="<%= Server.HTMLEncode(x_awayscore&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">fixtureyear</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_fixtureyear" size="30" value="<%= Server.HTMLEncode(x_fixtureyear&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">fixturenote</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_fixturenote" size="30" maxlength="100" value="<%= Server.HTMLEncode(x_fixturenote&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">fixtureid</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_fixtureid %><input type="hidden" name="x_fixtureid" value="<%= x_fixtureid %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">matchreport</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_matchreport" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_matchreport&"") %>"></span>&nbsp;</td>
	</tr>
</table>
<p>
<input type="submit" name="Action" value="EDIT">
</form>
<!--#include file="footer.asp"-->
