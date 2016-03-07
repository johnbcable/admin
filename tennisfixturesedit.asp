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
If key = "" Or IsNull(key) Then Response.Redirect "tennisfixtureslist.asp"

' Get action
a = Request.Form("a")
If a = "" Or IsNull(a) Then
	a = "I"	' Display with input box
End If

' Get fields from form
x_fixturedate = Request.Form("x_fixturedate")
x_homeoraway = Request.Form("x_homeoraway")
x_opponents = Request.Form("x_opponents")
x_hamptonresult = Request.Form("x_hamptonresult")
x_opponentresult = Request.Form("x_opponentresult")
x_fixtureyear = Request.Form("x_fixtureyear")
x_teamname = Request.Form("x_teamname")
x_fixturenote = Request.Form("x_fixturenote")
x_fixtureid = Request.Form("x_fixtureid")
x_matchreport = Request.Form("x_matchreport")
x_pair1 = Request.Form("x_pair1")
x_pair2 = Request.Form("x_pair2")

' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Select Case a
	Case "I": ' Get a record to display
		tkey = "" & key & ""
		strsql = "SELECT * FROM [tennisfixtures] WHERE [fixtureid]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.Eof Then
			Response.Clear
			Response.Redirect "tennisfixtureslist.asp"
		Else
			rs.MoveFirst
		End If

		' Get the field contents
		x_fixturedate = rs("fixturedate")
		x_homeoraway = rs("homeoraway")
		x_opponents = rs("opponents")
		x_hamptonresult = rs("hamptonresult")
		x_opponentresult = rs("opponentresult")
		x_fixtureyear = rs("fixtureyear")
		x_teamname = rs("teamname")
		x_fixturenote = rs("fixturenote")
		x_fixtureid = rs("fixtureid")
		x_matchreport = rs("matchreport")
		x_pair1 = rs("pair1")
		x_pair2 = rs("pair2")
		rs.Close
		Set rs = Nothing
	Case "U": ' Update

		' Open record
		tkey = "" & key & ""
		strsql = "SELECT * FROM [tennisfixtures] WHERE [fixtureid]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn, 1, 2
		If rs.Eof Then
			Response.Clear
			Response.Redirect "tennisfixtureslist.asp"
		End If
		tmpFld = EW_UnFormatDateTime(x_fixturedate,7)
		If IsDate(tmpFld) Then
		    rs("fixturedate") = CDate(tmpFld)
		Else
		    rs("fixturedate") = Null
		End If
		tmpFld = Trim(x_homeoraway)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("homeoraway") = tmpFld
		tmpFld = Trim(x_opponents)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("opponents") = tmpFld
		tmpFld = x_hamptonresult
		If Not IsNumeric(tmpFld) Then tmpFld = 0
		rs("hamptonresult") = CLng(tmpFld)
		tmpFld = x_opponentresult
		If Not IsNumeric(tmpFld) Then tmpFld = 0
		rs("opponentresult") = CLng(tmpFld)
		tmpFld = x_fixtureyear
		If Not IsNumeric(tmpFld) Then tmpFld = 0
		rs("fixtureyear") = CLng(tmpFld)
		tmpFld = Trim(x_teamname)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("teamname") = tmpFld
		tmpFld = Trim(x_fixturenote)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("fixturenote") = tmpFld
		tmpFld = Trim(x_matchreport)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("matchreport") = tmpFld
		tmpFld = Trim(x_pair1)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("pair1") = tmpFld
		tmpFld = Trim(x_pair2)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("pair2") = tmpFld
		rs.Update
		rs.Close
		Set rs = Nothing
		conn.Close
		Set conn = Nothing
		Response.Clear
		Response.Redirect "tennisfixtureslist.asp"
End Select
%>
<!--#include file="header.asp"-->
<p><span class="aspmaker">Edit TABLE: tennisfixtures<br><br><a href="tennisfixtureslist.asp">Back to List</a></span></p>
<script language="JavaScript" src="ew.js"></script>
<script language="JavaScript">
<!-- start Javascript
function  EW_checkMyForm(EW_this) {
if (EW_this.x_fixturedate && !EW_checkeurodate(EW_this.x_fixturedate.value)) {
        if (!EW_onError(EW_this, EW_this.x_fixturedate, "TEXT", "Incorrect date (dd/mm/yyyy) - fixturedate"))
            return false; 
        }
if (EW_this.x_hamptonresult && !EW_checkinteger(EW_this.x_hamptonresult.value)) {
        if (!EW_onError(EW_this, EW_this.x_hamptonresult, "TEXT", "Incorrect integer - hamptonresult"))
            return false; 
        }
if (EW_this.x_opponentresult && !EW_checkinteger(EW_this.x_opponentresult.value)) {
        if (!EW_onError(EW_this, EW_this.x_opponentresult, "TEXT", "Incorrect integer - opponentresult"))
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
<form onSubmit="return EW_checkMyForm(this);"  name="tennisfixturesedit" action="tennisfixturesedit.asp" method="post">
<p>
<input type="hidden" name="a" value="U">
<input type="hidden" name="key" value="<%= key %>">
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#CCCCCC">
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">fixturedate</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_fixturedate" value="<% If IsDate(x_fixturedate) Then Response.Write EW_FormatDateTime(x_fixturedate,7) Else Response.Write x_fixturedate End If %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">homeoraway</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_homeoraway" size="30" maxlength="1" value="<%= Server.HTMLEncode(x_homeoraway&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">opponents</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_opponents" size="30" maxlength="60" value="<%= Server.HTMLEncode(x_opponents&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">hamptonresult</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_hamptonresult" size="30" value="<%= Server.HTMLEncode(x_hamptonresult&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">opponentresult</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_opponentresult" size="30" value="<%= Server.HTMLEncode(x_opponentresult&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">fixtureyear</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_fixtureyear" size="30" value="<%= Server.HTMLEncode(x_fixtureyear&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">teamname</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_teamname" size="30" maxlength="30" value="<%= Server.HTMLEncode(x_teamname&"") %>"></span>&nbsp;</td>
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
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">pair 1</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_pair1" size="30" maxlength="80" value="<%= Server.HTMLEncode(x_pair1&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">pair 2</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_pair2" size="30" maxlength="80" value="<%= Server.HTMLEncode(x_pair2&"") %>"></span>&nbsp;</td>
	</tr>
</table>
<p>
<input type="submit" name="Action" value="EDIT">
</form>
<!--#include file="footer.asp"-->
