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
If key = "" Or IsNull(key) Then Response.Redirect "otherclubslist.asp"

' Get action
a = Request.Form("a")
If a = "" Or IsNull(a) Then
	a = "I"	' Display with input box
End If

' Get fields from form
x_clubid = Request.Form("x_clubid")
x_clubname = Request.Form("x_clubname")
x_address1 = Request.Form("x_address1")
x_address2 = Request.Form("x_address2")
x_address3 = Request.Form("x_address3")
x_address4 = Request.Form("x_address4")
x_town = Request.Form("x_town")
x_county = Request.Form("x_county")
x_postcode = Request.Form("x_postcode")
x_telephone = Request.Form("x_telephone")
x_fax = Request.Form("x_fax")
x_email = Request.Form("x_email")
x_contact = Request.Form("x_contact")
x_totaltennis = Request.Form("x_totaltennis")
x_cluburl = Request.Form("x_cluburl")
x_comments = Request.Form("x_comments")

' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Select Case a
	Case "I": ' Get a record to display
		tkey = "" & key & ""
		strsql = "SELECT * FROM [otherclubs] WHERE [clubid]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.Eof Then
			Response.Clear
			Response.Redirect "otherclubslist.asp"
		Else
			rs.MoveFirst
		End If

		' Get the field contents
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
		rs.Close
		Set rs = Nothing
	Case "U": ' Update

		' Open record
		tkey = "" & key & ""
		strsql = "SELECT * FROM [otherclubs] WHERE [clubid]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn, 1, 2
		If rs.Eof Then
			Response.Clear
			Response.Redirect "otherclubslist.asp"
		End If
		tmpFld = Trim(x_clubname)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("clubname") = tmpFld
		tmpFld = Trim(x_address1)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("address1") = tmpFld
		tmpFld = Trim(x_address2)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("address2") = tmpFld
		tmpFld = Trim(x_address3)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("address3") = tmpFld
		tmpFld = Trim(x_address4)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("address4") = tmpFld
		tmpFld = Trim(x_town)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("town") = tmpFld
		tmpFld = Trim(x_county)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("county") = tmpFld
		tmpFld = Trim(x_postcode)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("postcode") = tmpFld
		tmpFld = Trim(x_telephone)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("telephone") = tmpFld
		tmpFld = Trim(x_fax)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("fax") = tmpFld
		tmpFld = Trim(x_email)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("email") = tmpFld
		tmpFld = Trim(x_contact)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("contact") = tmpFld
		tmpFld = Trim(x_totaltennis)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("totaltennis") = tmpFld
		tmpFld = Trim(x_cluburl)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("cluburl") = tmpFld
		tmpFld = Trim(x_comments)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("comments") = tmpFld
		rs.Update
		rs.Close
		Set rs = Nothing
		conn.Close
		Set conn = Nothing
		Response.Clear
		Response.Redirect "otherclubslist.asp"
End Select
%>
<!--#include file="header.asp"-->
<p><span class="aspmaker">Edit TABLE: otherclubs<br><br><a href="otherclubslist.asp">Back to List</a></span></p>
<script language="JavaScript" src="ew.js"></script>
<script language="JavaScript">
<!-- start Javascript
function  EW_checkMyForm(EW_this) {
return true;
}
// end JavaScript -->
</script>
<form onSubmit="return EW_checkMyForm(this);"  name="otherclubsedit" action="otherclubsedit.asp" method="post">
<p>
<input type="hidden" name="a" value="U">
<input type="hidden" name="key" value="<%= key %>">
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#CCCCCC">
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">clubid</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_clubid %><input type="hidden" name="x_clubid" value="<%= x_clubid %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">clubname</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_clubname" size="30" maxlength="60" value="<%= Server.HTMLEncode(x_clubname&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">address 1</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_address1" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_address1&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">address 2</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_address2" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_address2&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">address 3</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_address3" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_address3&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">address 4</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_address4" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_address4&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">town</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_town" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_town&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">county</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_county" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_county&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">postcode</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_postcode" size="30" maxlength="12" value="<%= Server.HTMLEncode(x_postcode&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">telephone</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_telephone" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_telephone&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">fax</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_fax" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_fax&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">email</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_email" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_email&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">contact</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_contact" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_contact&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">totaltennis</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_totaltennis" size="30" maxlength="100" value="<%= Server.HTMLEncode(x_totaltennis&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">cluburl</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_cluburl" size="30" maxlength="100" value="<%= Server.HTMLEncode(x_cluburl&"") %>"></span>&nbsp;</td>
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
