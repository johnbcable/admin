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
If key = "" Or IsNull(key) Then Response.Redirect "circulation_listslist.asp"

' Get action
a = Request.Form("a")
If a = "" Or IsNull(a) Then
	a = "I"	' Display with input box
End If

' Get fields from form
x_name = Request.Form("x_name")
x_queryname = Request.Form("x_queryname")

' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Select Case a
	Case "I": ' Get a record to display
		tkey = "'" & key & "'"
		strsql = "SELECT * FROM [circulation_lists] WHERE [name]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.Eof Then
			Response.Clear
			Response.Redirect "circulation_listslist.asp"
		Else
			rs.MoveFirst
		End If

		' Get the field contents
		x_name = rs("name")
		x_queryname = rs("queryname")
		rs.Close
		Set rs = Nothing
	Case "U": ' Update

		' Open record
		tkey = "'" & key & "'"
		strsql = "SELECT * FROM [circulation_lists] WHERE [name]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn, 1, 2
		If rs.Eof Then
			Response.Clear
			Response.Redirect "circulation_listslist.asp"
		End If
		tmpFld = Trim(x_name)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("name") = tmpFld
		tmpFld = Trim(x_queryname)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("queryname") = tmpFld
		rs.Update
		rs.Close
		Set rs = Nothing
		conn.Close
		Set conn = Nothing
		Response.Clear
		Response.Redirect "circulation_listslist.asp"
End Select
%>
<!--#include file="header.asp"-->
<p><span class="aspmaker">Edit TABLE: circulation lists<br><br><a href="circulation_listslist.asp">Back to List</a></span></p>
<script language="JavaScript" src="ew.js"></script>
<script language="JavaScript">
<!-- start Javascript
function  EW_checkMyForm(EW_this) {
return true;
}
// end JavaScript -->
</script>
<form onSubmit="return EW_checkMyForm(this);"  name="circulation_listsedit" action="circulation_listsedit.asp" method="post">
<p>
<input type="hidden" name="a" value="U">
<input type="hidden" name="key" value="<%= key %>">
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#CCCCCC">
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">name</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_name" size="30" maxlength="150" value="<%= Server.HTMLEncode(x_name&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">queryname</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_queryname" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_queryname&"") %>"></span>&nbsp;</td>
	</tr>
</table>
<p>
<input type="submit" name="Action" value="EDIT">
</form>
<!--#include file="footer.asp"-->
