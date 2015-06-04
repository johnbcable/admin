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
		tkey = "'" & key & "'"
		strsql = "SELECT * FROM [circulation_lists] WHERE [name]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.Eof Then
			Response.Clear
			Response.Redirect "circulation_listslist.asp"
		Else
			rs.MoveFirst

			' Get the field contents
			x_name = rs("name")
			x_queryname = rs("queryname")
		End If
		rs.Close
		Set rs = Nothing
	Case "A": ' Add

		' Get fields from form
x_name = Request.Form("x_name")
x_queryname = Request.Form("x_queryname")

		' Open record
		strsql = "SELECT * FROM [circulation_lists] WHERE 0 = 1"
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn, 1, 2
		rs.AddNew
		tmpFld = Trim(x_name)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		srchFld = Replace(tmpFld&"","'","''")
		srchFld = Replace(srchFld,"[","[[]")
		srchFld = "'" & srchFld & "'"
		strsql = "SELECT * FROM [circulation_lists] WHERE [name] = " & srchFld
		Set rschk = conn.Execute(strsql)
		If Not rschk.Eof Then
		  Response.Write "Duplicate value for index or primary key -- name, value = " & tmpFld & "<br>"
		  Response.Write "Press [Previous Page] key to continue!"
		  Response.End
		End If
		rschk.Close
		Set rschk = Nothing
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
<p><span class="aspmaker">Add to TABLE: circulation lists<br><br><a href="circulation_listslist.asp">Back to List</a></span></p>
<script language="JavaScript" src="ew.js"></script>
<script language="JavaScript">
<!-- start Javascript
function  EW_checkMyForm(EW_this) {
return true;
}
// end JavaScript -->
</script>
<form onSubmit="return EW_checkMyForm(this);"  action="circulation_listsadd.asp" method="post">
<p>
<input type="hidden" name="a" value="A">
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
<input type="submit" name="Action" value="ADD">
</form>
<!--#include file="footer.asp"-->
