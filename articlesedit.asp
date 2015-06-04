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
If key = "" Or IsNull(key) Then Response.Redirect "articleslist.asp"

' Get action
a = Request.Form("a")
If a = "" Or IsNull(a) Then
	a = "I"	' Display with input box
End If

' Get fields from form
x_articleid = Request.Form("x_articleid")
x_title = Request.Form("x_title")
x_author = Request.Form("x_author")
x_courtcircularissueno = Request.Form("x_courtcircularissueno")
x_courtseq = Request.Form("x_courtseq")
x_courtcols = Request.Form("x_courtcols")
x_articletext = Request.Form("x_articletext")
x_newsitem = Request.Form("x_newsitem")
x_newsfrom = Request.Form("x_newsfrom")
x_newsuntil = Request.Form("x_newsuntil")
x_newspriority = Request.Form("x_newspriority")
x_articlesection = Request.Form("x_articlesection")

' Open Connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Select Case a
	Case "I": ' Get a record to display
		tkey = "" & key & ""
		strsql = "SELECT * FROM [articles] WHERE [articleid]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.Eof Then
			Response.Clear
			Response.Redirect "articleslist.asp"
		Else
			rs.MoveFirst
		End If

		' Get the field contents
		x_articleid = rs("articleid")
		x_title = rs("title")
		x_author = rs("author")
		x_courtcircularissueno = rs("courtcircularissueno")
		x_courtseq = rs("courtseq")
		x_courtcols = rs("courtcols")
		x_articletext = rs("articletext")
		x_newsitem = rs("newsitem")
		x_newsfrom = rs("newsfrom")
		x_newsuntil = rs("newsuntil")
		x_newspriority = rs("newspriority")
		x_articlesection = rs("articlesection")
		rs.Close
		Set rs = Nothing
	Case "U": ' Update

		' Open record
		tkey = "" & key & ""
		strsql = "SELECT * FROM [articles] WHERE [articleid]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn, 1, 2
		If rs.Eof Then
			Response.Clear
			Response.Redirect "articleslist.asp"
		End If
		tmpFld = Trim(x_title)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("title") = tmpFld
		tmpFld = Trim(x_author)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("author") = tmpFld
		tmpFld = x_courtcircularissueno
		If Not IsNumeric(tmpFld) Then tmpFld = 0
		rs("courtcircularissueno") = CLng(tmpFld)
		tmpFld = x_courtseq
		If Not IsNumeric(tmpFld) Then tmpFld = 0
		rs("courtseq") = CLng(tmpFld)
		tmpFld = x_courtcols
		If Not IsNumeric(tmpFld) Then tmpFld = 0
		rs("courtcols") = CLng(tmpFld)
		tmpFld = Trim(x_articletext)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("articletext") = tmpFld
		tmpFld = Trim(x_newsitem)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("newsitem") = tmpFld
		tmpFld = EW_UnFormatDateTime(x_newsfrom,7)
		If IsDate(tmpFld) Then
		    rs("newsfrom") = CDate(tmpFld)
		Else
		    rs("newsfrom") = Null
		End If
		tmpFld = EW_UnFormatDateTime(x_newsuntil,7)
		If IsDate(tmpFld) Then
		    rs("newsuntil") = CDate(tmpFld)
		Else
		    rs("newsuntil") = Null
		End If
		tmpFld = x_newspriority
		If Not IsNumeric(tmpFld) Then tmpFld = 0
		rs("newspriority") = CLng(tmpFld)
		tmpFld = Trim(x_articlesection)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("articlesection") = tmpFld
		rs.Update
		rs.Close
		Set rs = Nothing
		conn.Close
		Set conn = Nothing
		Response.Clear
		Response.Redirect "articleslist.asp"
End Select
%>
<!--#include file="header.asp"-->
<p><span class="aspmaker">Edit TABLE: articles<br><br><a href="articleslist.asp">Back to List</a></span></p>
<script language="JavaScript" src="ew.js"></script>
<script language="JavaScript">
<!-- start Javascript
function  EW_checkMyForm(EW_this) {
if (EW_this.x_courtcircularissueno && !EW_checkinteger(EW_this.x_courtcircularissueno.value)) {
        if (!EW_onError(EW_this, EW_this.x_courtcircularissueno, "TEXT", "Incorrect integer - courtcircularissueno"))
            return false; 
        }
if (EW_this.x_courtseq && !EW_checkinteger(EW_this.x_courtseq.value)) {
        if (!EW_onError(EW_this, EW_this.x_courtseq, "TEXT", "Incorrect integer - courtseq"))
            return false; 
        }
if (EW_this.x_courtcols && !EW_checkinteger(EW_this.x_courtcols.value)) {
        if (!EW_onError(EW_this, EW_this.x_courtcols, "TEXT", "Incorrect integer - courtcols"))
            return false; 
        }
if (EW_this.x_newsfrom && !EW_checkeurodate(EW_this.x_newsfrom.value)) {
        if (!EW_onError(EW_this, EW_this.x_newsfrom, "TEXT", "Incorrect date (dd/mm/yyyy) - newsfrom"))
            return false; 
        }
if (EW_this.x_newsuntil && !EW_checkeurodate(EW_this.x_newsuntil.value)) {
        if (!EW_onError(EW_this, EW_this.x_newsuntil, "TEXT", "Incorrect date (dd/mm/yyyy) - newsuntil"))
            return false; 
        }
if (EW_this.x_newspriority && !EW_checkinteger(EW_this.x_newspriority.value)) {
        if (!EW_onError(EW_this, EW_this.x_newspriority, "TEXT", "Incorrect integer - newspriority"))
            return false; 
        }
return true;
}
// end JavaScript -->
</script>
<form onSubmit="return EW_checkMyForm(this);"  name="articlesedit" action="articlesedit.asp" method="post">
<p>
<input type="hidden" name="a" value="U">
<input type="hidden" name="key" value="<%= key %>">
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#CCCCCC">
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">articleid</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_articleid %><input type="hidden" name="x_articleid" value="<%= x_articleid %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">title</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_title" size="30" maxlength="150" value="<%= Server.HTMLEncode(x_title&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">author</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_author" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_author&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">courtcircularissueno</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_courtcircularissueno" size="30" value="<%= Server.HTMLEncode(x_courtcircularissueno&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">courtseq</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_courtseq" size="30" value="<%= Server.HTMLEncode(x_courtseq&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">courtcols</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_courtcols" size="30" value="<%= Server.HTMLEncode(x_courtcols&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">articletext</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><textarea cols=35 rows=4 name="x_articletext"><%= x_articletext %></textarea></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">newsitem</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_newsitem" size="30" maxlength="1" value="<%= Server.HTMLEncode(x_newsitem&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">newsfrom</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_newsfrom" value="<% If IsDate(x_newsfrom) Then Response.Write EW_FormatDateTime(x_newsfrom,7) Else Response.Write x_newsfrom End If %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">newsuntil</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_newsuntil" value="<% If IsDate(x_newsuntil) Then Response.Write EW_FormatDateTime(x_newsuntil,7) Else Response.Write x_newsuntil End If %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">newspriority</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_newspriority" size="30" value="<%= Server.HTMLEncode(x_newspriority&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">articlesection</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_articlesection" size="30" maxlength="20" value="<%= Server.HTMLEncode(x_articlesection&"") %>"></span>&nbsp;</td>
	</tr>
</table>
<p>
<input type="submit" name="Action" value="EDIT">
</form>
<!--#include file="footer.asp"-->
