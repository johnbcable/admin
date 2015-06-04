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
		strsql = "SELECT * FROM [emails] WHERE [emailid]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.Eof Then
			Response.Clear
			Response.Redirect "emailslist.asp"
		Else
			rs.MoveFirst

			' Get the field contents
			x_subject = rs("subject")
			x_emailfile = rs("emailfile")
			x_attach1 = rs("attach1")
			x_attach2 = rs("attach2")
			x_attach3 = rs("attach3")
			x_circulation = rs("circulation")
			x_sent_on = rs("sent_on")
			x_number_sent = rs("number_sent")
			x_comment = rs("comment")
		End If
		rs.Close
		Set rs = Nothing
	Case "A": ' Add

		' Get fields from form
x_emailid = Request.Form("x_emailid")
x_subject = Request.Form("x_subject")
x_emailfile = Request.Form("x_emailfile")
x_attach1 = Request.Form("x_attach1")
x_attach2 = Request.Form("x_attach2")
x_attach3 = Request.Form("x_attach3")
x_circulation = Request.Form("x_circulation")
x_sent_on = Request.Form("x_sent_on")
x_number_sent = Request.Form("x_number_sent")
x_comment = Request.Form("x_comment")

		' Open record
		strsql = "SELECT * FROM [emails] WHERE 0 = 1"
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn, 1, 2
		rs.AddNew
		tmpFld = Trim(x_subject)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("subject") = tmpFld
		tmpFld = Trim(x_emailfile)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("emailfile") = tmpFld
		tmpFld = Trim(x_attach1)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("attach1") = tmpFld
		tmpFld = Trim(x_attach2)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("attach2") = tmpFld
		tmpFld = Trim(x_attach3)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("attach3") = tmpFld
		tmpFld = Trim(x_circulation)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("circulation") = tmpFld
		tmpFld = EW_UnFormatDateTime(x_sent_on,7)
		If IsDate(tmpFld) Then
		    rs("sent_on") = CDate(tmpFld)
		Else
		    rs("sent_on") = Null
		End If
		tmpFld = x_number_sent
		If Not IsNumeric(tmpFld) Then tmpFld = 0
		rs("number_sent") = CLng(tmpFld)
		tmpFld = Trim(x_comment)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("comment") = tmpFld
		rs.Update
		rs.Close
		Set rs = Nothing
		conn.Close
		Set conn = Nothing
		Response.Clear
		Response.Redirect "emailslist.asp"
End Select
%>
<!--#include file="header.asp"-->
<p><span class="aspmaker">Add to TABLE: emails<br><br><a href="emailslist.asp">Back to List</a></span></p>
<script language="JavaScript" src="ew.js"></script>
<script language="JavaScript">
<!-- start Javascript
function  EW_checkMyForm(EW_this) {
if (EW_this.x_sent_on && !EW_checkeurodate(EW_this.x_sent_on.value)) {
        if (!EW_onError(EW_this, EW_this.x_sent_on, "TEXT", "Incorrect date (dd/mm/yyyy) - sent on"))
            return false; 
        }
if (EW_this.x_number_sent && !EW_checkinteger(EW_this.x_number_sent.value)) {
        if (!EW_onError(EW_this, EW_this.x_number_sent, "TEXT", "Incorrect integer - number sent"))
            return false; 
        }
return true;
}
// end JavaScript -->
</script>
<form onSubmit="return EW_checkMyForm(this);"  action="emailsadd.asp" method="post">
<p>
<input type="hidden" name="a" value="A">
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#CCCCCC">
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">subject</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_subject" size="30" maxlength="200" value="<%= Server.HTMLEncode(x_subject&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">emailfile</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_emailfile" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_emailfile&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">attach 1</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_attach1" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_attach1&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">attach 2</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_attach2" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_attach2&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">attach 3</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_attach3" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_attach3&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">circulation</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_circulation" size="30" maxlength="100" value="<%= Server.HTMLEncode(x_circulation&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">sent on</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_sent_on" value="<% If IsDate(x_sent_on) Then Response.Write EW_FormatDateTime(x_sent_on,7) Else Response.Write x_sent_on End If %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">number sent</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_number_sent" size="30" value="<%= Server.HTMLEncode(x_number_sent&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">comment</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_comment" size="30" maxlength="200" value="<%= Server.HTMLEncode(x_comment&"") %>"></span>&nbsp;</td>
	</tr>
</table>
<p>
<input type="submit" name="Action" value="ADD">
</form>
<!--#include file="footer.asp"-->
