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
		strsql = "SELECT * FROM [photos] WHERE [mainphoto]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.Eof Then
			Response.Clear
			Response.Redirect "photoslist.asp"
		Else
			rs.MoveFirst

			' Get the field contents
			x_mainphoto = rs("mainphoto")
			x_thumbnail = rs("thumbnail")
			x_bottomcaption = rs("bottomcaption")
			x_topcaption = rs("topcaption")
			x_altcaption = rs("altcaption")
			x_takenon = rs("takenon")
			x_submittedby = rs("submittedby")
		End If
		rs.Close
		Set rs = Nothing
	Case "A": ' Add

		' Get fields from form
x_photoid = Request.Form("x_photoid")
x_mainphoto = Request.Form("x_mainphoto")
x_thumbnail = Request.Form("x_thumbnail")
x_bottomcaption = Request.Form("x_bottomcaption")
x_topcaption = Request.Form("x_topcaption")
x_altcaption = Request.Form("x_altcaption")
x_takenon = Request.Form("x_takenon")
x_submittedby = Request.Form("x_submittedby")

		' Open record
		strsql = "SELECT * FROM [photos] WHERE 0 = 1"
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn, 1, 2
		rs.AddNew
		tmpFld = Trim(x_mainphoto)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		srchFld = Replace(tmpFld&"","'","''")
		srchFld = Replace(srchFld,"[","[[]")
		srchFld = "'" & srchFld & "'"
		strsql = "SELECT * FROM [photos] WHERE [mainphoto] = " & srchFld
		Set rschk = conn.Execute(strsql)
		If Not rschk.Eof Then
		  Response.Write "Duplicate value for index or primary key -- mainphoto, value = " & tmpFld & "<br>"
		  Response.Write "Press [Previous Page] key to continue!"
		  Response.End
		End If
		rschk.Close
		Set rschk = Nothing
		rs("mainphoto") = tmpFld
		tmpFld = Trim(x_thumbnail)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("thumbnail") = tmpFld
		tmpFld = Trim(x_bottomcaption)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("bottomcaption") = tmpFld
		tmpFld = Trim(x_topcaption)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("topcaption") = tmpFld
		tmpFld = Trim(x_altcaption)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("altcaption") = tmpFld
		tmpFld = EW_UnFormatDateTime(x_takenon,7)
		If IsDate(tmpFld) Then
		    rs("takenon") = CDate(tmpFld)
		Else
		    rs("takenon") = Null
		End If
		tmpFld = Trim(x_submittedby)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("submittedby") = tmpFld
		rs.Update
		rs.Close
		Set rs = Nothing
		conn.Close
		Set conn = Nothing
		Response.Clear
		Response.Redirect "photoslist.asp"
End Select
%>
<!--#include file="header.asp"-->
<p><span class="aspmaker">Add to TABLE: photos<br><br><a href="photoslist.asp">Back to List</a></span></p>
<script language="JavaScript" src="ew.js"></script>
<script language="JavaScript">
<!-- start Javascript
function  EW_checkMyForm(EW_this) {
if (EW_this.x_takenon && !EW_checkeurodate(EW_this.x_takenon.value)) {
        if (!EW_onError(EW_this, EW_this.x_takenon, "TEXT", "Incorrect date (dd/mm/yyyy) - takenon"))
            return false; 
        }
return true;
}
// end JavaScript -->
</script>
<form onSubmit="return EW_checkMyForm(this);"  action="photosadd.asp" method="post">
<p>
<input type="hidden" name="a" value="A">
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#CCCCCC">
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">mainphoto</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_mainphoto" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_mainphoto&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">thumbnail</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_thumbnail" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_thumbnail&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">bottomcaption</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_bottomcaption" size="30" maxlength="100" value="<%= Server.HTMLEncode(x_bottomcaption&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">topcaption</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_topcaption" size="30" maxlength="100" value="<%= Server.HTMLEncode(x_topcaption&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">altcaption</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_altcaption" size="30" maxlength="100" value="<%= Server.HTMLEncode(x_altcaption&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">takenon</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_takenon" value="<% If IsDate(x_takenon) Then Response.Write EW_FormatDateTime(x_takenon,7) Else Response.Write x_takenon End If %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">submittedby</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_submittedby" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_submittedby&"") %>"></span>&nbsp;</td>
	</tr>
</table>
<p>
<input type="submit" name="Action" value="ADD">
</form>
<!--#include file="footer.asp"-->
