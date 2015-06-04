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
		strsql = "SELECT * FROM [members] WHERE [memberid]=" & tkey
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn
		If rs.Eof Then
			Response.Clear
			Response.Redirect "memberslist.asp"
		Else
			rs.MoveFirst

			' Get the field contents
			x_memberid = rs("memberid")
			x_membergrade = rs("membergrade")
			x_surname = rs("surname")
			x_forename1 = rs("forename1")
			x_initials = rs("initials")
			x_title = rs("title")
			x_address1 = rs("address1")
			x_address2 = rs("address2")
			x_address3 = rs("address3")
			x_address4 = rs("address4")
			x_postcode = rs("postcode")
			x_homephone = rs("homephone")
			x_workphone = rs("workphone")
			x_mobilephone = rs("mobilephone")
			x_email = rs("email")
			x_webpassword = rs("webpassword")
			x_webaccess = rs("webaccess")
			x_dob = rs("dob")
			x_mailing = rs("mailing")
			x_joined = rs("joined")
			x_dateleft = rs("dateleft")
			x_gender = rs("gender")
		End If
		rs.Close
		Set rs = Nothing
	Case "A": ' Add

		' Get fields from form
x_memberid = Request.Form("x_memberid")
x_membergrade = Request.Form("x_membergrade")
x_surname = Request.Form("x_surname")
x_forename1 = Request.Form("x_forename1")
x_initials = Request.Form("x_initials")
x_title = Request.Form("x_title")
x_address1 = Request.Form("x_address1")
x_address2 = Request.Form("x_address2")
x_address3 = Request.Form("x_address3")
x_address4 = Request.Form("x_address4")
x_postcode = Request.Form("x_postcode")
x_homephone = Request.Form("x_homephone")
x_workphone = Request.Form("x_workphone")
x_mobilephone = Request.Form("x_mobilephone")
x_email = Request.Form("x_email")
x_webpassword = Request.Form("x_webpassword")
x_webaccess = Request.Form("x_webaccess")
x_uniqueref = Request.Form("x_uniqueref")
x_dob = Request.Form("x_dob")
x_mailing = Request.Form("x_mailing")
x_joined = Request.Form("x_joined")
x_dateleft = Request.Form("x_dateleft")
x_gender = Request.Form("x_gender")

		' Open record
		strsql = "SELECT * FROM [members] WHERE 0 = 1"
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn, 1, 2
		rs.AddNew
		tmpFld = Trim(x_memberid)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		srchFld = Replace(tmpFld&"","'","''")
		srchFld = Replace(srchFld,"[","[[]")
		srchFld = "'" & srchFld & "'"
		strsql = "SELECT * FROM [members] WHERE [memberid] = " & srchFld
		Set rschk = conn.Execute(strsql)
		If Not rschk.Eof Then
		  Response.Write "Duplicate value for index or primary key -- memberid, value = " & tmpFld & "<br>"
		  Response.Write "Press [Previous Page] key to continue!"
		  Response.End
		End If
		rschk.Close
		Set rschk = Nothing
		rs("memberid") = tmpFld
		tmpFld = Trim(x_membergrade)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("membergrade") = tmpFld
		tmpFld = Trim(x_surname)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("surname") = tmpFld
		tmpFld = Trim(x_forename1)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("forename1") = tmpFld
		tmpFld = Trim(x_initials)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("initials") = tmpFld
		tmpFld = Trim(x_title)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("title") = tmpFld
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
		tmpFld = Trim(x_postcode)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("postcode") = tmpFld
		tmpFld = Trim(x_homephone)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("homephone") = tmpFld
		tmpFld = Trim(x_workphone)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("workphone") = tmpFld
		tmpFld = Trim(x_mobilephone)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("mobilephone") = tmpFld
		tmpFld = Trim(x_email)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("email") = tmpFld
		tmpFld = Trim(x_webpassword)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("webpassword") = tmpFld
		tmpFld = x_webaccess
		If Not IsNumeric(tmpFld) Then tmpFld = 0
		rs("webaccess") = CLng(tmpFld)
		tmpFld = EW_UnFormatDateTime(x_dob,7)
		If IsDate(tmpFld) Then
		    rs("dob") = CDate(tmpFld)
		Else
		    rs("dob") = Null
		End If
		tmpFld = Trim(x_mailing)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("mailing") = tmpFld
		tmpFld = EW_UnFormatDateTime(x_joined,7)
		If IsDate(tmpFld) Then
		    rs("joined") = CDate(tmpFld)
		Else
		    rs("joined") = Null
		End If
		tmpFld = EW_UnFormatDateTime(x_dateleft,7)
		If IsDate(tmpFld) Then
		    rs("dateleft") = CDate(tmpFld)
		Else
		    rs("dateleft") = Null
		End If
		tmpFld = Trim(x_gender)
		If Trim(tmpFld) & "x" = "x" Then tmpFld = Null
		rs("gender") = tmpFld
		rs.Update
		rs.Close
		Set rs = Nothing
		conn.Close
		Set conn = Nothing
		Response.Clear
		Response.Redirect "memberslist.asp"
End Select
%>
<!--#include file="header.asp"-->
<p><span class="aspmaker">Add to TABLE: members<br><br><a href="memberslist.asp">Back to List</a></span></p>
<script language="JavaScript" src="ew.js"></script>
<script language="JavaScript">
<!-- start Javascript
function  EW_checkMyForm(EW_this) {
if (EW_this.x_webaccess && !EW_checkinteger(EW_this.x_webaccess.value)) {
        if (!EW_onError(EW_this, EW_this.x_webaccess, "TEXT", "Incorrect integer - webaccess"))
            return false; 
        }
if (EW_this.x_dob && !EW_checkeurodate(EW_this.x_dob.value)) {
        if (!EW_onError(EW_this, EW_this.x_dob, "TEXT", "Incorrect date (dd/mm/yyyy) - dob"))
            return false; 
        }
if (EW_this.x_joined && !EW_checkeurodate(EW_this.x_joined.value)) {
        if (!EW_onError(EW_this, EW_this.x_joined, "TEXT", "Incorrect date (dd/mm/yyyy) - joined"))
            return false; 
        }
if (EW_this.x_dateleft && !EW_checkeurodate(EW_this.x_dateleft.value)) {
        if (!EW_onError(EW_this, EW_this.x_dateleft, "TEXT", "Incorrect date (dd/mm/yyyy) - dateleft"))
            return false; 
        }
return true;
}
// end JavaScript -->
</script>
<form onSubmit="return EW_checkMyForm(this);"  action="membersadd.asp" method="post">
<p>
<input type="hidden" name="a" value="A">
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#CCCCCC">
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">memberid</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_memberid" size="30" maxlength="10" value="<%= Server.HTMLEncode(x_memberid&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">membergrade</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_membergrade" size="30" maxlength="20" value="<%= Server.HTMLEncode(x_membergrade&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">surname</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_surname" size="30" maxlength="35" value="<%= Server.HTMLEncode(x_surname&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">forename 1</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_forename1" size="30" maxlength="20" value="<%= Server.HTMLEncode(x_forename1&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">initials</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_initials" size="30" maxlength="5" value="<%= Server.HTMLEncode(x_initials&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">title</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_title" size="30" maxlength="6" value="<%= Server.HTMLEncode(x_title&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">address 1</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_address1" size="30" maxlength="35" value="<%= Server.HTMLEncode(x_address1&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">address 2</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_address2" size="30" maxlength="35" value="<%= Server.HTMLEncode(x_address2&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">address 3</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_address3" size="30" maxlength="35" value="<%= Server.HTMLEncode(x_address3&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">address 4</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_address4" size="30" maxlength="35" value="<%= Server.HTMLEncode(x_address4&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">postcode</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_postcode" size="30" maxlength="12" value="<%= Server.HTMLEncode(x_postcode&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">homephone</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_homephone" size="30" maxlength="20" value="<%= Server.HTMLEncode(x_homephone&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">workphone</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_workphone" size="30" maxlength="20" value="<%= Server.HTMLEncode(x_workphone&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">mobilephone</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_mobilephone" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_mobilephone&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">email</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_email" size="30" maxlength="100" value="<%= Server.HTMLEncode(x_email&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">webpassword</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_webpassword" size="30" maxlength="10" value="<%= Server.HTMLEncode(x_webpassword&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">webaccess</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_webaccess" size="30" value="<%= Server.HTMLEncode(x_webaccess&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">dob</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_dob" value="<% If IsDate(x_dob) Then Response.Write EW_FormatDateTime(x_dob,7) Else Response.Write x_dob End If %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">mailing</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_mailing" size="30" maxlength="1" value="<%= Server.HTMLEncode(x_mailing&"") %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">joined</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_joined" value="<% If IsDate(x_joined) Then Response.Write EW_FormatDateTime(x_joined,7) Else Response.Write x_joined End If %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">dateleft</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_dateleft" value="<% If IsDate(x_dateleft) Then Response.Write EW_FormatDateTime(x_dateleft,7) Else Response.Write x_dateleft End If %>"></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">gender</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><input type="text" name="x_gender" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_gender&"") %>"></span>&nbsp;</td>
	</tr>
</table>
<p>
<input type="submit" name="Action" value="ADD">
</form>
<!--#include file="footer.asp"-->
