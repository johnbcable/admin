<!--#include file="db.asp"-->
<%
If Request.Form("submit") <> "" Then
	validpwd = False

	' Setup variables
	userid = Request.Form("userid")
	passwd = Request.Form("passwd")
    If Not validpwd Then
			Set conn = Server.CreateObject("ADODB.Connection")
			conn.open xDb_Conn_Str
			Set rs = conn.Execute( "Select * from [members] where [memberid] = '" & userid & "'")
			If Not rs.Eof Then
				If UCase(rs("webpassword")) = UCase(passwd) Then
					Session("project1_status_User") = rs("memberid")
					validpwd = True
				End If
			End If
			rs.Close
			Set rs = Nothing
			conn.Close
			Set conn = Nothing
	End If
	If validpwd Then

		' Write cookies
		If Request.Form("rememberme") <> "" Then
			Response.Cookies("project1")("userid") = userid
			Response.Cookies("project1").Expires = Date + 365 ' Change the expiry date of the cookies here
		End If		
		Session("project1_status") = "login"
		Response.Redirect "default.asp"
	End If
Else
	validpwd = True
End If
%>
<html>
<head>
	<title></title>
	<style type="text/css">
	<!-- 
 	INPUT, TEXTAREA, SELECT {font-family: Verdana; font-size: x-small;}
	.aspmaker {font-family: Verdana; font-size: x-small;}
	.ewTableOrderIndicator {font-family: Webdings;}
	-->
	</style>
<link href="tennis.css" rel="stylesheet" type="text/css" />
<meta name="generator" content="ASPMaker v3.1.0.1" />
</head>
<script language="JavaScript" src="ew.js"></script>
<script language="JavaScript">
<!-- start JavaScript
function  EW_checkMyForm(EW_this) {
if  (!EW_hasValue(EW_this.userid, "TEXT" )) {
            if  (!EW_onError(EW_this, EW_this.userid, "TEXT", "Please enter user ID"))
                return false; 
        }
if  (!EW_hasValue(EW_this.passwd, "PASSWORD" )) {
            if  (!EW_onError(EW_this, EW_this.passwd, "PASSWORD", "Please enter password"))
                return false; 
        }
return true;
}
// end JavaScript -->
</script>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<table border="0" cellspacing="0" cellpadding="2" align="center">
	<tr>
		<td><span class="aspmaker"></span></td>
	</tr>
</table>
<% If Not validpwd Then %>
<p align="center"><span class="aspmaker" style="color: Red;">Incorrect user ID or password</span></p>
<% End If %>
<form action="login.asp" method="post" onSubmit="return EW_checkMyForm(this);">
<table border="0" cellspacing="0" cellpadding="4" align="center">
	<tr>
		<td><span class="aspmaker">User Name</span></td>
		<td><span class="aspmaker"><input type="text" name="userid" size="20" value="<%= request.Cookies("project1")("userid") %>"></span></td>
	</tr>
	<tr>
		<td><span class="aspmaker">Password</span></td>
		<td><span class="aspmaker"><input type="password" name="passwd" size="20"></span></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><span class="aspmaker"><input type="checkbox" name="rememberme" value="true">Remember me</span></td>
	</tr>	
	<tr>
		<td colspan="2" align="center"><span class="aspmaker"><input type="submit" name="submit" value="Login"></span></td>
	</tr>
</table>
</form>
<br>
</body>
</html>
