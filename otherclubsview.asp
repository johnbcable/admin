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
End Select
%>
<!--#include file="header.asp"-->
<p><span class="aspmaker">View TABLE: otherclubs<br><br><a href="otherclubslist.asp">Back to List</a></span></p>
<p>
<form>
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#CCCCCC">
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">clubid</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_clubid %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">clubname</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_clubname %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">address 1</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_address1 %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">address 2</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_address2 %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">address 3</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_address3 %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">address 4</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_address4 %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">town</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_town %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">county</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_county %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">postcode</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_postcode %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">telephone</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_telephone %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">fax</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_fax %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">email</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_email %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">contact</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_contact %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">totaltennis</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_totaltennis %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">cluburl</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><% Response.Write x_cluburl %></span>&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#0099CC"><span class="aspmaker" style="color: #FFFFFF;">comments</span>&nbsp;</td>
		<td bgcolor="#F5F5F5"><span class="aspmaker"><%= Replace(x_comments & "", vbLf, "<br>") %></span>&nbsp;</td>
	</tr>
</table>
</form>
<p>
<!--#include file="footer.asp"-->
