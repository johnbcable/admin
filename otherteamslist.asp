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
displayRecs = 20
recRange = 10
%>
<%
dbwhere = ""
masterdetailwhere = ""
searchwhere = ""
a_search = ""
b_search = ""
whereClause = ""
%>
<%

' Get search criteria for basic search
pSearch = Request.QueryString("psearch")
pSearchType = Request.QueryString("psearchType")
If pSearch <> "" Then
	pSearch = Replace(pSearch,"'","''")
	pSearch = Replace(pSearch,"[","[[]")
	If pSearchType <> "" Then
		While InStr(pSearch, "  ") > 0
			pSearch = Replace(pSearch, "  ", " ")
		Wend
		arpSearch = Split(Trim(pSearch), " ")
		For Each kw In arpSearch
			b_search = b_search & "("
			b_search = b_search & "[clubname] LIKE '%" & Trim(kw) & "%' OR "
			b_search = b_search & "[teamname] LIKE '%" & Trim(kw) & "%' OR "
			b_search = b_search & "[comments] LIKE '%" & Trim(kw) & "%' OR "
			If Right(b_search, 4)=" OR " Then b_search = Left(b_search, Len(b_search)-4)
			b_search = b_search & ") " & pSearchType & " "
		Next
	Else
	b_search = b_search & "[clubname] LIKE '%" & pSearch & "%' OR "
	b_search = b_search & "[teamname] LIKE '%" & pSearch & "%' OR "
	b_search = b_search & "[comments] LIKE '%" & pSearch & "%' OR "
	End If
End If
If Right(b_search, 4) = " OR " Then b_search = Left(b_search, Len(b_search)-4)
If Right(b_search, 5) = " AND " Then b_search = Left(b_search, Len(b_search)-5)
%>
<%

' Build search criteria
If a_search <> "" Then
	searchwhere = a_search ' Advanced search
ElseIf b_search <> "" Then
	searchwhere = b_search ' Basic search
End If

' Save search criteria
If searchwhere <> "" Then
	Session("otherteams_searchwhere") = searchwhere

	' Reset start record counter (new search)
	startRec = 1
	Session("otherteams_REC") = startRec
Else
	searchwhere = Session("otherteams_searchwhere")
End If
%>
<%

' Get clear search cmd
If Request.QueryString("cmd").Count > 0 Then
	cmd = Request.QueryString("cmd")
	If UCase(cmd) = "RESET" Then

		' Reset search criteria
		searchwhere = ""
		Session("otherteams_searchwhere") = searchwhere
  ElseIf UCase(cmd) = "RESETALL" Then

		' Reset search criteria
		searchwhere = ""
		Session("otherteams_searchwhere") = searchwhere
	End If

	' Reset start record counter (reset command)
	startRec = 1
	Session("otherteams_REC") = startRec
End If

' Build dbwhere
If masterdetailwhere <> "" Then
	dbwhere = dbwhere & "(" & masterdetailwhere & ") AND "
End If
If searchwhere <> "" Then
	dbwhere = dbwhere & "(" & searchwhere & ") AND "
End If
If Len(dbwhere) > 5 Then
	dbwhere = Mid(dbwhere, 1, Len(dbwhere)-5) ' Trim rightmost AND
End If
%>
<%

' Load Default Order
DefaultOrder = ""
DefaultOrderType = ""

' No Default Filter
DefaultFilter = ""

' Check for an Order parameter
OrderBy = ""
If Request.QueryString("order").Count > 0 Then
	OrderBy = Request.QueryString("order")

	' Check if an ASC/DESC toggle is required
	If Session("otherteams_OB") = OrderBy Then
		If Session("otherteams_OT") = "ASC" Then
			Session("otherteams_OT") = "DESC"
		Else
			Session("otherteams_OT") = "ASC"
		End if
	Else
		Session("otherteams_OT") = "ASC"
	End If
	Session("otherteams_OB") = OrderBy
	Session("otherteams_REC") = 1
Else
	OrderBy = Session("otherteams_OB")
	If OrderBy = "" Then
		OrderBy = DefaultOrder
		Session("otherteams_OB") = OrderBy
		Session("otherteams_OT") = DefaultOrderType
	End If
End If

' Open connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str

' Build SQL
strsql = "SELECT * FROM [otherteams]"
If DefaultFilter <> "" Then
	whereClause = whereClause & "(" & DefaultFilter & ") AND "
End If
If dbwhere <> "" Then
	whereClause = whereClause & "(" & dbwhere & ") AND "
End If
If Right(whereClause, 5)=" AND " Then whereClause = Left(whereClause, Len(whereClause)-5)
If whereClause <> "" Then
	strsql = strsql & " WHERE " & whereClause
End If
If OrderBy <> "" Then 
	strsql = strsql & " ORDER BY [" & OrderBy & "] " & Session("otherteams_OT")
End If	

'Response.Write strsql
Set rs = Server.CreateObject("ADODB.Recordset")
rs.cursorlocation = 3
rs.Open strsql, conn, 1, 2
totalRecs = rs.RecordCount

' Check for a START parameter
If Request.QueryString("start").Count > 0 Then
	startRec = Request.QueryString("start")
	Session("otherteams_REC") = startRec
ElseIf Request.QueryString("pageno").Count > 0 Then
	pageno = Request.QueryString("pageno")
	If IsNumeric(pageno) Then
		startRec = (pageno-1)*displayRecs+1
		If startRec <= 0 Then
			startRec = 1
		ElseIf startRec >= ((totalRecs-1)\displayRecs)*displayRecs+1 Then
			startRec = ((totalRecs-1)\displayRecs)*displayRecs+1
		End If
		Session("otherteams_REC") = startRec
	Else
		startRec = Session("otherteams_REC")
		If Not IsNumeric(startRec) Or startRec = "" Then			
			startRec = 1 ' Reset start record counter
			Session("otherteams_REC") = startRec
		End If
	End If
Else
	startRec = Session("otherteams_REC")
	If Not IsNumeric(startRec) Or startRec = "" Then		
		startRec = 1 'Reset start record counter
		Session("otherteams_REC") = startRec
	End If
End If
%>
<!--#include file="header.asp"-->
<p><span class="aspmaker">TABLE: otherteams</span></p>
<form action="otherteamslist.asp">
<table border="0" cellspacing="0" cellpadding="4">
	<tr>
		<td><span class="aspmaker">Quick Search (*)</span></td>		
		<td><span class="aspmaker">
			<input type="text" name="psearch" size="20">
			<input type="Submit" name="Submit" value="GO">
		&nbsp;&nbsp;<a href="otherteamslist.asp?cmd=reset">Show all</a>
		</span></td>
	</tr>
	<tr><td>&nbsp;</td><td><span class="aspmaker"><input type="radio" name="psearchtype" value="" checked>Exact phrase&nbsp;&nbsp;<input type="radio" name="psearchtype" value="AND">All words&nbsp;&nbsp;<input type="radio" name="psearchtype" value="OR">Any word</span></td></tr>	
</table>
</form>
<form method="post">
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#CCCCCC">
	<tr bgcolor="#0099CC">
<td>&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
		<td><span class="aspmaker" style="color: #FFFFFF;">
<a href="otherteamslist.asp?order=<%= Server.URLEncode("uniqref") %>" style="color: #FFFFFF;">uniqref&nbsp;<% If OrderBy = "uniqref" Then %><span class="ewTableOrderIndicator"><% If Session("otherteams_OT") = "ASC" Then %>5<% ElseIf Session("otherteams_OT") = "DESC" Then %>6<% End If %></span><% End If %></a>
		</span></td>
		<td><span class="aspmaker" style="color: #FFFFFF;">
<a href="otherteamslist.asp?order=<%= Server.URLEncode("clubname") %>" style="color: #FFFFFF;">clubname&nbsp;(*)<% If OrderBy = "clubname" Then %><span class="ewTableOrderIndicator"><% If Session("otherteams_OT") = "ASC" Then %>5<% ElseIf Session("otherteams_OT") = "DESC" Then %>6<% End If %></span><% End If %></a>
		</span></td>
		<td><span class="aspmaker" style="color: #FFFFFF;">
<a href="otherteamslist.asp?order=<%= Server.URLEncode("teamname") %>" style="color: #FFFFFF;">teamname&nbsp;(*)<% If OrderBy = "teamname" Then %><span class="ewTableOrderIndicator"><% If Session("otherteams_OT") = "ASC" Then %>5<% ElseIf Session("otherteams_OT") = "DESC" Then %>6<% End If %></span><% End If %></a>
		</span></td>
</tr>
<%

' Avoid starting record > total records
If CLng(startRec) > CLng(totalRecs) Then
	startRec = totalRecs
End If

' Set the last record to display
stopRec = startRec + displayRecs - 1

' Move to first record directly for performance reason
recCount = startRec - 1
If Not rs.Eof Then
	rs.MoveFirst
	rs.Move startRec - 1
End If
recActual = 0
Do While (Not rs.Eof) And (recCount < stopRec)
	recCount = recCount + 1
	If CLng(recCount) >= CLng(startRec) Then 
		recActual = recActual + 1 %>
<%

	' Set row color
	bgcolor = "#FFFFFF"
%>
<%	

	' Display alternate color for rows
	If recCount Mod 2 <> 0 Then
		bgcolor = "#F5F5F5"
	End If
%>
<%

	' Load Key for record
	key = rs("uniqref")
	x_uniqref = rs("uniqref")
	x_clubname = rs("clubname")
	x_teamname = rs("teamname")
	x_comments = rs("comments")
%>
	<tr bgcolor="<%= bgcolor %>">
<td><span class="aspmaker"><a href="<% key = rs("uniqref") : If Not IsNull(key) Then Response.Write "otherteamsview.asp?key=" & Server.URLEncode(key) Else Response.Write "javascript:alert('Invalid Record! Key is null');" End If %>">View</a></span></td>
<td><span class="aspmaker"><a href="<% key = rs("uniqref") : If Not IsNull(key) Then Response.Write "otherteamsedit.asp?key=" & Server.URLEncode(key) Else Response.Write "javascript:alert('Invalid Record! Key is null');" End If %>">Edit</a></span></td>
<td><span class="aspmaker"><a href="<% key = rs("uniqref") : If Not IsNull(key) Then Response.Write "otherteamsadd.asp?key=" & Server.URLEncode(key) Else Response.Write "javascript:alert('Invalid Record! Key is null');" End If %>">Copy</a></span></td>
<td><span class="aspmaker"><a href="<% key = rs("uniqref") : If Not IsNull(key) Then Response.Write "otherteamsdelete.asp?key=" & Server.URLEncode(key) Else Response.Write "javascript:alert('Invalid Record! Key is null');" End If %>">Delete</a></span></td>
		<td><span class="aspmaker"><% Response.Write x_uniqref %></span>&nbsp;</td>
		<td><span class="aspmaker"><% Response.Write x_clubname %></span>&nbsp;</td>
		<td><span class="aspmaker"><% Response.Write x_teamname %></span>&nbsp;</td>
	</tr>
<%
	End If
	rs.MoveNext
Loop
%>
</table>
</form>
<%

' Close recordset and connection
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%>
<table border="0" cellspacing="0" cellpadding="10"><tr><td>
<%
If totalRecs > 0 Then
	rsEof = (totalRecs < (startRec + displayRecs))
	PrevStart = startRec - displayRecs
	If PrevStart < 1 Then PrevStart = 1
	NextStart = startRec + displayRecs
	If NextStart > totalRecs Then NextStart = startRec
	LastStart = ((totalRecs-1)\displayRecs)*displayRecs+1
	%>
<form>	
	<table border="0" cellspacing="0" cellpadding="0"><tr><td><span class="aspmaker">Page</span>&nbsp;</td>
<!--first page button-->
	<% If CLng(startRec)=1 Then %>
	<td><img src="images/firstdisab.gif" alt="First" width="20" height="15" border="0"></td>
	<% Else %>
	<td><a href="otherteamslist.asp?start=1"><img src="images/first.gif" alt="First" width="20" height="15" border="0"></a></td>
	<% End If %>
<!--previous page button-->
	<% If CLng(PrevStart) = CLng(startRec) Then %>
	<td><img src="images/prevdisab.gif" alt="Previous" width="20" height="15" border="0"></td>
	<% Else %>
	<td><a href="otherteamslist.asp?start=<%=PrevStart%>"><img src="images/prev.gif" alt="Previous" width="20" height="15" border="0"></a></td>
	<% End If %>
<!--current page number-->
	<td><input type="text" name="pageno" value="<%=(startRec-1)\displayRecs+1%>" size="4"></td>
<!--next page button-->
	<% If CLng(NextStart) = CLng(startRec) Then %>
	<td><img src="images/nextdisab.gif" alt="Next" width="20" height="15" border="0"></td>
	<% Else %>
	<td><a href="otherteamslist.asp?start=<%=NextStart%>"><img src="images/next.gif" alt="Next" width="20" height="15" border="0"></a></td>
	<% End If %>
<!--last page button-->
	<% If CLng(LastStart) = CLng(startRec) Then %>
	<td><img src="images/lastdisab.gif" alt="Last" width="20" height="15" border="0"></td>
	<% Else %>
	<td><a href="otherteamslist.asp?start=<%=LastStart%>"><img src="images/last.gif" alt="Last" width="20" height="15" border="0"></a></td>
	<% End If %>
	<td><a href="otherteamsadd.asp"><img src="images/addnew.gif" alt="Add new" width="20" height="15" border="0"></a></td>
	<td><span class="aspmaker">&nbsp;of <%=(totalRecs-1)\displayRecs+1%></span></td>
	</tr></table>	
</form>	
	<% If CLng(startRec) > CLng(totalRecs) Then startRec = totalRecs
	stopRec = startRec + displayRecs - 1
	recCount = totalRecs - 1
	If rsEOF Then recCount = totalRecs
	If stopRec > recCount Then stopRec = recCount %>
	<span class="aspmaker">Records <%= startRec %> to <%= stopRec %> of <%= totalRecs %></span>
<% Else %>
	<span class="aspmaker">No records found</span>
<p>
<a href="otherteamsadd.asp"><img src="images/addnew.gif" alt="Add new" width="20" height="15" border="0"></a>
</p>
<% End If %>
</td></tr></table>
<!--#include file="footer.asp"-->
