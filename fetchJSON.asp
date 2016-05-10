<%@language="VBScript"%>
<!--#include file="JSON_2.0.4.asp" -->
<!--#include file="JSON_UTIL_0.1.1.asp" -->
<%
' This page sends back JSON data to a requesting agent 
' From and identified DB query
'Dimension variables
Dim adoCon         'Holds the Database Connection Object
Dim rsDB		   'Holds the recordset for the records in the database
Dim strSQL         'Holds the SQL query to query the database	
Dim dataSource	   'Holds the name of the data source from the Application object
Dim dataResults    'Holds results from the JSON query
Dim querylist(40)  'Array of queries
Dim queryref       'Reference to query in querylist (default to 1)
Dim p1, p2, p3     'Parameters (text)
Dim paramknt       'Count of the number of parameters
Dim debugging      'Are we running in debug mode
Dim debug          'Placeholder for if we are to display debug info
Dim origSQL

p1 = ""
p2 = ""
p3 = ""
paramknt = 0

debugging = False  'Default production provision

'Now retrieve query details from submitting form or querystring
'and initialise the strSQL variable with an SQL statement to query the database
queryref = Request.Form("id")
If queryref = "" Then
	queryref = Request.QueryString("id")
	If queryref = "" Then
		queryref = -1
	End If
End If

' Optionally consider dealing with fill-in fields for the identified query
If queryref > -1 Then	

	p1 = Request.Form("p1")
	If p1 = "" Then
		p1 = Request.QueryString("p1")
		If p1 <> "" Then
			paramknt = paramknt + 1
		End If
	End If

	p2 = Request.Form("p2")
	If p2 = "" Then
		p2 = Request.QueryString("p2")
		If p2 <> "" Then 
			paramknt = paramknt + 2
		End If
	End If

	p3 = Request.Form("p3")
	If p3 = "" Then
		p3 = Request.QueryString("p3")
		If p3 <> "" Then
			paramknt = paramknt + 4
		End If
	End If

Else

 	'No query Id so use the supplied query
	thequery = Request.Form("thequery")
	If thequery = "" Then
		thequery = Request.QueryString("thequery")
		If thequery = "" Then
			thequery = "SELECT ""There are no results"" FROM coaches"
		End If
	End If

End If

debug = Request.Form("debug")
If debug = "" Then
	debug = Request.QueryString("debug")
	If debug = "Y" Then
		debugging = True
	End If
Else
	If debug = "Y" Then
		debugging = True
	End If
End If

'Retrieve the datSource name from the Application object
dataSource = Application("hamptonsportsdb")
'Create an ADO connection object
Set adoCon = Server.CreateObject("ADODB.Connection")	
'Set an active connection to the Connection object using DSN connection
adoCon.Open dataSource
'Create an ADO recordset object
Set rsDB = Server.CreateObject("ADODB.Recordset")	

If queryref > -1 Then

	'Initialise querylist with queries
	querylist(0) = "SELECT * FROM coaches ORDER BY [displayorder] ASC;"
	querylist(1) = "SELECT * FROM coaches WHERE [uniqueref] = {{p1}};"
	querylist(2) = "SELECT * FROM futureevents ORDER BY [sortabledate] ASC;"
	querylist(3) = "SELECT Format([eventdate],'yyyy-mm-dd') as eventdate, eventtime, eventtype, eventnote, eventreport, Format([enddate],'yyyy-mm-dd') as enddate, endtime FROM futureevents WHERE [eventid] = {{p1}};"
	querylist(4) = "SELECT * FROM futuretournaments ORDER BY [tourstart] ASC;"
	querylist(5) = "SELECT * FROM alltournaments WHERE [tournamentid] = {{p1}};"
	querylist(6) = "SELECT * FROM allmembers WHERE [onlinebookingid] < 5000 ORDER BY [surname], [forename1];"
	querylist(7) = "SELECT * FROM allmembers WHERE [homephone] NOT LIKE '' AND [adultphone] = [homephone] ORDER BY [surname], [forename1];"
	querylist(8) = "SELECT * FROM tournaments ORDER BY [tourstart] ASC;"
	querylist(9) = "SELECT * FROM clubcommunications ORDER by [datesent] DESC;"
	querylist(10) = "SELECT [membergrade], [forename1], [surname], [britishtennisno] FROM {{p1}} WHERE [membergrade] = '{{p2}}' ORDER BY [surname], [forename1];"
	querylist(11) = "SELECT [membergrade], COUNT(*) as categorytotal FROM notinbritishtennis GROUP BY [membergrade] ORDER BY [membergrade];"
	querylist(12) = "SELECT * FROM alltournaments;"
	querylist(13) = "SELECT * FROM allmembers WHERE email LIKE '{{p1}}';"
	querylist(14) = "SELECT * FROM {{p1}}"
	querylist(15) = "SELECT * from allmembers where mobilephone = '{{p1}}'"
	querylist(16) = "SELECT title FROM winners WHERE year = {{p1}}"
	querylist(17) = "SELECT Format(fixturedate,'YYYY-mm-dd') AS fixturedate, homeoraway, opponents, hamptonresult, opponentresult, fixtureyear, teamname, fixturenote, fixtureid, matchreport, pair1, pair2 FROM fixturesetup where [teamname] = '{{p1}}' ORDER BY fixtureid ASC"

	querylist(19) = "SELECT [Name], [Type] FROM MSysObjects where [Type] = {{p1}} AND [Flags] = 0;"



	strSQL = querylist(queryref)
	origSQL = strSQL

	'Now, replace any fill-in fields with their params

	If (paramknt > 0)  Then
		strSQL = replace(strSQL, "{{p1}}", p1)
	End If

	If (paramknt > 1)  Then
		strSQL = replace(strSQL, "{{p2}}", p2)
	End If

	If (paramknt > 2)  Then
		strSQL = replace(strSQL, "{{p3}}", p3)
	End If

	'Now for any predecided substitutions
	
Else

	strSQL = thequery
	origSQL = strSQL

End If

If debugging Then
	Response.Write("p1 = [" & p1 & "]<br />")
	Response.Write("p2 = [" & p2 & "]<br />")
	Response.Write("p3 = [" & p3 & "]<br />")
	Response.Write("origSQL = [" & origSQL & "]<br />")
	Response.Write("strSQL = [" & strSQL & "]<br />")
	Response.End
End If

dataResults = QueryToJSON(adoCon, strSQL).Flush

Response.ContentType = "application/json"
Response.Write(dataResults)
Response.End
%>
