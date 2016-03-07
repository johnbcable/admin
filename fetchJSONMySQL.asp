


<%
Dim oConn, oRs
Dim qry, connectstr
Dim db_name, db_username, db_userpassword
Dim db_server

db_server = "mysql.secureserver.net"
db_name = "your_dbusername"
db_username = "your_dbusername"
db_userpassword = "your_dbpassword"
fieldname = "your_field"
tablename = "your_table"

connectstr = "Driver={MySQL ODBC 3.51 Driver};SERVER=" & db_server & ";DATABASE=" & db_name & ";UID=" & db_username & ";PWD=" & db_userpassword

Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open connectstr

qry = "SELECT * FROM " & tablename

Set oRS = oConn.Execute(qry)

if not oRS.EOF then
while not oRS.EOF
response.write ucase(fieldname) & ": " & oRs.Fields(fieldname) & "<br>"
oRS.movenext
wend
oRS.close
end if

Set oRs = nothing
Set oConn = nothing

%>