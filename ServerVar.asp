<%@ LANGUAGE="VBSCRIPT" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2//EN">
<HTML>
<HEAD>
<TITLE>Server Variables</TITLE>
</HEAD><BODY>
<TABLE border=1>
<TR><TD><B>Server Variable</B></TD><TD><B>Value</B></TD></TR>
<% For Each ServerVariable In Request.ServerVariables %> 
<TR><TD> <%= ServerVariable %> </TD><TD>  <%= Request.ServerVariables(ServerVariable) %> </TD></TR>

<% Next %> 
</TABLE>

</BODY></HTML>
<%
%>
