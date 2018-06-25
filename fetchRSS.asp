<%@language="JScript"%>
<%

var rssResults = new String("").toString();
var rssStart = new String("").toString();
var rssEnd = new String("").toString();
var rssItems = new String("").toString();


/*
-- Test data
*/

rssStart = "<?xml version='1.0' encoding='utf-8'?><rss version='2.0'><channel><title>Test RSS Channel</title><description>Test RSS Channel</description><link>https://hamptontennis.org.uk</link>";

rssEnd = "</channel></rss>";

rssItems = "<item><title>Item Title</title><description>Updated on: 5/20/2012</description><link>https://hamptontennis.org.uk/fetchRSS.asp.asp</link>";

rssResults = rssStart+rssItems+rssEnd;

Response.ContentType = "application/rss+xml";
Response.Write(rssResults);
Response.End();
%>
