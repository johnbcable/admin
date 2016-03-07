<script language=JScript runat=Server>
//
//  getTeamCaptains(category)                    	returns string
//  getTeam(teamid)                         		returns team object
//  getLeagueTeam(division,teamnumber)     			returns leagueteam object
//  getTeamByName(teamname)                    		returns team object
//  printTeam(teamObj)								returns string
//
function getTeamCaptains(teamcategory)
{
	// Establish local variables
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	var debugging = false;
	var editline, emailsubject;
	var captainlist = new String("").toString();
	var captainknt = 0;
	var thecategory = new String(teamcategory).toString();
	thecategory = thecategory.toUpperCase();
	if (thecategory == "" || thecategory == "NULL" || thecategory == "UNDEFINED")
		thecategory = new String("ADULT").toString();
	// Set up database connections
	dbconnect=Application("hamptonsportsdb"); 
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	SQL1 = new String("SELECT * from tennisteams where teamcategory = '"+thecategory+"' order by teamname").toString();
	captainlist += '<ul>';
	RS = Conn.Execute(SQL1);
	captainknt = 0;
	while (! RS.EOF)
	{
		dummy = new String(RS("teamname")+"                                                                 ").toString();
		captainlist += "<li>"+dummy.substr(0,30)+" - "+RS("teamcaptain")+" ("+RS("teamnote")+")</li>";
		captainknt++;
		RS.MoveNext();
	}
	captainlist += "</ul>";
	RS.Close();
	return (captainlist);
}
// ================================================================
function getTeam(teamid)
{
	// Establish local variables
	var theteam = new Object();
	var sTeam = new String(teamid);
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	var mDateObj, dummy1, dummy2;
	//
	theteam.teamid = teamid;
	theteam.teamname = new String("").toString();
	theteam.teamcategory = new String("").toString();
	theteam.teamcaptain = new String("").toString();
	theteam.teamnote = new String("").toString();
	theteam.fixtureline = new String("").toString();
	theteam.uniqueref = new String("").toString();
	theteam.email = new String("").toString();
	theteam.blogline = new String("").toString();
	theteam.teamnumber = 0;
	theteam.division = new String("").toString();
	//
	dbconnect=Application("hamptonsportsdb");
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	// Need to query particular email database entry from fileref
	SQL1 = new String("SELECT * from tennisteams where teamid="+sTeam).toString();
	RS = Conn.Execute(SQL1);
	if (! RS.EOF)
	{
		theteam.teamname = new String(RS("teamname")).toString();
		theteam.teamcategory = new String(RS("teamcategory")).toString();
		theteam.teamcaptain = new String(RS("teamcaptain")).toString();
		theteam.teamnote = new String(RS("teamnote")).toString();
		theteam.fixtureline = new String(RS("fixtureline")).toString();
		theteam.uniqueref = new String(RS("uniqueref")).toString();
		theteam.email = new String(RS("email")).toString();
		theteam.blogline = new String(RS("blogline")).toString();
		dummy2 = new String(RS("teamnumber")).toString();
		theteam.division = new String(RS("division")).toString();
		if (theteam.teamcaptain == "undefined" || theteam.teamcaptain == "null" )
			theteam.teamcaptain = new String("");
		if (theteam.teamnote == "undefined" || theteam.teamnote == "null" )
			theteam.teamnote = new String("");
		if (theteam.fixtureline == "undefined" || theteam.fixtureline == "null" )
			theteam.fixtureline = new String("");
		if (theteam.email == "undefined" || theteam.email == "null" )
			theteam.email = new String("");
		if (theteam.blogline == "undefined" || theteam.blogline == "null" )
			theteam.blogline = new String("");
		if (theteam.uniqueref == "undefined" || theteam.uniqueref == "null" )
			theteam.uniqueref = new String("");
		if (dummy2 == "undefined" || dummy2 == "null" )
		{
			theteam.teamnumber = 0;
		}
		else
		{
			theteam.teamnumber = new Number(dummy2);
			theteam.teamnumber = theteam.teamnumber.valueOf();
		}
		if (theteam.division == "undefined" || theteam.division == "null" )
			theteam.division = new String("");
	}
	return(theteam);
}
// ================================================================
function getLeagueTeam(division,teamnumber)
{
	// Establish local variables
	var theleagueteam = new Object();
	var sDivision = new String(division);
	var sTeamNumber = new String(teamnumber);
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	var mDateObj, dummy1, dummy2;
	//
	theleagueteam.division = new String(sDivision).toString();
	theleagueteam.teamnumber = new String(sTeamNumber).toString();
	theleagueteam.homenightoffset = new String("").toString();
	theleagueteam.teamname = new String("NOT FOUND").toString();
	//
	dbconnect=Application("hamptonsportsdb");
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	// Need to query particular email database entry from fileref
	SQL1 = new String("SELECT * from LeagueDivisions where division='"+sDivision+"' and teamnumber="+sTeamNumber).toString();
	RS = Conn.Execute(SQL1);
	if (! RS.EOF)
	{
		theleagueteam.division = new String(RS("division")).toString();
		theleagueteam.teamnumber = new String(RS("teamnumber")).toString();
		theleagueteam.homenightoffset = new String(RS("homenightoffset")).toString();
		theleagueteam.teamname = new String(RS("teamname")).toString();
		if (theleagueteam.division == "undefined" || theleagueteam.division == "null" )
			theleagueteam.division = new String(sDivision);
		if (theleagueteam.teamnumber == "undefined" || theleagueteam.teamnumber == "null" )
			theleagueteam.teamnumber = new String(sTeamNumber);
	}
	return(theleagueteam);
}
// ================================================================
function getTeamByName(teamname)
{
	// Establish local variables
	var theteam = new Object();
	var sTeam = new String(teamname).toString();
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	var mDateObj, dummy1, dummy2;
	//
	theteam.teamid = 0;
	theteam.teamname = new String("").toString();
	theteam.teamcategory = new String("").toString();
	theteam.teamcaptain = new String("").toString();
	theteam.teamnote = new String("").toString();
	theteam.fixtureline = new String("").toString();
	theteam.uniqueref = new String("").toString();
	theteam.email = new String("").toString();
	theteam.blogline = new String("").toString();
	theteam.teamnumber = 0;
	theteam.division = new String("").toString();
	//
	dbconnect=Application("hamptonsportsdb");
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	// Need to query particular email database entry from fileref
	SQL1 = new String("SELECT * from tennisteams where teamname like '"+sTeam+"'").toString();
	RS = Conn.Execute(SQL1);
	if (! RS.EOF)
	{
		theteam.teamid = RS("teamid");
		theteam.teamname = new String(RS("teamname")).toString();
		theteam.teamcategory = new String(RS("teamcategory")).toString();
		theteam.teamcaptain = new String(RS("teamcaptain")).toString();
		theteam.teamnote = new String(RS("teamnote")).toString();
		theteam.fixtureline = new String(RS("fixtureline")).toString();
		theteam.uniqueref = new String(RS("uniqueref")).toString();
		theteam.email = new String(RS("email")).toString();
		dummy2 = new String(RS("teamnumber")).toString();
		theteam.division = new String(RS("division")).toString();
		if (theteam.teamcaptain == "undefined" || theteam.teamcaptain == "null" )
			theteam.teamcaptain = new String("");
		if (theteam.teamnote == "undefined" || theteam.teamnote == "null" )
			theteam.teamnote = new String("");
		if (theteam.fixtureline == "undefined" || theteam.fixtureline == "null" )
			theteam.fixtureline = new String("");
		if (theteam.uniqueref == "undefined" || theteam.uniqueref == "null" )
			theteam.uniqueref = new String("");
		if (theteam.email == "undefined" || theteam.email == "null" )
			theteam.email = new String("");
		if (theteam.blogline == "undefined" || theteam.blogline == "null" )
			theteam.blogline = new String("");
		if (dummy2 == "undefined" || dummy2 == "null" )
		{
			theteam.teamnumber = 0;
		}
		else
		{
			theteam.teamnumber = new Number(dummy2);
			theteam.teamnumber = theteam.teamnumber.valueOf();
		}
		if (theteam.division == "undefined" || theteam.division == "null" )
			theteam.division = new String("");
	}
	return(theteam);
}
// ================================================================
function printTeam(teamobj)
{
	// Establish local variables
	var sReport = new String("").toString();
		sReport += teamobj.teamid +"<br />";
		sReport += teamobj.teamname +"<br />";
   		sReport += teamobj.teamcategory +"<br />";
		sReport += teamobj.teamcaptain +"<br />";
		sReport += teamobj.teamnote +"<br />";
		sReport += teamobj.fixtureline +"<br />";
		sReport += teamobj.uniqueref +"<br />";
		sReport += teamobj.email +"<br />";
		sReport += teamobj.blogline +"<br />";
		sReport += teamobj.teamnumber +"<br />";
		sReport += teamobj.division +"<br />";
	return(sReport);
}
// ================================================================
</script>
