<script language="JScript" CODEPAGE="65001" runat=Server>
//
//  getWinner(winnerid)                 returns winner object
//  setWinner(winnerobject)				returns boolean
//
// ================================================================
function getWinner(winnerid)
{
	// Establish local variables
	var thewinner = new Object();
	var sWinner = new String(winnerid);
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	var mDateObj, dummy1;
	//
	thewinner.winnerid = new String(sWinner).toString();
	thewinner.year = new String("").toString();
	thewinner.title = new String("").toString();
	thewinner.winners = new String("").toString();
	thewinner.section = new String("").toString();
	thewinner.member1 = new String("").toString();
	thewinner.member2 = new String("").toString();
	thewinner.runnerup = new String("").toString();
	//
	dbconnect=Application("hamptonsportsdb");
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	// Need to query particular email database entry from fileref
	SQL1 = new String("SELECT * from winners where winnerid="+sWinner).toString();
	RS = Conn.Execute(SQL1);
	if (! RS.EOF)
	{
		thewinner.year = new String(RS("year")).toString();
		thewinner.title = new String(RS("title")).toString();
		thewinner.winners = new String(RS("winners")).toString();
		thewinner.section = new String(RS("section")).toString();
		thewinner.member1 = new String(RS("member1")).toString();
		thewinner.member2 = new String(RS("member2")).toString();
		thewinner.runnerup = new String(RS("runnerup")).toString();
		if (thewinner.year == "" || thewinner.year == "null" || thewinner.year == "undefined")
			thewinner.year = new String("").toString();
		if (thewinner.title == "" || thewinner.title == "null" || thewinner.title == "undefined")
			thewinner.title = new String("").toString();
		if (thewinner.winners == "" || thewinner.winners == "null" || thewinner.winners == "undefined")
			thewinner.winners = new String("").toString();
		if (thewinner.section == "" || thewinner.section == "null" || thewinner.section == "undefined")
			thewinner.section = new String("ADULT TENNIS").toString();
		if (thewinner.member1 == "" || thewinner.member1 == "null" || thewinner.member1 == "undefined")
			thewinner.member1 = new String("").toString();
		if (thewinner.member2 == "" || thewinner.member2 == "null" || thewinner.member2 == "undefined")
			thewinner.member2 = new String("").toString();
		if (thewinner.runnerup == "" || thewinner.runnerup == "null" || thewinner.runnerup == "undefined")
			thewinner.runnerup = new String("").toString();
	}
	return(thewinner);
}
// ================================================================
function setWinner(winnerobj)
{
	// Establish local variables
	var sWinner = new String(winnerobj.winnerid);
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	var mDateObj, dummy1;
	//
	dbconnect=Application("hamptonsportsdb");
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	SQLstart = new String("UPDATE winners ")
	SQLend = new String(" WHERE winnerid="+winnerobj.winnerid).toString();
	SQLmiddle = new String("SET ").toString();
	SQLmiddle += "[year]="+winnerobj.year+",";
	SQLmiddle += "[title] = '"+winnerobj.title+"',";
	SQLmiddle += "[winners] = '"+winnerobj.winners+"',";
	SQLmiddle += "[section] = '"+winnerobj.section+"',";
	SQLmiddle += "[member1] = "+winnerobj.member1+",";
	SQLmiddle += "[member2] = "+winnerobj.member2+", ";
	SQLmiddle += "[runnerup] = '"+winnerobj.runnerup+"'";
	SQL1 = new String(SQLstart+SQLmiddle+SQLend).toString();;
	try {
		RS = Conn.Execute(SQL1);
	}
	catch(e) {
		resultObj.result = false;
		resultObj.errno = (e.number & 0xFFFF);
		resultObj.description += e.description;
		resultObj.sql = new String(SQLstart+SQLmiddle+SQLend).toString();
	}
	return(resultObj);
}
// ================================================================</script>