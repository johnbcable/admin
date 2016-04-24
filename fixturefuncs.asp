<script language="JScript" CODEPAGE="65001" runat=Server>
//
//  getFixture(fixtureid)                     		returns fixture object
//  setFixture(fixtureobject)				returns boolean
//
// ================================================================
function getFixture(fixtureid)
{
	// Establish local variables
	var thefixture = new Object();
	var sFixture = new String(fixtureid);
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	var mDateObj, dummy1;
	//
	thefixture.fixtureid = new String(fixtureid).toString();
	thefixture.fixturedate = new String("").toString();
	thefixture.teamname = new String("").toString();
	thefixture.venue = new String("").toString();
	thefixture.opponents = new String("").toString();
	thefixture.pair1 = new String("").toString();
	thefixture.pair2 = new String("").toString();
	thefixture.ourscore = new String("").toString();
	thefixture.theirscore = new String("").toString();
	thefixture.note = new String("").toString();
	thefixture.season = new String("").toString();

	//
	dbconnect=Application("hamptonsportsdb");
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	// Need to query particular email database entry from fileref
	SQL1 = new String("SELECT * from tennisfixtures where fixtureid="+sFixture).toString();
	RS = Conn.Execute(SQL1);
	if (! RS.EOF)
	{
		thefixture.teamname = new String(RS("teamname")).toString();
		thefixture.venue = new String(RS("homeoraway")).toString();
		mDateObj=new Date(RS("fixturedate"));
		dummy1 = mDateObj.valueOf();
		if (dummy1 == 0) // no date in database
		{
			thefixture.fixturedate = new String("").toString();
		}
		else
		{
			thefixture.fixturedate = new String(ddmmyyyy(mDateObj)).toString();
		}
		thefixture.fixtureyear
		thefixture.fixturedate = new String(RS("fixturedate")).toString();
		thefixture.opponents = new String(RS("opponents")).toString();
		thefixture.pair1 = new String(RS("pair1")).toString();
		thefixture.pair2 = new String(RS("pair2")).toString();
		thefixture.note = new String(RS("fixturenote")).toString();
		thefixture.ourscore = new String(RS("hamptonresult")).toString();
		thefixture.theirscore = new String(RS("opponentresult")).toString();
		thefixture.season = new String(RS("fixtureyear")).toString();
	}
	return(thefixture);
}
// ================================================================
function setFixture(fixtureobj)
{
	// Establish local variables
	var sFixture = new String(fixtureobj.fixtureid);
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	var mDateObj, dummy1;
	//
	dbconnect=Application("hamptonsportsdb");
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	SQLstart = new String("UPDATE tennisfixtures ")
	SQLend = new String(" WHERE fixtureid="+fixtureobj.fixtureid).toString();
	SQLmiddle = new String("SET ").toString();
	SQLmiddle += "teamname='"+fixtureobj.teamname+"',";
	SQLmiddle += "homeoraway = '"+fixtureobj.venue+"',";
	SQLmiddle += "fixturedate = '"+fixtureobj.fixturedate+"',";
	SQLmiddle += "opponents = '"+fixtureobj.opponents+"',";
	SQLmiddle += "pair1 = '"+fixtureobj.pair1+"',";
	SQLmiddle += "pair2 = '"+fixtureobj.pair2+"',";
	SQLmiddle += "fixturenote = '"+fixtureobj.note+"',";
	SQLmiddle += "hamptonresult = "+fixtureobj.ourscore+",";
	SQLmiddle += "opponentresult = "+fixtureobj.theirscore+",";
	SQLmiddle += "fixtureyear = '"+fixtureobj.season+"'";
	SQL1 = new String(SQLstart+SQLmiddle+SQLend).toString();;
	RS = Conn.Execute(SQL1);
	return(true);
}
// ================================================================</script>