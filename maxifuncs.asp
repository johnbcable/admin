<script language="JScript" CODEPAGE="65001" runat=Server>
//
//  getMaxiCaptainsAsString(maxiteamname)         	returns string
//
// ================================================================
//
function getMaxiCaptainsAsString(maxiteamname)
{
	// Establish local variables
	var thecaptains = new String("").toString();
	var sTeam = new String(maxiteamname);
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	var mDateObj, dummy1;
	var first = true;
	//
	dbconnect=Application("hamptonsportsdb");
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	// 
	// Put in "Maxi-Tennis:"+  prefix if maxiteamname doesnt contain it?
	sdummy = 
	SQL1 = new String("SELECT * from maxicaptains where maxiteam = '" + sTeam + "'");
	RS = Conn.Execute(SQL1);
	// Retrieve captains name details into string
	while (! RS.EOF)
	{
		if (!first)
			thecaptains += " and ";
		thecaptains += new String(RS("forename1")).toString();
		thecaptains += new String(" "+RS("surname")).toString();
		first = false;
		RS.MoveNext();
	}
	
	return(thecaptains);
}
// ================================================================
</script>