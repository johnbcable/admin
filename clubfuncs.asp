<script language="JScript" CODEPAGE="65001" runat=Server>
//
//  getClubEmailAddressList()                   returns string
//  getClub(clubid)                                      returns club object
//
// ================================================================
function getClubEmailAddressList()
{
	// Establish local variables
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	var debugging = false;
	var editline, emailsubject;
	var cluburllist = new String("").toString();
	var cluburlknt = 0;
	var dummy, m_cluburl, m_totaltennis;
	// Set up database connections
	dbconnect=Application("hamptonsportsdb");
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	SQL1 = new String("SELECT * from otherclubs where totaltennis is not null OR cluburl is not null order by clubname").toString();
	cluburllist += '<ul>';
	RS = Conn.Execute(SQL1);
	cluburlknt = 0;
	while (! RS.EOF)
	{
		dummy = new String(RS("clubname")+"                                                                 ").toString();
		m_cluburl = new String(RS("cluburl")).toString();
		if (m_cluburl == "" || m_cluburl == "null" || m_cluburl == "undefined")
			m_cluburl = new String(RS("totaltennis")).toString();
		cluburllist += '<li>&nbsp;&nbsp;&nbsp;&nbsp;<a href="'+m_cluburl+'">'+dummy.substr(0,30)+'</a></li>';
		cluburlknt++;
		RS.MoveNext();
	}
	cluburllist += "</ul>";
	RS.Close();
	return (cluburllist);
}
// ================================================================
function getClub(clubid)
{
	// Establish local variables
	var theclub = new Object();
	var sclub = new String(clubid);
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	//
	theclub.clubname = new String("").toString();
	theclub.address1 = new String("").toString();
	theclub.address2 = new String("").toString();
	theclub.address3 = new String("").toString();
	theclub.address4 = new String("").toString();
	theclub.town = new String("").toString();
	theclub.county = new String("").toString();
	theclub.postcode = new String("").toString();
	theclub.telephone = new String("").toString();
	theclub.fax = new String("").toString();
	theclub.email = new String("").toString();
	theclub.contact = new String("").toString();
	theclub.totaltennis = new String("").toString();
	theclub.cluburl = new String("").toString();
	theclub.comments = new String("").toString();
	theclub.email2 = new String("").toString();
	theclub.contact2 = new String("").toString();
	theclub.telephone2 = new String("").toString();
	theclub.email3 = new String("").toString();
	theclub.contact3 = new String("").toString();
	theclub.telephone3 = new String("").toString();
	//
	dbconnect=Application("hamptonsportsdb");
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	SQL1 = new String("SELECT * from otherclubs where clubid="+sclub).toString();
	RS = Conn.Execute(SQL1);
	if (! RS.EOF)
	{
		theclub.clubname = new String(RS("clubname")).toString();
		if (theclub.clubname == "null")
			theclub.clubname = new String("").toString();
		theclub.address1 = new String(RS("address1")).toString();
		if (theclub.address1 == "null")
			theclub.address1 = new String("").toString();
		theclub.address2 = new String(RS("address2")).toString();
		if (theclub.address2 == "null")
			theclub.address2 = new String("").toString();
		theclub.address3 = new String(RS("address3")).toString();
		if (theclub.address3 == "null")
			theclub.address3 = new String("").toString();
		theclub.address4 = new String(RS("address4")).toString();
		if (theclub.address4 == "null")
			theclub.address4 = new String("").toString();
		theclub.town = new String(RS("town")).toString();
		if (theclub.town == "null")
			theclub.town = new String("").toString();
		theclub.county = new String(RS("county")).toString();
		if (theclub.county == "null")
			theclub.county = new String("").toString();
		theclub.postcode = new String(RS("postcode")).toString();
		if (theclub.postcode == "null")
			theclub.postcode = new String("").toString();
		theclub.telephone = new String(RS("telephone")).toString();
		if (theclub.telephone == "null")
			theclub.telephone = new String("").toString();
		theclub.fax = new String(RS("fax")).toString();
		if (theclub.fax == "null")
			theclub.fax = new String("").toString();
		theclub.email = new String(RS("email")).toString();
		if (theclub.email == "null")
			theclub.email = new String("").toString();
		theclub.contact = new String(RS("contact")).toString();
		if (theclub.contact == "null")
			theclub.contact = new String("").toString();
		theclub.totaltennis = new String(RS("totaltennis")).toString();
		if (theclub.totaltennis == "null")
			theclub.totaltennis = new String("").toString();
		theclub.cluburl = new String(RS("cluburl")).toString();
		if (theclub.cluburl == "null")
			theclub.cluburl = new String("").toString();
		theclub.comments = new String(RS("comments")).toString();
		if (theclub.comments == "null")
			theclub.comments = new String("").toString();
		theclub.email2 = new String(RS("email2")).toString();
		if (theclub.email2 == "null")
			theclub.email2 = new String("").toString();
		theclub.contact2 = new String(RS("contact2")).toString();
		if (theclub.contact2 == "null")
			theclub.contact2 = new String("").toString();
		theclub.telephone2 = new String(RS("telephone2")).toString();
		if (theclub.telephone2 == "null")
			theclub.telephone2 = new String("").toString();
		theclub.email3 = new String(RS("email3")).toString();
		if (theclub.email3 == "null")
			theclub.email3 = new String("").toString();
		theclub.contact3 = new String(RS("contact3")).toString();
		if (theclub.contact3 == "null")
			theclub.contact3 = new String("").toString();
		theclub.telephone3 = new String(RS("telephone3")).toString();	
		if (theclub.telephone3 == "null")
			theclub.telephone3 = new String("").toString();
	}
	return(theclub);
}
</script>
