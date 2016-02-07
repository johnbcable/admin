
<!--
METADATA
TYPE="typelib"
UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"
NAME="CDO for Windows 2000 Library"
-->
<script language="JScript" CODEPAGE="65001" runat=Server>
//
//  newMailObject()                             returns CDOSYS.Message object
//  getRecentEmails( howmany)          					returns string
//  getRecentMessages( howmany )          				returns string
//  getEmailByFile(fileref)            					returns email object
//  getEmailByID(mailid)               					returns email object
//  getMessageByID(messageid)							returns message object
//  processSquadMergeFields(message,fixture) 			returns string
//  processPersonalMergeFields(message,uniqueref)		returns string
//  processGeneralMergeFields(message)					returns string
//  processMAXITennisMergeFields(message,uniqueref)		returns string
//
function newMailObject()
{

  var cdoConfig = Server.CreateObject("CDO.Configuration");
	
  var cdoFields = cdoConfig.Fields;
  cdoFields.Item(cdoSendUsingMethod) = cdoSendUsingPort;
  cdoFields.Item(cdoSMTPServerPort) = 25;
  cdoFields.Item(cdoSMTPServer) = "SMTPMAIL";
  cdoFields.Item(cdoSMTPconnectiontimeout) = 30;
  cdoFields.Update();

/*	
// .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
// .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
// .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
// .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
// .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
// .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "user@domain.com"
// .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "PASSWORD"
*/	
	
	objCDOMail = Server.CreateObject("CDO.Message");
	
  	objCDOMail.Configuration = cdoConfig;

	objCDOMail.BodyPart.charset = "utf-8";
	
	return(objCDOMail);

}
//======================================================================
function getRecentEmails(howmany)
{
	// Establish local variables
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	var debugging = false;
	var editline, emailsubject;
	var emaillist = new String("").toString();
	var emailknt = 0;
	var totaltoget = new Number(howmany);
	var emaildir = new String("emails").toString();
	totaltoget = totaltoget.valueOf();
	// Set up database connections
	dbconnect=Application("hamptonsportsdb"); 
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	// Need to query last 5 emails sent and return as string
	// title<hr><br>event1<br>event2<br> .... eventN<hr>
	SQL1 = new String("SELECT * from sentemails").toString();
	emaillist += '<ol id="emails"><div id="emailheader"><h3><a href="email_list.asp">Recent Emails</a></h3></div>';
	RS = Conn.Execute(SQL1);
	emailknt = 0;
	while (! RS.EOF && emailknt < totaltoget)
	{
		emaillist += '<li><p><a href="'+emaildir+'/'+RS("emailfile")+'" target="_emailframe">'+RS("subject")+'</a></p>';
		emailknt++;
		RS.MoveNext();
	}
	emaillist += "</ol>";
	RS.Close();
	return (emaillist);
}
//=====================================================================
function getRecentMessages(howmany)
{
	// Establish local variables
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	var debugging = false;
	var editline, emailsubject;
	var qmesslist = new String("").toString();
	var qmessknt = 0;
	var totaltoget = new Number(howmany);
	totaltoget = totaltoget.valueOf();
	// Set up database connections
	dbconnect=Application("hamptonsportsdb"); 
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	// Need to query last 5 quick messages sent and return as string
	// title<hr><br>event1<br>event2<br> .... eventN<hr>
	SQL1 = new String("SELECT * from quickmessages order by messageid desc").toString();
	qmesslist += '<ol id="quickmessages"><div id="qmessageheader"><h3><a href="qmess_list.asp">Recent Quick Messages</a></h3></div>';
	RS = Conn.Execute(SQL1);
	qmessknt = 0;
	while (! RS.EOF && qmessknt < totaltoget)
	{
		qmesslist += '<li><p><a href="showqmess.asp?id='+RS("messageid")+'" target="_blank">'+RS("subject")+'</a></p>';
		qmessknt++;
		RS.MoveNext();
	}
	qmesslist += "</ol>";
	RS.Close();
	return (qmesslist);
}
//=====================================================================
function getEmailByFile(fileref)
{
	// Establish local variables
	var emailobj = new Object();
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	// Set up database connections
	dbconnect=Application("hamptonsportsdb"); 
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	// Need to query particular email database entry from fileref
	SQL1 = new String("SELECT * from sentemails where emailfile='"+fileref+"'").toString();
	RS = Conn.Execute(SQL1);
	emailobj.subject=new String("NOT FOUND").toString();
	while (! RS.EOF)
	{
		emailobj.emailfile = new String(RS("emailfile")).toString();
		emailobj.subject = new String(RS("subject")).toString();
		emailobj.senton = new String(RS("dateofemail")).toString();
		emailobj.attach1 = new String(RS("attach1")).toString();
		emailobj.attach2 = new String(RS("attach2")).toString();
		emailobj.attach3 = new String(RS("attach3")).toString();
		emailobj.circulation = new String(RS("circulation")).toString();
		emailobj.numbersent = new String(RS("number_sent")).toString();
		emailobj.emailid = new String(RS("emailid")).toString();
		RS.MoveNext();
	}
	RS.Close();
	return (emailobj);
}
//=====================================================================
function getEmailByID(mailid)
{
	// Establish local variables
	var emailobj = new Object();
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	// Set up database connections
	dbconnect=Application("hamptonsportsdb"); 
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	// Need to query particular email database entry from fileref
	SQL1 = new String("SELECT * from sentemails where emailid="+mailid).toString();
	RS = Conn.Execute(SQL1);
	emailobj.subject=new String("NOT FOUND").toString();
	while (! RS.EOF)
	{
		emailobj.emailfile = new String(RS("emailfile")).toString();
		emailobj.subject = new String(RS("subject")).toString();
		emailobj.senton = new String(RS("dateofemail")).toString();
		emailobj.attach1 = new String(RS("attach1")).toString();
		emailobj.attach2 = new String(RS("attach2")).toString();
		emailobj.attach3 = new String(RS("attach3")).toString();
		emailobj.circulation = new String(RS("circulation")).toString();
		emailobj.numbersent = new String(RS("number_sent")).toString();
		emailobj.emailid = new String(RS("emailid")).toString();
		RS.MoveNext();
	}
	RS.Close();
	return (emailobj);
}
//=====================================================================
function getMessageByID(messageid)
{
	// Establish local variables
	var messageobj = new Object();
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	// Set up database connections
	dbconnect=Application("hamptonsportsdb"); 
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	// Need to query particular email database entry from fileref
	SQL1 = new String("SELECT * from quickmessages where messageid="+messageid).toString();
	RS = Conn.Execute(SQL1);
	messageobj.subject=new String("NOT FOUND").toString();
	while (! RS.EOF)
	{
		messageobj.subject = new String(RS("subject")).toString();
		messageobj.messagetext = new String(RS("messagetext")).toString();
		messageobj.senton = new String(RS("sent_on")).toString();
		messageobj.circulation = new String(RS("circulation")).toString();
		messageobj.numbersent = new String(RS("number_sent")).toString();
		messageobj.emailid = new String(RS("messageid")).toString();
		messageobj.sender = new String(RS("sender")).toString();
		messageobj.postid = new String(RS("postid")).toString();
		RS.MoveNext();
	}
	RS.Close();
	return (messageobj);
}

//=====================================================================
function processSquadMergeFields(message,fixture)
{
	// Establish local variables
	var sMessage = new String(message);
	var oFixture = new Object();
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	//
	oFixture = getFixture(fixture);
	//
	// Now, do any text substitutions for this message
	sMessage = new String(sMessage.replace("ZZteamZZ",oFixture.teamname)).toString();
	sMessage = new String(sMessage.replace("ZZvenueZZ",oFixture.venue)).toString();
	sMessage = new String(sMessage.replace("ZZfixturedateZZ",oFixture.fixturedate)).toString();
	sMessage = new String(sMessage.replace("ZZopponentsZZ",oFixture.opponents)).toString();
	sMessage = new String(sMessage.replace("ZZpair1ZZ",oFixture.pair1)).toString();
	sMessage = new String(sMessage.replace("ZZpair2ZZ",oFixture.pair2)).toString();
	return(sMessage);
}
//=====================================================================
function processPersonalMergeFields(message,personid)
{
	// Establish local variables
	var sMessage = new String(message);
	var sPerson = new String(personid);
	var oMember = new Object();
	var dummy;
	var fulladdress = new String("").toString();

	oMember = getMember(sPerson);

	// Create the fulladdress field
	fulladdress += oMember.address1;
	if (! (oMember.address2 == "")) {
		fulladdress += "," + oMember.address2;
	}
	if (! (oMember.address3 == "")) {
		fulladdress += "," + oMember.address3;
	}
	if (! (oMember.address4 == "")) {
		fulladdress += "," + oMember.address4;
	}
	if (! (oMember.postcode == "")) {
		fulladdress += "," + oMember.postcode;
	}

	sMessage = replaceAll(sMessage, "ZZidZZ", sPerson);
	sMessage = replaceAll(sMessage, "ZZforenameZZ", oMember.forename1);
	sMessage = replaceAll(sMessage, "ZZsurnameZZ", oMember.surname);
	sMessage = replaceAll(sMessage, "ZZaddressZZ", fulladdress);
	sMessage = replaceAll(sMessage, "ZZgenderZZ", oMember.gender == "M" ? "Male" : "Famale");
	sMessage = replaceAll(sMessage, "ZZemailZZ", oMember.email);
	sMessage = replaceAll(sMessage, "ZZhomephoneZZ", oMember.homephone);
	sMessage = replaceAll(sMessage, "ZZmobileZZ", oMember.mobile);
	sMessage = replaceAll(sMessage, "ZZdobZZ", oMember.dob);
	sMessage = replaceAll(sMessage, "ZZemergencynameZZ",oMember.emergencyname);
	sMessage = replaceAll(sMessage, "ZZemergencyphoneZZ",oMember.emergencyphone);
	sMessage = replaceAll(sMessage, "ZZphotoconsentZZ",oMember.photoconsent);
	sMessage = replaceAll(sMessage, "ZZmedicalZZ",oMember.medicalinfo);
	sMessage = replaceAll(sMessage, "ZZsinglesZZ",oMember.singles);
	sMessage = replaceAll(sMessage, "ZZdoublesZZ",oMember.doubles);
	sMessage = replaceAll(sMessage, "ZZdoublespartnerZZ",oMember.doublespartner);
	sMessage = replaceAll(sMessage, "ZZmixedZZ",oMember.mixeddoubles);
	sMessage = replaceAll(sMessage, "ZZmixedpartnerZZ",oMember.mixedpartner);
	sMessage = replaceAll(sMessage, "ZZmixedplateZZ",oMember.mixedplate);
	sMessage = replaceAll(sMessage, "ZZvets40ZZ",oMember.vetsdoubles);
	sMessage = replaceAll(sMessage, "ZZvets40partnerZZ",oMember.vetspartner);
	sMessage = replaceAll(sMessage, "ZZvets55ZZ",oMember.supervetsdoubles);
	sMessage = replaceAll(sMessage, "ZZvets55partnerZZ",oMember.supervetspartner);
	sMessage = replaceAll(sMessage, "ZZsummerfinalsZZ",oMember.summerfinalsday);
	sMessage = replaceAll(sMessage, "ZZautumnfinalsZZ",oMember.autumnfinalsday);
	sMessage = replaceAll(sMessage, "ZZteamhelpZZ",oMember.teamhelp);
	sMessage = replaceAll(sMessage, "ZZsocialhelpZZ",oMember.socialhelp);
	sMessage = replaceAll(sMessage, "ZZadminhelpZZ",oMember.adminhelp);
	sMessage = replaceAll(sMessage, "ZZarden9helpZZ",oMember.arden9help);
	sMessage = replaceAll(sMessage, "ZZonlinebookingidZZ",oMember.onlinebookingid);
	sMessage = replaceAll(sMessage, "ZZbtmnumberZZ", oMember.britishtennisno);
	sMessage = replaceAll(sMessage, "ZZwimbledonweekdayZZ", oMember.wimbledonweekday);
	sMessage = replaceAll(sMessage, "ZZwimbledonweekendZZ", oMember.wimbledonweekend);
	sMessage = replaceAll(sMessage, "ZZbritishtennisnoZZ", oMember.britishtennisno);


	dummy = new String(oMember.onlinebookingpin).toString();
	if (dummy.length > 4)
		sMessage = replaceAll(sMessage, "ZZonlinebookingpinZZ","Already allocated and changed by member");
	else
		sMessage = replaceAll(sMessage, "ZZonlinebookingpinZZ",oMember.onlinebookingpin);
	
	return(sMessage);
}
//=====================================================================
function processMAXITennisMergeFields(message,personid)
{
	// Establish local variables
	var sMessage = new String(message);
	var sPerson = new String(personid);
	var sTeam = new String("").toString();
	var sCaptains = new String("").toString();
	var oMember = new Object();
	
	oMember = getMember(sPerson);

	sTeam += oMember.maxitennis;
	sCaptains += getMaxiCaptainsAsString(sTeam)
	
	sMessage = new String(sMessage.replace("ZZmaxiteamZZ",sTeam)).toString();
	sMessage = new String(sMessage.replace("ZZmaxicaptainsZZ",sCaptains)).toString();

	return(sMessage);
}
//=====================================================================
function processGeneralMergeFields(message)
{
	// Establish local variables
	var sMessage = new String(message);

	return(sMessage);
}
</script>
