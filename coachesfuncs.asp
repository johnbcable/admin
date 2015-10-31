<script language=JScript charset="utf-8" CODEPAGE="65001" runat=Server>
//
//  getCoach(memberid)		       		returns coach object
//  setCoach(coachobject)				returns boolean
//  printCoach(coachobject)				returns string
//  newCoach(memberid)					returns boolean
//  deleteCoach(memberid)			returns SQL used for deletion    
//
// ================================================================

function getCoach(memberid, debugflag)
{
	debugflag = debugflag || false;
	// Establish local variables
	var thecoach = new Object();
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	var mDateObj, dummy1;
	// Set up defaults for the coach object

	//
	dbconnect=Application("hamptonsportsdb");
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	// Need to query particular coachnament entry from coachid
	SQL1 = new String("SELECT * FROM coaches WHERE uniqueref="+memberid).toString();
	Response.Write("<br />"+SQL1+"<br />");
	RS = Conn.Execute(SQL1);
	thecoach.uniqueref = memberid;
	thecoach.surname = new String("NOT FOUND").toString();
	while (! RS.EOF)
	{
		thecoach.id = Trim(new String(RS("memberid")).toString());
		thecoach.grade = Trim(new String(RS("membergrade")).toString());
 		thecoach.surname= Trim(new String(RS("surname")).toString());
		thecoach.forename1= Trim(new String(RS("forename1")).toString());
		// Set initials to empty string as we dont use these any more
		thecoach.initials = new String("").toString();
		thecoach.gender = Trim(new String(RS("gender")).toString());
		thecoach.address1= Trim(new String(RS("address1")).toString());
		thecoach.address2= Trim(new String(RS("address2")).toString());
		thecoach.address3= Trim(new String(RS("address3")).toString());
		thecoach.address4= Trim(new String(RS("address4")).toString());
		thecoach.postcode= Trim(new String(RS("postcode")).toString());
		thecoach.homephone= Trim(new String(RS("homephone")).toString());
		thecoach.mobile= Trim(new String(RS("mobilephone")).toString());
		thecoach.email= Trim(new String(RS("email")).toString());
		thecoach.webpassword = Trim(new String(RS("webpassword")).toString());
		thecoach.pool = Trim(new String(RS("pool")).toString());
		thecoach.maxitennis = Trim(new String(RS("maxiteam")).toString());
		thecoach.britishtennisno = Trim(new String(RS("britishtennisno")).toString());
		thecoach.adultcontact = Trim(new String(RS("adultcontact")).toString());
		thecoach.adultrelationship = Trim(new String(RS("adultrelationship")).toString());
		thecoach.adultphone = Trim(new String(RS("adultphone")).toString());
		thecoach.adultmobile = Trim(new String(RS("adultmobile")).toString());
		thecoach.specialcare = Trim(new String(RS("specialcare")).toString());
		thecoach.photoconsent = Trim(new String(RS("photoconsent")).toString());
		thecoach.paid = StringFromDB(new String(RS("paid")));
		thecoach.mailing = Trim(new String(RS("mailing")).toString());
		thecoach.internalleague = Trim(new String(RS("internalleague")).toString());
		thecoach.onlinebookingid = Trim(new String(RS("onlinebookingid")).toString());
		thecoach.onlinebookingpin = Trim(new String(RS("onlinebookingpin")).toString());
		thecoach.iscoach = Trim(new String(RS("iscoach")).toString());
		// 	mWimbledonDraw = Trim(new String(RS("wimbledondraw")));
		dummy=new Number(RS("webaccess")).valueOf();
		thecoach.webaccess=dummy.valueOf();
		// now the date bits of the member record
		// First, date of birth
		mDateObj=new Date(RS("dob"));
		dummy1 = mDateObj.valueOf();
		if (dummy1 == 0) // no date in database
			thecoach.dob = "";
		else
			thecoach.dob = ddmmyyyy(mDateObj);
		// Second, joining date
		mDateObj=new Date(RS("joined"));
		dummy1 = mDateObj.valueOf();
		if (dummy1 == 0) // no date in database
			thecoach.joined = "";
		else
			thecoach.joined = ddmmyyyy(mDateObj);
		// Third, date of leaving the club
		mDateObj=new Date(RS("dateleft"));
		dummy1 = mDateObj.valueOf();
		if (dummy1 == 0) // no date in database
			thecoach.left = "";
		else
			thecoach.left = ddmmyyyy(mDateObj);
		// End of date bits
		RS.MoveNext();
	}
	if (thecoach.grade=="null" || thecoach.grade=="undefined")
		thecoach.grade="";
	if (thecoach.surname=="null" || thecoach.surname=="undefined")
		thecoach.surname="";
	if (thecoach.forename1=="null" || thecoach.forename1=="undefined")
		thecoach.forename1="";
	if (thecoach.initials=="null" || thecoach.initials=="undef")
		thecoach.initials="";
	if (thecoach.address1=="null" || thecoach.address1=="undefined")
		thecoach.address1="";
	if (thecoach.address2=="null" || thecoach.address2=="undefined")
		thecoach.address2="";
	if (thecoach.address3=="null" || thecoach.address3=="undefined")
		thecoach.address3="";
	if (thecoach.address4=="null" || thecoach.address4=="undefined")
		thecoach.address4="";
	if (thecoach.postcode=="null" || thecoach.postcode=="undefined")
		thecoach.postcode="";
	if (thecoach.homephone=="null" || thecoach.homephone=="undefined")
		thecoach.homephone="";
	if (thecoach.mobile=="null" || thecoach.mobile=="undefined")
		thecoach.mobile="";
	if (thecoach.email=="null" || thecoach.email=="undefined")
		thecoach.email="";
	if (thecoach.webpassword=="null" || thecoach.webpassword=="undefined")
		thecoach.webpassword="tennis";
	if (thecoach.pool == "null" || thecoach.pool == "undefined")
		thecoach.pool = "";
	if (thecoach.maxitennis == "null" || thecoach.maxitennis == "undefined")
		thecoach.maxitennis = "";
	if (thecoach.britishtennisno == "null" || thecoach.britishtennisno=="undefined")
		thecoach.britishtennisno = "";
	if (thecoach.adultcontact == "null" || thecoach.adultcontact=="undefined")
		thecoach.adultcontact = "";
	if (thecoach.adultrelationship == "null" || thecoach.adultrelationship=="undefined")
		thecoach.adultrelationship = "";
	if (thecoach.adultphone == "null" || thecoach.adultphone == "undefined" )
		thecoach.adultphone = "";
	if (thecoach.adultmobile == "null" || thecoach.adultmobile == "undefined" )
		thecoach.adultmobile = "";
	if (thecoach.specialcare == "null" || thecoach.specialcare == "undefined" )
		thecoach.specialcare = "";
	if (thecoach.photoconsent == "null" || thecoach.photoconsent == "undefined" || thecoach.photoconsent == "")
		thecoach.photoconsent = new String("N").toString();
	if (thecoach.paid == "null" || thecoach.paid == "undefined" || thecoach.paid == "")
		thecoach.paid = new String("N").toString();
	if (thecoach.webaccess < 20)
		thecoach.webaccess = 20;
	if (thecoach.mailing=="null" || thecoach.mailing=="" || thecoach.mailing=="undefined")
			thecoach.mailing = "N";	
	if (thecoach.internalleague=="null" || thecoach.internalleague=="" || thecoach.internalleague=="undefined")
			thecoach.internalleague = "";	
	if (thecoach.onlinebookingid=="null" || thecoach.onlinebookingid=="")
	{
	 	  dummy1 = new Number(memberid)+5000;
			thecoach.onlinebookingid = new String(dummy1).toString();	
	}
	if (thecoach.onlinebookingpin=="null" || thecoach.onlinebookingpin=="" || thecoach.onlinebookingpin=="undefined")
			thecoach.onlinebookingpin = new String("").toString();	
	if (thecoach.gender=="null" || thecoach.gender=="" || thecoach.gender=="undefined")
			thecoach.gender = "";	
	if (thecoach.iscoach == "null" || thecoach.iscoach == "undefined" || thecoach.iscoach == "")
		thecoach.iscoach = new String("N").toString();

	RS.Close();
	Conn.Close();
	RS = null;
	Conn = null;

	return(thecoach);
}
// ================================================================
function setCoach(coachobj, debugflag)
{
	// Establish local variables
	var sMember = new String(coachobj.memberid);
	var RS, RS2, Conn, SQL1, SQL2, uniqref;
	var mDateObj, dummy1;
	var resultObj = new Object();
	var e;
	//
	resultObj.result = true;
	resultObj.errno = 0;
	resultObj.description = new String("").toString();
	//
	if (coachobj.grade=="null" || coachobj.grade =="undefined")
		coachobj.grade="";
	if (coachobj.surname=="null" || coachobj.surname =="undefined")
		coachobj.surname="";
	if (coachobj.forename1=="null" || coachobj.forename1 =="undefined")
		coachobj.forename1="";
	// Set initials to empty string as we dont use these any more
	coachobj.initials="";
	if (coachobj.address1=="null" || coachobj.address1 =="undefined")
		coachobj.address1="";
	if (coachobj.address2=="null" || coachobj.address2 =="undefined")
		coachobj.address2="";
	if (coachobj.address3=="null" || coachobj.address3 =="undefined")
		coachobj.address3="";
	if (coachobj.address4=="null" || coachobj.address4 =="undefined")
		coachobj.address4="";
	if (coachobj.postcode=="null" || coachobj.postcode =="undefined")
		coachobj.postcode="";
	if (coachobj.homephone=="null" || coachobj.homephone =="undefined")
		coachobj.homephone="";
	if (coachobj.mobile=="null" || coachobj.mobile =="undefined")
		coachobj.mobile="";
	if (coachobj.email=="null" || coachobj.email =="undefined")
		coachobj.email="";
	if (coachobj.webpassword=="null" || coachobj.webpassword =="undefined")
		coachobj.webpassword="tennis";
	if (coachobj.pool == "null" || coachobj.pool =="undefined")
		coachobj.pool = "";
	if (coachobj.maxitennis == "null" || coachobj.maxitennis =="undefined")
		coachobj.maxitennis = "";
	if (coachobj.britishtennisno == "null" || coachobj.britishtennisno =="undefined")
		coachobj.britishtennisno = "";
	if (coachobj.adultcontact == "null" || coachobj.adultcontact =="undefined")
		coachobj.adultcontact = "";
	if (coachobj.adultrelationship == "null" || coachobj.adultrelationship =="undefined")
		coachobj.adultrelationship = "";
	if (coachobj.adultphone == "null" || coachobj.adultphone =="undefined")
		coachobj.adultphone = "";
	if (coachobj.adultmobile == "null" || coachobj.adultmobile =="undefined")
		coachobj.adultmobile = "";
	if (coachobj.specialcare == "null" || coachobj.specialcare =="undefined")
		coachobj.specialcare = "";
	if (coachobj.photoconsent == "null" || coachobj.photoconsent == "undefined" || coachobj.photoconsent == "")
		coachobj.photoconsent = "N";
	if (coachobj.paid == "null" || coachobj.paid == "undefined" || coachobj.paid == "")
		coachobj.paid = "N";
	if (coachobj.internalleague=="null" || coachobj.internalleague=="" || coachobj.internalleague=="undefined")
			coachobj.internalleague = "";	
	if (coachobj.onlinebookingid=="null" || coachobj.onlinebookingid=="" || coachobj.onlinebookingid=="undefined")
	{
		if (coachobj.grade == "Adult" || coachobj.grade == "18-25s" || coachobj.grade == "Junior")  
			 coachobj.onlinebookingid = new String(new Number(coachobj.uniqueref)+5000).toString();
		else
			 coachobj.onlinebookingid = new String("0").toString();
	}
	if (coachobj.onlinebookingpin=="null" || coachobj.onlinebookingpin=="" || coachobj.onlinebookingpin=="undefined")
			coachobj.onlinebookingpin = new String("").toString();
	if (coachobj.webaccess < 20)
		coachobj.webaccess = 20;
	if (coachobj.mailing=="null" || coachobj.mailing =="undefined" || coachobj.mailing=="")
			coachobj.mailing = "N";	
	if (coachobj.dob=="null" || coachobj.dob =="undefined")
			coachobj.dob = "";	
	if (coachobj.left=="null" || coachobj.left =="undefined")
			coachobj.left = "";	
	if (coachobj.joined=="null" || coachobj.joined =="undefined")
			coachobj.joined = "";	
	if (coachobj.gender=="null" || coachobj.gender =="undefined" || coachobj.gender=="")
			coachobj.gender = "";	
	if (coachobj.iscoach == "null" || coachobj.iscoach == "undefined" || coachobj.iscoach == "")
		coachobj.iscoach = "N";
	
	var dbconnect=Application("hamptonsportsdb"); 
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	SQLstart = new String("UPDATE members ")
	SQLend = new String(" WHERE uniqueref="+coachobj.uniqueref).toString();
	SQLmiddle = new String("SET ").toString();
	SQLmiddle += " memberid='"+coachobj.id+"',";
	SQLmiddle += " membergrade='"+coachobj.grade+"',";
	SQLmiddle += " surname='"+coachobj.surname+"',";
	SQLmiddle += " forename1='"+coachobj.forename1+"',";
	SQLmiddle += " initials='"+coachobj.initials+"',";
	SQLmiddle += " gender='"+coachobj.gender+"', ";
	SQLmiddle += " address1='"+coachobj.address1+"',";
	SQLmiddle += " address2='"+coachobj.address2+"',";
	SQLmiddle += " address3='"+coachobj.address3+"',";
	SQLmiddle += " address4='"+coachobj.address4+"',";
	SQLmiddle += " postcode='"+coachobj.postcode+"',";
	SQLmiddle += " homephone='"+coachobj.homephone+"',";
	SQLmiddle += " mobilephone='"+coachobj.mobile+"',";
	SQLmiddle += " email='"+coachobj.email+"',";
	SQLmiddle += " webpassword='"+coachobj.webpassword+"',";
	SQLmiddle += " pool='"+coachobj.pool+"',";
	SQLmiddle += " maxiteam='"+coachobj.maxitennis+"',";
	SQLmiddle += " britishtennisno='"+coachobj.britishtennisno+"',";
	SQLmiddle += " adultcontact='"+coachobj.adultcontact+"',";
	SQLmiddle += " adultrelationship='"+coachobj.adultrelationship+"',";
	SQLmiddle += " adultphone='"+coachobj.adultphone+"',";
	SQLmiddle += " adultmobile='"+coachobj.adultmobile+"',";
	SQLmiddle += " specialcare='"+coachobj.specialcare+"',";
	SQLmiddle += " photoconsent='"+coachobj.photoconsent+"',";
	SQLmiddle += " paid='"+coachobj.paid+"',";
	SQLmiddle += " webaccess="+coachobj.webaccess+",";
	// Now do date fields. If null dont insert them as part of the update clause
	//  Access doesnt like setting date fields to ''
	if (! (coachobj.dob == ""))
		SQLmiddle += " dob='"+coachobj.dob+"',";
	else
		SQLmiddle += " dob=null,";
	if (! (coachobj.joined == ""))
		SQLmiddle += " joined='"+coachobj.joined+"',";
	else
		SQLmiddle += " joined=null,";
	if (! (coachobj.left == ""))
		SQLmiddle += " dateleft='"+coachobj.left+"',";
	else
		SQLmiddle += " dateleft=null,";
	SQLmiddle += " mailing='"+coachobj.mailing+"', ";
	SQLmiddle += " internalleague='"+coachobj.internalleague+"', ";
	if (coachobj.onlinebookingid != "0")
		 SQLmiddle += " onlinebookingid="+coachobj.onlinebookingid+", ";
	SQLmiddle += " onlinebookingpin='"+coachobj.onlinebookingpin+"', ";
	SQLmiddle += " iscoach='"+coachobj.iscoach+"' ";
//
	SQL1 = new String(SQLstart+SQLmiddle+SQLend).toString();
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

// ================================================================
function printCoach(coachObj)
{
	// Establish local variables
	var sReport = new String("").toString();
		sReport += "Coachnament ID: "+coachObj.coachnamentid +"<br />";
		sReport += "Coachnament title: "+coachObj.coachtitle +"<br />";
   		sReport += "Coachnament start date: "+coachObj.coachstart +"<br />";
		sReport += "Coachnament end date: "+coachObj.coachend +"<br />";
		sReport += "Coachnament finals day: "+coachObj.coachfinalsday +"<br />";
		sReport += "Who coachnament is for: "+coachObj.coachwho +"<br />";
		sReport += "Who to contact about this coachnament: "+coachObj.coachcontact +"<br />";
		sReport += "Start time on start day: "+coachObj.coachstarttime +"<br />";
		sReport += "End time on end date: "+coachObj.coachendtime +"<br />";
		sReport += "Finals day start time: "+coachObj.finalsstarttime +"<br />";
		sReport += "Finals day end time: "+coachObj.finalsendtime +"<br />";
		sReport += "Cost to enter: "+coachObj.coachcost +"<br />";
		sReport += "Web URL: "+coachObj.coachurl +"<br />";
		sReport += "Category of coachnament: "+coachObj.coachcategory +"<br />";
		sReport += "Additional information link: "+coachObj.coachblurb +"<br />";

	return(sReport);
}

// ================================================================
function newCoach(memberid,debugflag)
{
	// Insert new skeleton coach row into table based on the data from 
	// the member with a uniqueref of memberid and return the memberid
	// to caller
	debugflag = debugflag || false;
	// Establish local variables
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	var SQLstart, SQLmiddle, SQLend;
	var coachid;
	//
	dbconnect=Application("hamptonsportsdb");
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);

	// Generate random 5-character reference string

   	var text = "";
    var possible = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";

    for( var i=0; i < 5; i++ )
        text += possible.charAt(Math.floor(Math.random() * possible.length));

	SQLstart = new String("INSERT INTO coaches ([coachtitle], [coachstart],[coachend],[coachcategory]) ").toString();
	SQLend = new String("").toString();
	SQLmiddle = new String("VALUES ('NEW TOURNAMENT', '01/01/2099', '01/01/2199', '"+text+"')").toString();

	SQL1 = new String(SQLstart+SQLmiddle+SQLend).toString();

	// Insert dummy row 

	RS = Conn.Execute(SQL1);

	// Now get the eventid back
	coachid = memberid;
	SQL2 = new String("SELECT coachnamentid FROM coachnaments WHERE coachcategory = '"+text+"'").toString();

	RS = Conn.Execute(SQL2);
	if (! RS.EOF) {
		coachid = new Number(RS("coachnamentid"));
	}
	RS.Close();
	Conn.Close();
	RS = null;
	Conn = null;

	return(coachid);
}

// ================================================================
function deleteCoach(memberid, debugflag)
{
	debugflag = debugflag || false;
	// Establish local variables
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	var mDateObj, dummy1;
	//
	dbconnect=Application("hamptonsportsdb");
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	SQL1 = new String("DELETE FROM coaches WHERE uniqueref = "+memberid);

	if ( ! debugflag) {
		RS = Conn.Execute(SQL1);
		// RS.Close();
	}

	Conn.Close();
	RS = null;
	Conn = null;
	
	return(SQL1);
}
// ================================================================
</script>