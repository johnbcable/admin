<script language=JScript runat=Server>
//
//  getCoach(Coachid)           returns Coach Object
//  setCoach(CoachObject)		  returns result Object
//  printCoach(CoachObject)	  returns string
//  deleteCoach(Coachid)        returns result Object 
//
// ================================================================
function getCoach(Coachid)
{
	// Establish local variable
	var thecoach = new Object();
	var sCoach = new String(Coachid);
	var RS, RS2, Conn, SQL1, SQL2, uniqref;
	var mDateObj, dummy1;
	var mDob;
	//
	thecoach.Coachid = new String(sCoach).toString();
	//
	var dbconnect=Application("hamptonsportsdb"); 
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	// 
	SQL1 = new String("SELECT * FROM coaches WHERE uniqueref = " + sCoach);
	RS = Conn.Execute(SQL1);
	// Retrieve database values for Coach
	mDob = new String("").toString();
	thecoach.uniqueref = sCoach;
	while (! RS.EOF)
	{
		thecoach.id = Trim(new String(RS("uniqueref")).toString());
		thecoach.membergrade = Trim(new String(RS("membergrade")).toString());
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
		// now the date bits of the Coach record
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
	if (thecoach.membergrade=="null" || thecoach.membergrade=="undefined")
		thecoach.membergrade="";
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
	 	  dummy1 = new Number(Coachid)+5000;
			thecoach.onlinebookingid = new String(dummy1).toString();	
	}
	if (thecoach.onlinebookingpin=="null" || thecoach.onlinebookingpin=="" || thecoach.onlinebookingpin=="undefined")
			thecoach.onlinebookingpin = new String("").toString();	
	if (thecoach.gender=="null" || thecoach.gender=="" || thecoach.gender=="undefined")
			thecoach.gender = "";	
	if (thecoach.iscoach == "null" || thecoach.iscoach == "undefined" || thecoach.iscoach == "")
		thecoach.iscoach = new String("N").toString();
	
	return(thecoach);
}
// ================================================================
function setCoach(Coachobj)
{
	// Establish local variables
	var sCoach = new String(Coachobj.Coachid);
	var RS, RS2, Conn, SQL1, SQL2, uniqref;
	var mDateObj, dummy1;
	var resultObj = new Object();
	var e;
	//
	resultObj.result = true;
	resultObj.errno = 0;
	resultObj.description = new String("").toString();
	//
	if (Coachobj.grade=="null" || Coachobj.grade =="undefined")
		Coachobj.grade="";
	if (Coachobj.surname=="null" || Coachobj.surname =="undefined")
		Coachobj.surname="";
	if (Coachobj.forename1=="null" || Coachobj.forename1 =="undefined")
		Coachobj.forename1="";
	// Set initials to empty string as we dont use these any more
	Coachobj.initials="";
	if (Coachobj.address1=="null" || Coachobj.address1 =="undefined")
		Coachobj.address1="";
	if (Coachobj.address2=="null" || Coachobj.address2 =="undefined")
		Coachobj.address2="";
	if (Coachobj.address3=="null" || Coachobj.address3 =="undefined")
		Coachobj.address3="";
	if (Coachobj.address4=="null" || Coachobj.address4 =="undefined")
		Coachobj.address4="";
	if (Coachobj.postcode=="null" || Coachobj.postcode =="undefined")
		Coachobj.postcode="";
	if (Coachobj.homephone=="null" || Coachobj.homephone =="undefined")
		Coachobj.homephone="";
	if (Coachobj.mobile=="null" || Coachobj.mobile =="undefined")
		Coachobj.mobile="";
	if (Coachobj.email=="null" || Coachobj.email =="undefined")
		Coachobj.email="";
	if (Coachobj.webpassword=="null" || Coachobj.webpassword =="undefined")
		Coachobj.webpassword="tennis";
	if (Coachobj.pool == "null" || Coachobj.pool =="undefined")
		Coachobj.pool = "";
	if (Coachobj.maxitennis == "null" || Coachobj.maxitennis =="undefined")
		Coachobj.maxitennis = "";
	if (Coachobj.britishtennisno == "null" || Coachobj.britishtennisno =="undefined")
		Coachobj.britishtennisno = "";
	if (Coachobj.adultcontact == "null" || Coachobj.adultcontact =="undefined")
		Coachobj.adultcontact = "";
	if (Coachobj.adultrelationship == "null" || Coachobj.adultrelationship =="undefined")
		Coachobj.adultrelationship = "";
	if (Coachobj.adultphone == "null" || Coachobj.adultphone =="undefined")
		Coachobj.adultphone = "";
	if (Coachobj.adultmobile == "null" || Coachobj.adultmobile =="undefined")
		Coachobj.adultmobile = "";
	if (Coachobj.specialcare == "null" || Coachobj.specialcare =="undefined")
		Coachobj.specialcare = "";
	if (Coachobj.photoconsent == "null" || Coachobj.photoconsent == "undefined" || Coachobj.photoconsent == "")
		Coachobj.photoconsent = "N";
	if (Coachobj.paid == "null" || Coachobj.paid == "undefined" || Coachobj.paid == "")
		Coachobj.paid = "N";
	if (Coachobj.internalleague=="null" || Coachobj.internalleague=="" || Coachobj.internalleague=="undefined")
			Coachobj.internalleague = "";	
	if (Coachobj.onlinebookingid=="null" || Coachobj.onlinebookingid=="" || Coachobj.onlinebookingid=="undefined")
	{
		if (Coachobj.grade == "Adult" || Coachobj.grade == "18-25s" || Coachobj.grade == "Junior")  
			 Coachobj.onlinebookingid = new String(new Number(Coachobj.uniqueref)+5000).toString();
		else
			 Coachobj.onlinebookingid = new String("0").toString();
	}
	if (Coachobj.onlinebookingpin=="null" || Coachobj.onlinebookingpin=="" || Coachobj.onlinebookingpin=="undefined")
			Coachobj.onlinebookingpin = new String("").toString();
	if (Coachobj.webaccess < 20)
		Coachobj.webaccess = 20;
	if (Coachobj.mailing=="null" || Coachobj.mailing =="undefined" || Coachobj.mailing=="")
			Coachobj.mailing = "N";	
	if (Coachobj.dob=="null" || Coachobj.dob =="undefined")
			Coachobj.dob = "";	
	if (Coachobj.left=="null" || Coachobj.left =="undefined")
			Coachobj.left = "";	
	if (Coachobj.joined=="null" || Coachobj.joined =="undefined")
			Coachobj.joined = "";	
	if (Coachobj.gender=="null" || Coachobj.gender =="undefined" || Coachobj.gender=="")
			Coachobj.gender = "";	
	if (Coachobj.iscoach == "null" || Coachobj.iscoach == "undefined" || Coachobj.iscoach == "")
		Coachobj.iscoach = "N";
	
	var dbconnect=Application("hamptonsportsdb"); 
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	SQLstart = new String("UPDATE Coachs ")
	SQLend = new String(" WHERE uniqueref="+Coachobj.uniqueref).toString();
	SQLmiddle = new String("SET ").toString();
	SQLmiddle += " Coachid='"+Coachobj.id+"',";
	SQLmiddle += " Coachgrade='"+Coachobj.grade+"',";
	SQLmiddle += " surname='"+Coachobj.surname+"',";
	SQLmiddle += " forename1='"+Coachobj.forename1+"',";
	SQLmiddle += " initials='"+Coachobj.initials+"',";
	SQLmiddle += " gender='"+Coachobj.gender+"', ";
	SQLmiddle += " address1='"+Coachobj.address1+"',";
	SQLmiddle += " address2='"+Coachobj.address2+"',";
	SQLmiddle += " address3='"+Coachobj.address3+"',";
	SQLmiddle += " address4='"+Coachobj.address4+"',";
	SQLmiddle += " postcode='"+Coachobj.postcode+"',";
	SQLmiddle += " homephone='"+Coachobj.homephone+"',";
	SQLmiddle += " mobilephone='"+Coachobj.mobile+"',";
	SQLmiddle += " email='"+Coachobj.email+"',";
	SQLmiddle += " webpassword='"+Coachobj.webpassword+"',";
	SQLmiddle += " pool='"+Coachobj.pool+"',";
	SQLmiddle += " maxiteam='"+Coachobj.maxitennis+"',";
	SQLmiddle += " britishtennisno='"+Coachobj.britishtennisno+"',";
	SQLmiddle += " adultcontact='"+Coachobj.adultcontact+"',";
	SQLmiddle += " adultrelationship='"+Coachobj.adultrelationship+"',";
	SQLmiddle += " adultphone='"+Coachobj.adultphone+"',";
	SQLmiddle += " adultmobile='"+Coachobj.adultmobile+"',";
	SQLmiddle += " specialcare='"+Coachobj.specialcare+"',";
	SQLmiddle += " photoconsent='"+Coachobj.photoconsent+"',";
	SQLmiddle += " paid='"+Coachobj.paid+"',";
	SQLmiddle += " webaccess="+Coachobj.webaccess+",";
	// Now do date fields. If null dont insert them as part of the update clause
	//  Access doesnt like setting date fields to ''
	if (! (Coachobj.dob == ""))
		SQLmiddle += " dob='"+Coachobj.dob+"',";
	else
		SQLmiddle += " dob=null,";
	if (! (Coachobj.joined == ""))
		SQLmiddle += " joined='"+Coachobj.joined+"',";
	else
		SQLmiddle += " joined=null,";
	if (! (Coachobj.left == ""))
		SQLmiddle += " dateleft='"+Coachobj.left+"',";
	else
		SQLmiddle += " dateleft=null,";
	SQLmiddle += " mailing='"+Coachobj.mailing+"', ";
	SQLmiddle += " internalleague='"+Coachobj.internalleague+"', ";
	if (Coachobj.onlinebookingid != "0")
		 SQLmiddle += " onlinebookingid="+Coachobj.onlinebookingid+", ";
	SQLmiddle += " onlinebookingpin='"+Coachobj.onlinebookingpin+"', ";
	SQLmiddle += " iscoach='"+Coachobj.iscoach+"' ";
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
function printCoach(Coachobj)
{
	// Establish local variables
	var sReport = new String("").toString();
		sReport += Coachobj.id +"<br />";
		sReport += Coachobj.grade +"<br />";
   		sReport += Coachobj.surname +"<br />";
		sReport += Coachobj.forename1 +"<br />";
		sReport += Coachobj.initials +"<br />";
		sReport += Coachobj.gender +"<br />";
		sReport += Coachobj.address1 +"<br />";
		sReport += Coachobj.address2 +"<br />";
		sReport += Coachobj.address3 +"<br />";
		sReport += Coachobj.address4 +"<br />";
		sReport += Coachobj.postcode +"<br />";
		sReport += Coachobj.homephone +"<br />";
		sReport += Coachobj.mobile +"<br />";
		sReport += Coachobj.email +"<br />";
		sReport += Coachobj.webpassword +"<br />";
		sReport += Coachobj.pool +"<br />";
		sReport += Coachobj.maxitennis +"<br />";
		sReport += Coachobj.britishtennisno +"<br />";
		sReport += Coachobj.adultcontact +"<br />";
		sReport += Coachobj.adultrelationship +"<br />";
		sReport += Coachobj.adultphone +"<br />";
		sReport += Coachobj.adultmobile +"<br />";
		sReport += Coachobj.specialcare +"<br />";
		sReport += Coachobj.photoconsent +"<br />";
		sReport += Coachobj.dob +"<br />";
		sReport += Coachobj.joined +"<br />";
		sReport += Coachobj.left +"<br />";
		sReport += Coachobj.mailing +"<br />";
		sReport += Coachobj.paid +"<br />";
		sReport += Coachobj.internalleague +"<br />";
		sReport += Coachobj.onlinebookingid +"<br />";
		sReport += Coachobj.onlinebookingpin +"<br />";
		sReport += Coachobj.iscoach +"<br />";
	return(sReport);
}
// ================================================================
function deleteCoach(Coachid)
{
	// Establish local variables
	var sCoach = new String(Coachid);
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	var mDateObj, dummy1;
	var resultObj = new Object();
	var CoachObj = new Object();

	Coachobj = getCoach(Coachid);
	SQL1 = new String("DELETE FROM Coachs WHERE uniqueref = "+sCoach).toString();
	dbconnect=Application("hamptonsportsdb"); 
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	
	resultObj.result = true;
	resultObj.errno = (0);
	resultObj.description += "Coach deleted";
	resultObj.sql = new String(SQL1).toString();
	try {
		RS = Conn.Execute(SQL1);
	}
	catch(e) {
		resultObj.result = true;
		resultObj.errno = (e.number & 0xFFFF);
		resultObj.description += e.description;
		resultObj.sql = new String(SQL1).toString();
	}
	
	// Write audit record
	if (resultObj.result) {
		// Confirm who has been deleted
		me = new String(getUserID()).toString();
		SQL2 = new String("INSERT INTO Coach_audits(Coachid,action) VALUES('"+me+"','DELETE "+CoachObj.forename1+" "+CoachObj.surname+"')").toString();
		resultObj.result = true;
		resultObj.errno = (0);
		resultObj.description += "Coach deleted";
		resultObj.sql = new String(SQL2).toString();
		try {
				RS = Conn.Execute(SQL2);
		}
		catch(e) {
				resultObj.result = true;
				resultObj.errno = (e.number & 0xFFFF);
				resultObj.description += e.description;
				resultObj.sql = new String(SQL2).toString();
		}
	
	}
	return(resultObj);

}
</script>