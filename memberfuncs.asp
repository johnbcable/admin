<script language=JScript runat=Server>
//
//  getMember(memberid)           returns member Object
//  setMember(memberObject)		  returns result Object
//  printMember(memberObject)	  returns string
//  deleteMember(memberid)        returns result Object 
//
// ================================================================
function getMember(memberid)
{
	// Establish local variable
	var themember = new Object();
	var sMember = new String(memberid);
	var RS, RS2, Conn, SQL1, SQL2, uniqref;
	var mDateObj, dummy1;
	var mDob;
	//
	themember.memberid = new String(sMember).toString();
	//
	var dbconnect=Application("hamptonsportsdb"); 
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	// 
	SQL1 = new String("SELECT * from members where uniqueref = " + sMember);
	RS = Conn.Execute(SQL1);
	// Retrieve database values for member
	mDob = new String("").toString();
	themember.uniqueref = sMember;
	while (! RS.EOF)
	{
		themember.id = Trim(new String(RS("memberid")).toString());
		themember.grade = Trim(new String(RS("membergrade")).toString());
 		themember.surname= Trim(new String(RS("surname")).toString());
		themember.forename1= Trim(new String(RS("forename1")).toString());
		// Set initials to empty string as we dont use these any more
		themember.initials = new String("").toString();
		themember.gender = Trim(new String(RS("gender")).toString());
		themember.address1= Trim(new String(RS("address1")).toString());
		themember.address2= Trim(new String(RS("address2")).toString());
		themember.address3= Trim(new String(RS("address3")).toString());
		themember.address4= Trim(new String(RS("address4")).toString());
		themember.postcode= Trim(new String(RS("postcode")).toString());
		themember.homephone= Trim(new String(RS("homephone")).toString());
		themember.mobile= Trim(new String(RS("mobilephone")).toString());
		themember.email= Trim(new String(RS("email")).toString());
		themember.webpassword = Trim(new String(RS("webpassword")).toString());
		themember.pool = Trim(new String(RS("pool")).toString());
		themember.maxitennis = Trim(new String(RS("maxiteam")).toString());
		themember.britishtennisno = Trim(new String(RS("britishtennisno")).toString());
		themember.adultcontact = Trim(new String(RS("adultcontact")).toString());
		themember.adultrelationship = Trim(new String(RS("adultrelationship")).toString());
		themember.adultphone = Trim(new String(RS("adultphone")).toString());
		themember.adultmobile = Trim(new String(RS("adultmobile")).toString());
		themember.specialcare = Trim(new String(RS("specialcare")).toString());
		themember.photoconsent = Trim(new String(RS("photoconsent")).toString());
		themember.paid = StringFromDB(new String(RS("paid")));
		themember.mailing = Trim(new String(RS("mailing")).toString());
		themember.internalleague = Trim(new String(RS("internalleague")).toString());
		themember.onlinebookingid = Trim(new String(RS("onlinebookingid")).toString());
		themember.onlinebookingpin = Trim(new String(RS("onlinebookingpin")).toString());
		themember.iscoach = Trim(new String(RS("iscoach")).toString());
		// 	mWimbledonDraw = Trim(new String(RS("wimbledondraw")));
		dummy=new Number(RS("webaccess")).valueOf();
		themember.webaccess=dummy.valueOf();
		// now the date bits of the member record
		// First, date of birth
		mDateObj=new Date(RS("dob"));
		dummy1 = mDateObj.valueOf();
		if (dummy1 == 0) // no date in database
			themember.dob = "";
		else
			themember.dob = ddmmyyyy(mDateObj);
		// Second, joining date
		mDateObj=new Date(RS("joined"));
		dummy1 = mDateObj.valueOf();
		if (dummy1 == 0) // no date in database
			themember.joined = "";
		else
			themember.joined = ddmmyyyy(mDateObj);
		// Third, date of leaving the club
		mDateObj=new Date(RS("dateleft"));
		dummy1 = mDateObj.valueOf();
		if (dummy1 == 0) // no date in database
			themember.left = "";
		else
			themember.left = ddmmyyyy(mDateObj);
		// End of date bits
		RS.MoveNext();
	}
	if (themember.grade=="null" || themember.grade=="undefined")
		themember.grade="";
	if (themember.surname=="null" || themember.surname=="undefined")
		themember.surname="";
	if (themember.forename1=="null" || themember.forename1=="undefined")
		themember.forename1="";
	if (themember.initials=="null" || themember.initials=="undef")
		themember.initials="";
	if (themember.address1=="null" || themember.address1=="undefined")
		themember.address1="";
	if (themember.address2=="null" || themember.address2=="undefined")
		themember.address2="";
	if (themember.address3=="null" || themember.address3=="undefined")
		themember.address3="";
	if (themember.address4=="null" || themember.address4=="undefined")
		themember.address4="";
	if (themember.postcode=="null" || themember.postcode=="undefined")
		themember.postcode="";
	if (themember.homephone=="null" || themember.homephone=="undefined")
		themember.homephone="";
	if (themember.mobile=="null" || themember.mobile=="undefined")
		themember.mobile="";
	if (themember.email=="null" || themember.email=="undefined")
		themember.email="";
	if (themember.webpassword=="null" || themember.webpassword=="undefined")
		themember.webpassword="tennis";
	if (themember.pool == "null" || themember.pool == "undefined")
		themember.pool = "";
	if (themember.maxitennis == "null" || themember.maxitennis == "undefined")
		themember.maxitennis = "";
	if (themember.britishtennisno == "null" || themember.britishtennisno=="undefined")
		themember.britishtennisno = "";
	if (themember.adultcontact == "null" || themember.adultcontact=="undefined")
		themember.adultcontact = "";
	if (themember.adultrelationship == "null" || themember.adultrelationship=="undefined")
		themember.adultrelationship = "";
	if (themember.adultphone == "null" || themember.adultphone == "undefined" )
		themember.adultphone = "";
	if (themember.adultmobile == "null" || themember.adultmobile == "undefined" )
		themember.adultmobile = "";
	if (themember.specialcare == "null" || themember.specialcare == "undefined" )
		themember.specialcare = "";
	if (themember.photoconsent == "null" || themember.photoconsent == "undefined" || themember.photoconsent == "")
		themember.photoconsent = new String("N").toString();
	if (themember.paid == "null" || themember.paid == "undefined" || themember.paid == "")
		themember.paid = new String("N").toString();
	if (themember.webaccess < 20)
		themember.webaccess = 20;
	if (themember.mailing=="null" || themember.mailing=="" || themember.mailing=="undefined")
			themember.mailing = "N";	
	if (themember.internalleague=="null" || themember.internalleague=="" || themember.internalleague=="undefined")
			themember.internalleague = "";	
	if (themember.onlinebookingid=="null" || themember.onlinebookingid=="")
	{
	 	  dummy1 = new Number(memberid)+5000;
			themember.onlinebookingid = new String(dummy1).toString();	
	}
	if (themember.onlinebookingpin=="null" || themember.onlinebookingpin=="" || themember.onlinebookingpin=="undefined")
			themember.onlinebookingpin = new String("").toString();	
	if (themember.gender=="null" || themember.gender=="" || themember.gender=="undefined")
			themember.gender = "";	
	if (themember.iscoach == "null" || themember.iscoach == "undefined" || themember.iscoach == "")
		themember.iscoach = new String("N").toString();
	
	return(themember);
}
// ================================================================
function setMember(memberobj)
{
	// Establish local variables
	var sMember = new String(memberobj.memberid);
	var RS, RS2, Conn, SQL1, SQL2, uniqref;
	var mDateObj, dummy1;
	var resultObj = new Object();
	var e;
	//
	resultObj.result = true;
	resultObj.errno = 0;
	resultObj.description = new String("").toString();
	//
	if (memberobj.grade=="null" || memberobj.grade =="undefined")
		memberobj.grade="";
	if (memberobj.surname=="null" || memberobj.surname =="undefined")
		memberobj.surname="";
	if (memberobj.forename1=="null" || memberobj.forename1 =="undefined")
		memberobj.forename1="";
	// Set initials to empty string as we dont use these any more
	memberobj.initials="";
	if (memberobj.address1=="null" || memberobj.address1 =="undefined")
		memberobj.address1="";
	if (memberobj.address2=="null" || memberobj.address2 =="undefined")
		memberobj.address2="";
	if (memberobj.address3=="null" || memberobj.address3 =="undefined")
		memberobj.address3="";
	if (memberobj.address4=="null" || memberobj.address4 =="undefined")
		memberobj.address4="";
	if (memberobj.postcode=="null" || memberobj.postcode =="undefined")
		memberobj.postcode="";
	if (memberobj.homephone=="null" || memberobj.homephone =="undefined")
		memberobj.homephone="";
	if (memberobj.mobile=="null" || memberobj.mobile =="undefined")
		memberobj.mobile="";
	if (memberobj.email=="null" || memberobj.email =="undefined")
		memberobj.email="";
	if (memberobj.webpassword=="null" || memberobj.webpassword =="undefined")
		memberobj.webpassword="tennis";
	if (memberobj.pool == "null" || memberobj.pool =="undefined")
		memberobj.pool = "";
	if (memberobj.maxitennis == "null" || memberobj.maxitennis =="undefined")
		memberobj.maxitennis = "";
	if (memberobj.britishtennisno == "null" || memberobj.britishtennisno =="undefined")
		memberobj.britishtennisno = "";
	if (memberobj.adultcontact == "null" || memberobj.adultcontact =="undefined")
		memberobj.adultcontact = "";
	if (memberobj.adultrelationship == "null" || memberobj.adultrelationship =="undefined")
		memberobj.adultrelationship = "";
	if (memberobj.adultphone == "null" || memberobj.adultphone =="undefined")
		memberobj.adultphone = "";
	if (memberobj.adultmobile == "null" || memberobj.adultmobile =="undefined")
		memberobj.adultmobile = "";
	if (memberobj.specialcare == "null" || memberobj.specialcare =="undefined")
		memberobj.specialcare = "";
	if (memberobj.photoconsent == "null" || memberobj.photoconsent == "undefined" || memberobj.photoconsent == "")
		memberobj.photoconsent = "N";
	if (memberobj.paid == "null" || memberobj.paid == "undefined" || memberobj.paid == "")
		memberobj.paid = "N";
	if (memberobj.internalleague=="null" || memberobj.internalleague=="" || memberobj.internalleague=="undefined")
			memberobj.internalleague = "";	
	if (memberobj.onlinebookingid=="null" || memberobj.onlinebookingid=="" || memberobj.onlinebookingid=="undefined")
	{
		if (memberobj.grade == "Adult" || memberobj.grade == "18-25s" || memberobj.grade == "Junior")  
			 memberobj.onlinebookingid = new String(new Number(memberobj.uniqueref)+5000).toString();
		else
			 memberobj.onlinebookingid = new String("0").toString();
	}
	if (memberobj.onlinebookingpin=="null" || memberobj.onlinebookingpin=="" || memberobj.onlinebookingpin=="undefined")
			memberobj.onlinebookingpin = new String("").toString();
	if (memberobj.webaccess < 20)
		memberobj.webaccess = 20;
	if (memberobj.mailing=="null" || memberobj.mailing =="undefined" || memberobj.mailing=="")
			memberobj.mailing = "N";	
	if (memberobj.dob=="null" || memberobj.dob =="undefined")
			memberobj.dob = "";	
	if (memberobj.left=="null" || memberobj.left =="undefined")
			memberobj.left = "";	
	if (memberobj.joined=="null" || memberobj.joined =="undefined")
			memberobj.joined = "";	
	if (memberobj.gender=="null" || memberobj.gender =="undefined" || memberobj.gender=="")
			memberobj.gender = "";	
	if (memberobj.iscoach == "null" || memberobj.iscoach == "undefined" || memberobj.iscoach == "")
		memberobj.iscoach = "N";
	
	var dbconnect=Application("hamptonsportsdb"); 
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	SQLstart = new String("UPDATE members ")
	SQLend = new String(" WHERE uniqueref="+memberobj.uniqueref).toString();
	SQLmiddle = new String("SET ").toString();
	SQLmiddle += " memberid='"+memberobj.id+"',";
	SQLmiddle += " membergrade='"+memberobj.grade+"',";
	SQLmiddle += " surname='"+memberobj.surname+"',";
	SQLmiddle += " forename1='"+memberobj.forename1+"',";
	SQLmiddle += " initials='"+memberobj.initials+"',";
	SQLmiddle += " gender='"+memberobj.gender+"', ";
	SQLmiddle += " address1='"+memberobj.address1+"',";
	SQLmiddle += " address2='"+memberobj.address2+"',";
	SQLmiddle += " address3='"+memberobj.address3+"',";
	SQLmiddle += " address4='"+memberobj.address4+"',";
	SQLmiddle += " postcode='"+memberobj.postcode+"',";
	SQLmiddle += " homephone='"+memberobj.homephone+"',";
	SQLmiddle += " mobilephone='"+memberobj.mobile+"',";
	SQLmiddle += " email='"+memberobj.email+"',";
	SQLmiddle += " webpassword='"+memberobj.webpassword+"',";
	SQLmiddle += " pool='"+memberobj.pool+"',";
	SQLmiddle += " maxiteam='"+memberobj.maxitennis+"',";
	SQLmiddle += " britishtennisno='"+memberobj.britishtennisno+"',";
	SQLmiddle += " adultcontact='"+memberobj.adultcontact+"',";
	SQLmiddle += " adultrelationship='"+memberobj.adultrelationship+"',";
	SQLmiddle += " adultphone='"+memberobj.adultphone+"',";
	SQLmiddle += " adultmobile='"+memberobj.adultmobile+"',";
	SQLmiddle += " specialcare='"+memberobj.specialcare+"',";
	SQLmiddle += " photoconsent='"+memberobj.photoconsent+"',";
	SQLmiddle += " paid='"+memberobj.paid+"',";
	SQLmiddle += " webaccess="+memberobj.webaccess+",";
	// Now do date fields. If null dont insert them as part of the update clause
	//  Access doesnt like setting date fields to ''
	if (! (memberobj.dob == ""))
		SQLmiddle += " dob='"+memberobj.dob+"',";
	else
		SQLmiddle += " dob=null,";
	if (! (memberobj.joined == ""))
		SQLmiddle += " joined='"+memberobj.joined+"',";
	else
		SQLmiddle += " joined=null,";
	if (! (memberobj.left == ""))
		SQLmiddle += " dateleft='"+memberobj.left+"',";
	else
		SQLmiddle += " dateleft=null,";
	SQLmiddle += " mailing='"+memberobj.mailing+"', ";
	SQLmiddle += " internalleague='"+memberobj.internalleague+"', ";
	if (memberobj.onlinebookingid != "0")
		 SQLmiddle += " onlinebookingid="+memberobj.onlinebookingid+", ";
	SQLmiddle += " onlinebookingpin='"+memberobj.onlinebookingpin+"', ";
	SQLmiddle += " iscoach='"+memberobj.iscoach+"' ";
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
function printMember(memberobj)
{
	// Establish local variables
	var sReport = new String("").toString();
		sReport += memberobj.id +"<br />";
		sReport += memberobj.grade +"<br />";
   		sReport += memberobj.surname +"<br />";
		sReport += memberobj.forename1 +"<br />";
		sReport += memberobj.initials +"<br />";
		sReport += memberobj.gender +"<br />";
		sReport += memberobj.address1 +"<br />";
		sReport += memberobj.address2 +"<br />";
		sReport += memberobj.address3 +"<br />";
		sReport += memberobj.address4 +"<br />";
		sReport += memberobj.postcode +"<br />";
		sReport += memberobj.homephone +"<br />";
		sReport += memberobj.mobile +"<br />";
		sReport += memberobj.email +"<br />";
		sReport += memberobj.webpassword +"<br />";
		sReport += memberobj.pool +"<br />";
		sReport += memberobj.maxitennis +"<br />";
		sReport += memberobj.britishtennisno +"<br />";
		sReport += memberobj.adultcontact +"<br />";
		sReport += memberobj.adultrelationship +"<br />";
		sReport += memberobj.adultphone +"<br />";
		sReport += memberobj.adultmobile +"<br />";
		sReport += memberobj.specialcare +"<br />";
		sReport += memberobj.photoconsent +"<br />";
		sReport += memberobj.dob +"<br />";
		sReport += memberobj.joined +"<br />";
		sReport += memberobj.left +"<br />";
		sReport += memberobj.mailing +"<br />";
		sReport += memberobj.paid +"<br />";
		sReport += memberobj.internalleague +"<br />";
		sReport += memberobj.onlinebookingid +"<br />";
		sReport += memberobj.onlinebookingpin +"<br />";
		sReport += memberobj.iscoach +"<br />";
	return(sReport);
}
// ================================================================
function deleteMember(memberid)
{
	// Establish local variables
	var sMember = new String(memberid);
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	var mDateObj, dummy1;
	var resultObj = new Object();
	var memberObj = new Object();

	memberobj = getMember(memberid);
	SQL1 = new String("DELETE FROM members WHERE uniqueref = "+sMember).toString();
	dbconnect=Application("hamptonsportsdb"); 
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	
	resultObj.result = true;
	resultObj.errno = (0);
	resultObj.description += "Member deleted";
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
		SQL2 = new String("INSERT INTO member_audits(memberid,action) VALUES('"+me+"','DELETE "+memberObj.forename1+" "+memberObj.surname+"')").toString();
		resultObj.result = true;
		resultObj.errno = (0);
		resultObj.description += "Member deleted";
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