<script language=JScript charset="utf-8" CODEPAGE="65001" runat=Server>
//
//  getAnnualTour(tournamentyear)       returns annualtournament object
//  getTour(tournamentid)       		returns tournament object
//  setTour(tournamentobject)			returns boolean
//  printTour(tournamentobject)			returns string
//  newTour()							returns tournamentid for the new tournament
//  deleteTour(tournamewntid)			returns SQL used for deletion    
//
// ================================================================
function getAnnualTour(touryear)
{
	// Establish local variables
	var thetour = new Object();
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	var mDateObj, dummy1;
	//
	thetour.annualtournamentid = new String("").toString();
	thetour.finalsday = new String("").toString();
	thetour.touryear = new String("").toString();
	thetour.mainphotolink = new String("").toString();
	thetour.mainthanks = new String("").toString();
	thetour.finalsdaycomments = new String("").toString();
	//
	dbconnect=Application("hamptonsportsdb");
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	// Need to query particular tournament entry from tourid
	SQL1 = new String("SELECT * FROM annualtournaments WHERE touryear="+touryear).toString();
	RS = Conn.Execute(SQL1);
	if (! RS.EOF)
	{
		thetour.annualtournamentid = new String(RS("annualtournamentid")).toString();
		thetour.touryear = new String(RS("touryear")).toString();
		thetour.mainphotolink = new String(RS("mainphotolink")).toString();
		thetour.finalsday = new String(RS("finalsday")).toString();
		thetour.mainthanks = new String(RS("mainthanks")).toString();
		thetour.finalsdaycomments = new String(RS("finalsdaycomments")).toString();
		// OK now deal with formatting and validation issues

		// Firstly, null fields
		if (thetour.mainphotolink == "null" || thetour.mainphotolink == "undefined")
			thetour.mainphotolink = new String("").toString();
		if (thetour.mainthanks == "null" || thetour.mainthanks == "undefined")
			thetour.mainthanks = new String("").toString();
		if (thetour.finalsdaycomments == "null" || thetour.finalsdaycomments == "undefined")
			thetour.finalsdaycomments = new String("").toString();
	
		// Secondly, date fields

	}

	RS.Close();
	Conn.Close();
	RS = null;
	Conn = null;

	return(thetour);
}


function getTour(tourid)
{
	// Establish local variables
	var thetour = new Object();
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	var mDateObj, dummy1;
	//
	thetour.tournamentid = new String(tourid).toString();
	thetour.tourtitle = new String("").toString();
	thetour.tourstart = new String("").toString();
	thetour.tourend = new String("").toString();
	thetour.tourfinalsday = new String("").toString();
	thetour.tourwho = new String("").toString();
	thetour.tourcontact = new String("").toString();
	thetour.tourstarttime = new String("").toString();
	thetour.tourendtime = new String("").toString();
	thetour.finalsstarttime = new String("").toString();
	thetour.finalsendtime = new String("").toString();
	thetour.tourcost = new String("").toString();
	thetour.toururl = new String("tournaments.html").toString();
	thetour.tourcategory = new String("JUNIOR").toString();
	thetour.tourblurb = new String("").toString();
	//
	dbconnect=Application("hamptonsportsdb");
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	// Need to query particular tournament entry from tourid
	SQL1 = new String("SELECT * FROM tournaments WHERE tournamentid="+tourid).toString();
	Response.Write("<br />"+SQL1+"<br />");
	RS = Conn.Execute(SQL1);
	if (! RS.EOF)
	{
		thetour.tourtitle = new String(RS("tourtitle")).toString();
		thetour.tourstart = new String(RS("tourstart")).toString();
		thetour.tourend = new String(RS("tourend")).toString();
		thetour.tourfinalsday = new String(RS("tourfinalsday")).toString();
		thetour.tourwho = new String(RS("tourwho")).toString();
		thetour.tourcontact = new String(RS("tourcontact")).toString();
		thetour.tourstarttime = new String(RS("tourstarttime")).toString();
		thetour.tourendtime = new String(RS("tourendtime")).toString();
		thetour.finalsstarttime = new String(RS("finalsstarttime")).toString();
		thetour.finalsendtime = new String(RS("finalsendtime")).toString();
		thetour.tourcost = new String(RS("tourcost")).toString();
		thetour.toururl = new String(RS("toururl")).toString();
		thetour.tourcategory = new String(RS("tourcategory")).toString();
		thetour.tourblurb = new String(RS("tourblurb")).toString();
		// OK now deal with formatting and validation issues

		// Firstly, null fields
		if (thetour.tourtitle == "null" || thetour.tourtitle == "undefined")
			thetour.tourtitle = new String("").toString();
		if (thetour.tourwho == "null" || thetour.tourwho == "undefined")
			thetour.tourwho = new String("").toString();
		if (thetour.tourcontact == "null" || thetour.tourcontact == "undefined")
			thetour.tourcontact = new String("").toString();
		if (thetour.tourcost == "null" || thetour.tourcost == "undefined")
			thetour.tourcost = new String("").toString();
		if (thetour.toururl == "null" || thetour.toururl == "undefined")
			thetour.toururl = new String("").toString();
		if (thetour.tourcategory == "null" || thetour.tourcategory == "undefined")
			thetour.tourcategory = new String("").toString();
		if (thetour.tourblurb == "null" || thetour.tourblurb == "undefined")
			thetour.tourblurb = new String("").toString();

		// Secondly, date fields
		mDateObj=new Date(RS("tourstart"));
		dummy1 = mDateObj.valueOf();
		if (dummy1 == 0) // no date in database
			thetour.tourstart = "";
		else
			thetour.tourstart = ddmmyyyy(mDateObj);

		mDateObj=new Date(RS("tourend"));
		dummy1 = mDateObj.valueOf();
		if (dummy1 == 0) // no date in database
			thetour.tourend = "";
		else
			thetour.tourend = ddmmyyyy(mDateObj);

		mDateObj=new Date(RS("tourfinalsday"));
		dummy1 = mDateObj.valueOf();
		if (dummy1 == 0) // no date in database
			thetour.tourfinalsday = "";
		else
			thetour.tourfinalsday = ddmmyyyy(mDateObj);
		
		// Now time fields - must reformat as hhmm

		if (! (thetour.tourstarttime.length==8)) {
			if ( ! (thetour.tourstarttime.length==4))
				thetour.tourstarttime = new String("").toString();
		} else {
			thetour.tourstarttime = new String(Left(thetour.tourstarttime,2)+thetour.tourstarttime.substring(3,2)).toString();
		}


	}
	RS.Close();
	Conn.Close();
	RS = null;
	Conn = null;

	return(thetour);
}
// ================================================================
function setTour(tourobj, debugflag)
{
	debugflag = debugflag || false;
	// Establish local variables
	var sWinner = new String(tourobj.touryear);
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	var mDateObj, dummy1;
	//
	dbconnect=Application("hamptonsportsdb");
	Conn = Server.CreateObject("ADODB.Connection");
	RS2 = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	SQLstart = new String("UPDATE tournaments ")
	SQLend = new String(" WHERE tournamentid="+tourobj.tournamentid).toString();
	SQLmiddle = new String("SET ").toString();
	SQLmiddle += "tourtitle = '"+tourobj.tourtitle+"', ";
	SQLmiddle += "tourwho = '"+tourobj.tourwho+"', ";
	SQLmiddle += "tourcontact = '"+tourobj.tourcontact+"', ";
	SQLmiddle += "tourcost = '"+tourobj.tourcost+"', ";
	SQLmiddle += "toururl = '"+tourobj.toururl+"', ";
	SQLmiddle += "tourcategory = '"+tourobj.tourcategory+"', ";
	SQLmiddle += "tourblurb = '"+tourobj.tourblurb+"', ";

	// Now deal with the date and time fields, dates first
	if (! (tourobj.tourstart == "" || tourobj.tourstart == "null" || tourobj.tourstart == "undefined"  ))
		SQLmiddle += " tourstart='"+tourobj.tourstart+"',";
	else
		SQLmiddle += " tourstart=null,";
	if (! (tourobj.tourend == "" || tourobj.tourend == "null" || tourobj.tourend == "undefined"  ))
		SQLmiddle += " tourend='"+tourobj.tourend+"',";
	else
		SQLmiddle += " tourend=null, ";
	if (! (tourobj.tourfinalsday == "" || tourobj.tourfinalsday == "null" || tourobj.tourfinalsday == "undefined"  ))
		SQLmiddle += " tourfinalsday='"+tourobj.tourfinalsday+"',";
	else
		SQLmiddle += " tourfinalsday=null,";
	
	// Now for the time fields
	
	if (tourobj.tourstarttime == "" || tourobj.tourstarttime == "null" || tourobj.tourstarttime == "undefined"  )
		tourobj.tourstarttime = new String("").toString();
	if (tourobj.tourendtime == "" || tourobj.tourendtime == "null" || tourobj.tourendtime == "undefined"  )
		tourobj.tourendtime = new String("").toString();
	if (tourobj.finalsstarttime == "" || tourobj.finalsstarttime == "null" || tourobj.finalsstarttime == "undefined"  )
		tourobj.finalsstarttime = new String("").toString();
	if (tourobj.finalsendtime == "" || tourobj.finalsendtime == "null" || tourobj.finalsendtime == "undefined"  )
		tourobj.finalsendtime = new String("").toString();

	if (! (tourobj.tourstarttime.length==8)) {
		if ( ! (tourobj.tourstarttime.length==4))
			SQLmiddle += " tourstarttime=null, "
		else 
			SQLmiddle += " tourstarttime = '"+tourobj.tourstarttime+"',";
	}
	if (! (tourobj.tourendtime.length==8)) {
		if ( ! (tourobj.tourendtime.length==4))
			SQLmiddle += " tourendtime=null, "
		else 
			SQLmiddle += " tourendtime = '"+tourobj.tourendtime+"',";
	}
	if (! (tourobj.finalsstarttime.length==8)) {
		if ( ! (tourobj.finalsstarttime.length==4))
			SQLmiddle += " finalsstarttime=null, "
		else 
			SQLmiddle += " finalsstarttime = '"+tourobj.finalsstarttime+"', ";
	}
	if (! (tourobj.finalsendtime.length==8)) {
		if ( ! (tourobj.finalsendtime.length==4))
			SQLmiddle += " finalsendtime=null "
		else 
			SQLmiddle += " finalsendtime = '"+tourobj.finalsendtime+"' ";
	}


	SQL1 = new String(SQLstart+SQLmiddle+SQLend).toString();;
	// if ( ! debugflag)
	if (! debugflag) {
		RS2 = Conn.Execute(SQL1);
	}
	// RS2.Close();
	Conn.Close();
	RS2 = null;
	Conn = null;

	return(SQL1);
}
// ================================================================
function printTour(tourObj)
{
	// Establish local variables
	var sReport = new String("").toString();
		sReport += "Tournament ID: "+tourObj.tournamentid +"<br />";
		sReport += "Tournament title: "+tourObj.tourtitle +"<br />";
   		sReport += "Tournament start date: "+tourObj.tourstart +"<br />";
		sReport += "Tournament end date: "+tourObj.tourend +"<br />";
		sReport += "Tournament finals day: "+tourObj.tourfinalsday +"<br />";
		sReport += "Who tournament is for: "+tourObj.tourwho +"<br />";
		sReport += "Who to contact about this tournament: "+tourObj.tourcontact +"<br />";
		sReport += "Start time on start day: "+tourObj.tourstarttime +"<br />";
		sReport += "End time on end date: "+tourObj.tourendtime +"<br />";
		sReport += "Finals day start time: "+tourObj.finalsstarttime +"<br />";
		sReport += "Finals day end time: "+tourObj.finalsendtime +"<br />";
		sReport += "Cost to enter: "+tourObj.tourcost +"<br />";
		sReport += "Web URL: "+tourObj.toururl +"<br />";
		sReport += "Category of tournament: "+tourObj.tourcategory +"<br />";
		sReport += "Additional information link: "+tourObj.tourblurb +"<br />";

	return(sReport);
}

// ================================================================
function newTour(debugflag)
{
	// Insert new skeleton tournament row into table and then 
	// return the tournament ID so the tournament can be fetched 
	// afterwards if required
	debugflag = debugflag || false;
	// Establish local variables
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	var SQLstart, SQLmiddle, SQLend;
	var tourid;
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

	SQLstart = new String("INSERT INTO tournaments ([tourtitle], [tourstart],[tourend],[tourcategory]) ").toString();
	SQLend = new String("").toString();
	SQLmiddle = new String("VALUES ('NEW TOURNAMENT', '01/01/2099', '01/01/2199', '"+text+"')").toString();

	SQL1 = new String(SQLstart+SQLmiddle+SQLend).toString();

	// Insert dummy row 

	RS = Conn.Execute(SQL1);

	// Now get the eventid back
	tourid = -1;
	SQL2 = new String("SELECT tournamentid FROM tournaments WHERE tourcategory = '"+text+"'").toString();

	RS = Conn.Execute(SQL2);
	if (! RS.EOF) {
		tourid = new Number(RS("tournamentid"));
	}
	RS.Close();
	Conn.Close();
	RS = null;
	Conn = null;

	return(tourid);
}

// ================================================================
function deleteTour(tournamentid, debugflag)
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
	SQL1 = new String("DELETE FROM tournaments WHERE tournamentid = "+tournamentid);

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