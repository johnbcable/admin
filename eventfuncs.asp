<script language="JScript" CODEPAGE="65001" runat=Server>
//
//  getFutureEvents( limit)               	returns string
//  getEventByID(eventid)               	returns Event object
//  getEventForTournament(tournamentid)		returns Event object
//  printEvent()							returns string
//  newEvent()								returns Event ID
//  setEvent(eventObj)						returns SQL used for update
//  deleteEvent(eventid)					returns SQL used for deletion
//
var globaldebugflag = true;

function getFutureEvents(howmanylimit)
{
	// Establish local variables
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	var debugging = false;
	var editline, Eventsubject;
	var Eventlist = new String("").toString();
	var Eventknt = 0;
	var eventclass, evreport;
	var totaltoget = new Number(howmanylimit);
	totaltoget = totaltoget.valueOf();
	// Set up database connections
	var dbconnect=Application("hamptonsportsdb"); 
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	// Need to query last 5 Events sent and return as string
	// title<hr><br>event1<br>event2<br> .... eventN<hr>
	SQL1 = new String("SELECT * from futureevents").toString();
	Eventlist += '<div id="futureevents"><ol>';
	RS = Conn.Execute(SQL1);
	Eventknt = 0;
	while (! RS.EOF && Eventknt < totaltoget)
	{
		evreport = new String(RS("eventreport")).toString();
		if (evreport == "" || evreport == "null" || evreport == "undefined")
			evreport = new String("NONE").toString();
		eventclass = new String(RS("eventtype")).toString();
		eventclass = eventclass.toLowerCase();
		if (evreport == "NONE")
		{
			Eventlist += '<li class="'+eventclass+'">'+RS("dateofevent")+' - '+RS("eventnote")+'</li>';
		}
		else
		{
			Eventlist += '<li class="'+eventclass+'"><a href="'+evreport+'">'+RS("dateofevent")+' - '+RS("eventnote")+'</a></li>';
		}

		Eventknt++;
		RS.MoveNext();
	}
	Eventlist += '</ol><p><a href="eventlist.asp">more events</a></p>';
	Eventlist += '</div>';
	RS.Close();
	return (Eventlist);
}
//=====================================================================
function getEventByID(eventid)
{
	// Establish local variables
	var eventObj = new Object();
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref, mDateObj, dummy1;
	// Set up database connections
	var dbconnect=Application("hamptonsportsdb"); 
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	// Need to query particular Event database entry from fileref
	SQL1 = new String("SELECT * from events where eventid="+eventid).toString();
	RS = Conn.Execute(SQL1);
	while (! RS.EOF)
	{
		eventObj.eventdate = new String(RS("eventdate")).toString();
		eventObj.eventtime = new String(RS("eventtime")).toString();
		eventObj.eventyear = new String(RS("eventyear")).toString();
		eventObj.eventtype = new String(RS("eventtype")).toString();
		eventObj.eventnote = new String(RS("eventnote")).toString();
		eventObj.eventid = new Number(RS("eventid")).valueOf();
		eventObj.eventreport = new String(RS("eventreport")).toString();
		eventObj.enddate = new String(RS("enddate")).toString();
		eventObj.endtime = new String(RS("endtime")).toString();
		eventObj.fixturelink = new String(RS("fixturelink")).toString();
		eventObj.tourlink = new String(RS("tourlink")).toString();
		eventObj.holidaylink = new String(RS("holidaylink")).toString();
		eventObj.advert = new String(RS("advert")).toString();
		RS.MoveNext();
	}
	RS.Close();
	Conn.Close();
	RS = null;
	Conn = null;

	//
	if (eventObj.eventdate=="null" || eventObj.eventdate=="undefined")
		eventObj.eventdate="";
	if (eventObj.eventtime=="null" || eventObj.eventtime=="undefined")
		eventObj.eventtime="";
	if (eventObj.eventtype=="null" || eventObj.eventtype=="undefined")
		eventObj.eventtype="";
	if (eventObj.eventnote=="null" || eventObj.eventnote=="undefined")
		eventObj.eventnote="";
	if (eventObj.eventreport=="null" || eventObj.eventreport=="undefined")
		eventObj.eventreport="";
	// If no end date, default to event (start) date
	if (eventObj.enddate=="null" || eventObj.enddate=="undefined")
		eventObj.enddate=eventObj.eventdate;
	// If no end time, default to start time
	if (eventObj.endtime=="null" || eventObj.endtime=="undefined")
		eventObj.eventdtime="";
	if (eventObj.fixturelink=="null" || eventObj.fixturelink=="undefined")
		eventObj.fixturelink="";
	if (eventObj.tourlink=="null" || eventObj.tourlink=="undefined")
		eventObj.tourlink="";
	if (eventObj.holidaylink=="null" || eventObj.holidaylink=="undefined")
		eventObj.holidaylink="";
	if (eventObj.advert=="null" || eventObj.advert=="undefined")
		eventObj.advert="";

	if (eventObj.eventyear=="null" || eventObj.eventyear=="undefined")
		eventObj.eventyear=new String("").toString();

	// Make sure date fields are reformatted as dd/mm/yyyy
	mDateObj=new Date(eventObj.eventdate);
	dummy1 = mDateObj.valueOf();
	if (dummy1 == 0) // no date in database
		eventObj.eventdate = "";
	else
		eventObj.eventdate = ddmmyyyy(mDateObj);

	mDateObj=new Date(eventObj.enddate);
	dummy1 = mDateObj.valueOf();
	if (dummy1 == 0) // no date in database
		eventObj.enddate = "";
	else
		eventObj.enddate = ddmmyyyy(mDateObj);

	return (eventObj);
}
//=====================================================================
function getEventForTournament(tournamentid)
{
	// Establish local variables
	var Eventobj = new Object();
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	// Set up database connections
	var dbconnect=Application("hamptonsportsdb"); 
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	// Need to query particular Event database entry from fileref
	SQL1 = new String("SELECT * from events where linktable='tournaments' AND linkid = "+m_id).toString();
	RS = Conn.Execute(SQL1);
	if (! RS.EOF)
	{
		Eventobj.eventdate = new String(RS("eventdate")).toString();
		Eventobj.eventtime = new String(RS("eventtime")).toString();
		Eventobj.eventyear = new Number(RS("eventyear")).valueOf();
		Eventobj.eventtype = new String(RS("eventtype")).toString();
		Eventobj.eventnote = new String(RS("eventnote")).toString();
		Eventobj.eventid = new Number(RS("eventid")).valueOf();
		Eventobj.eventreport = new String(RS("eventreport")).toString();
		EventObj.enddate = new String(RS("enddate")).toString();
		EventObj.endtime = new String(RS("endtime")).toString();
		EventObj.fixturelink = new String(RS("fixturelink")).toString();
		EventObj.tourlink = new String(RS("tourlink")).toString();
		EventObj.holidaylink = new String(RS("holidaylink")).toString();
		EventObj.advert = new String(RS("advert")).toString();
		RS.MoveNext();
	}
	else
	{
		Eventobj = null;
	}
	RS.Close();
	return (Eventobj);
}

// ================================================================
function printEvent(eventObj)
{
	// Establish local variables
	var sReport = new String("").toString();

	if (eventObj)
	{
		sReport += "Event ID: "+eventObj.eventid +"<br />";
		sReport += "Title for event: "+eventObj.eventnote +"<br />";
   		sReport += "Date of event: "+eventObj.eventdate +"<br />";
		sReport += "Start time for event: "+eventObj.eventtime +"<br />";
		sReport += "Year of event: "+eventObj.eventyear +"<br />";
		sReport += "Type of event: "+eventObj.eventtype +"<br />";
		sReport += "Further info about event: "+eventObj.eventreport +"<br />";
		sReport += "End date for event: "+eventObj.enddate +"<br />";
		sReport += "End tiem for event: "+eventObj.endtime +"<br />";
		sReport += "Link to fixtures (if present): "+eventObj.fixturelink+"<br />";
		sReport += "Link to tournaments (if present): "+eventObj.tourlink +"<br />";
		sReport += "Link to holiday clubs (if present): "+eventObj.holidaylink +"<br />";
		sReport += "Link to advert for this event (if present): "+eventObj.advert +"<br />";

	}	
	else
	{
		sReport += "There is no data for this event<br />";
	}

	return(sReport);
}

// ================================================================
function newEvent(debugflag)
{
	// Insert new skeleton tournament row into table and then 
	// return the event ID so the evnt can be fetched afterwards
	debugflag = debugflag || false;
	// Establish local variables
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	var SQLstart, SQLmiddle, SQLend;
	var resultObj = new Object();
	var eventid;
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

	SQLstart = new String("INSERT INTO events ([eventdate],[eventtime],[eventtype]) ").toString();
	SQLend = new String("").toString();
	SQLmiddle = new String("VALUES ('01/01/2099', '10:00:00', '"+text+"')").toString();

	SQL1 = new String(SQLstart+SQLmiddle+SQLend).toString();

	// Insert dummy row 

	RS = Conn.Execute(SQL1);

	// Now get the eventid back
	eventid = -1;
	SQL2 = new String("SELECT eventid FROM events WHERE eventtype = '"+text+"'").toString();

	RS = Conn.Execute(SQL2);
	if (! RS.EOF) {
		eventid = new Number(RS("eventid"));
		// eventid = eventid.value();
	}
	RS.Close();
	Conn.Close();
	RS = null;
	Conn = null;

	return(eventid);
}

// ================================================================
function setEvent(eventobj, debugflag)
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
	SQLstart = new String("UPDATE events ")
	SQLend = new String(" WHERE eventid="+eventobj.eventid).toString();
	SQLmiddle = new String("SET ").toString();
	SQLmiddle += "eventtype = '"+eventobj.eventtype+"', ";
	if (eventobj.eventnote == "" | eventobj.eventnote == "null" || eventobj.eventnote == "undefined") {
		SQLmiddle += "eventnote = null, ";
	} else {
		SQLmiddle += "eventnote = '"+eventobj.eventnote+"', ";
	}
	if (eventobj.eventreport == "" | eventobj.eventreport == "null" || eventobj.eventreport == "undefined") {
		SQLmiddle += "eventreport = null, ";
	} else {
		SQLmiddle += "eventreport = '"+eventobj.eventreport+"', ";
	}


	// Now deal with the date and time fields, dates first
	if (! (eventobj.eventdate == "" || eventobj.eventdate == "null" || eventobj.eventdate == "undefined"  ))
		SQLmiddle += " eventdate='"+eventobj.eventdate+"', ";
	else
		SQLmiddle += " eventdate=null, ";

	// Make sure and uypdate the event year to match the event date
	mDateObj = new Date(eventobj.eventdate);
	dummy1 = mDateObj.getFullYear();
	eventobj.eventyear = new Number(dummy1).valueOf();
	SQLmiddle += "eventyear = "+eventobj.eventyear+", ";

	// Now for the time fields
	
	if (! (eventobj.eventtime.length==8)) {
		if ( ! (eventobj.eventtime.length==4))
			SQLmiddle += " eventtime=null"
		else 
			SQLmiddle += " eventtime = '"+eventobj.eventtime+"'";
	}

	SQL1 = new String(SQLstart+SQLmiddle+SQLend).toString();;
	if ( ! debugflag) {
		RS = Conn.Execute(SQL1);

	}
		
	// RS.Close();
	Conn.Close();
	RS = null;
	Conn = null;

	return(SQL1);
}

// ================================================================
function deleteEvent(eventid, debugflag)
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
	SQL1 = new String("DELETE FROM events WHERE eventid = "+eventid);

	if ( ! debugflag)
		RS = Conn.Execute(SQL1);

	// RS.Close();
	Conn.Close();
	RS = null;
	Conn = null;
	
	return(SQL1);
}

// ================================================================
</script>
