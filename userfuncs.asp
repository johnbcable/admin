<script language="JScript" CODEPAGE="65001" runat=Server>
// List of function prototypes
// -
// getUserID()
// getUserCode()
// getOnlineID()
// getUserName();
// getUserAccessLevel()
// setUser(theid,thename,thelevel)
// getSessionMode()
// current_debug_status()
//
//==================================================
function getUserID() {
	// Now need to do this via user_sessions table and not via cookies
	// DB connection stuff
	var zz, result;
	var myusername, myaddress;
	var thedate, thetime;
	var dbconnect=Application("hamptonsportsdb"); 
	var connobj, RS, SQL2;
	myusername = new String(Request.ServerVariables("REMOTE_USER")).toString();
	myaddress = new String(Request.ServerVariables("REMOTE_ADDR")).toString();
	if (myusername == "" || myusername=="null" || myusername == "undefined")
		myusername = new String(myaddress).toString();
	connobj=Server.CreateObject("ADODB.Connection");
	RS=Server.CreateObject("ADODB.Recordset");
	connobj.Open(dbconnect);
	// Get info from sign-on record
	SQL2 = new String("SELECT memberid FROM User_Sessions where remoteaddress = '"+myaddress+"'").toString();
	v_memberid = "NONE";
	RS=connobj.Execute(SQL2);
	if (! RS.EOF)
		v_memberid = new String(RS("memberid")).toString();
	RS.Close();
	RS=null;
	connobj.Close();
	connobj=null;
	return(v_memberid);
}
//==================================================
function getUserCode() {
	var zz, result;
	var dbconnect=Application("hamptonsportsdb"); 
	var connobj, RS, SQL1;
	zz=getUserID();
	result=new Number(0);
	SQL1="select uniqueref from members where memberid = '"+zz+"'";
	connobj=Server.CreateObject("ADODB.Connection");
	RS=Server.CreateObject("ADODB.Recordset");
	connobj.Open(dbconnect);
	RS=connobj.Execute(SQL1);
	while(! RS.EOF)
	{
		result=new Number(RS("uniqueref"));
		RS.MoveNext();
	}
	RS=null;
	connobj.Close();
	connobj=null;
	return(result.valueOf());
}
//==================================================
function getOnlineID() {
	var zz, result;
	var dbconnect=Application("hamptonsportsdb"); 
	var connobj, RS, SQL1;
	zz=getUserCode();
	result=new Number(0);
	SQL1="select onlinebookingid from members where uniqueref = "+zz;
	connobj=Server.CreateObject("ADODB.Connection");
	RS=Server.CreateObject("ADODB.Recordset");
	connobj.Open(dbconnect);
	RS=connobj.Execute(SQL1);
	while(! RS.EOF)
	{
		result=new Number(RS("onlinebookingid"));
		RS.MoveNext();
	}
	RS=null;
	connobj.Close();
	connobj=null;
	return(result.valueOf());
}
//==================================================
function getUserName() {
	var zz, result;
	var dbconnect=Application("hamptonsportsdb"); 
	var connobj, RS, SQL1;
	zz=getUserID();
	result=new String("").toString();
	SQL1="select membername from User_Sessions where memberid = '"+zz+"'";
	connobj=Server.CreateObject("ADODB.Connection");
	RS=Server.CreateObject("ADODB.Recordset");
	connobj.Open(dbconnect);
	RS=connobj.Execute(SQL1);
	if(! RS.EOF)
		result=new String(RS("membername")).toString();
	RS.Close();
	RS=null;
	connobj.Close();
	connobj=null;
	return(result.valueOf());
}
//==================================================
function getUserAccessLevel() {
	var zz, result;
	var dbconnect=Application("hamptonsportsdb"); 
	var connobj, RS, SQL1;
	zz=getUserCode();
	result=new Number(0);
	SQL1=new String("select webaccess from members where uniqueref = "+zz).toString();
	connobj=Server.CreateObject("ADODB.Connection");
	RS=Server.CreateObject("ADODB.Recordset");
	connobj.Open(dbconnect);
	RS=connobj.Execute(SQL1);
	while(! RS.EOF)
	{
		result=new Number(RS("webaccess"));
		RS.MoveNext();
	}
	RS=null;
	connobj.Close();
	connobj=null;
	return(result.valueOf());
}
//==================================================
function setUser(theref,theid,thename,thelevel) {
	// Now need to do this via user_sessions table and not via cookies
	// DB connection stuff
	var zz, result;
	var myusername, myaddress;
	var thedate, thetime;
	var dbconnect=Application("hamptonsportsdb"); 
	var connobj, RS, SQL1;
	myusername = new String(Request.ServerVariables("REMOTE_USER")).toString();
	myaddress = new String(Request.ServerVariables("REMOTE_ADDR")).toString();
	if (myusername == "" || myusername=="null" || myusername == "undefined")
		myusername = new String(myaddress).toString();
	connobj=Server.CreateObject("ADODB.Connection");
	RS=Server.CreateObject("ADODB.Recordset");
	connobj.Open(dbconnect);
	thedate = new Date();
	thetime = timeasstring(thedate);
	// Clear out any old sign-in record for this theref or theid
	SQL1=new String("delete from user_sessions where remoteaddress = '"+myaddress+"' or remoteuser = '"+myusername+"'").toString();
	RS=connobj.Execute(SQL1);
	// Add new sign-on record - always default to non-debug mode
	SQL2 = new String("INSERT into User_Sessions values('"+myusername+"','"+myaddress+"',#"+ddmmyyyy(thedate)+"#,#"+thetime+"#,'"+theid+"','"+thename+"','Normal')").toString();
	RS=connobj.Execute(SQL2);
	// Response.Cookies("member")("memberid")=new String(theid);
	// Response.Cookies("member")("membername")=new String(thename);
	// Response.Cookies("member").Expires="December 31,2037";
	// Response.Cookies("member")("acclevel")=thelevel;
	connObj=null;
	return(true);
}
//==================================================
function myAddress() {
	var theaddress = new String("").toString();
	theaddress = new String(Request.ServerVariables("REMOTE_ADDR")).toString();
	if (theaddress=="" || theaddress == "null" || theaddress == "undefined")
		theaddress = new String("NOT KNOWN").toString();
	return(theaddress.valueOf());
}
//==================================================
function getSessionMode() {
// Get debug indicator for current user session
	var zz, result;
	var mysessionmode;
	var myaddress;
	var dbconnect=Application("hamptonsportsdb"); 
	var connobj, RS, SQL2;
	mysessionmode = new String("Normal").toString();
	myaddress = new String(Request.ServerVariables("REMOTE_ADDR")).toString();
	connobj=Server.CreateObject("ADODB.Connection");
	RS=Server.CreateObject("ADODB.Recordset");
	connobj.Open(dbconnect);
	// Get info from sign-on record
	SQL2 = new String("SELECT sessionmode FROM User_Sessions where remoteaddress = '"+myaddress+"'").toString();
	v_memberid = "NONE";
	RS=connobj.Execute(SQL2);
	if (! RS.EOF)
	{
		mysessionmode = new String(RS("sessionmode")).toString();
		if (mysessionmode == "" || mysessionmode=="null" || mysessionmode == "undefined")
			mysessionmode = new String("Normal").toString();
	}
	RS.Close();
	RS=null;
	connobj.Close();
	connobj=null;
	return(mysessionmode.valueOf());
}
//==================================================
function setSessionMode(themode) {
	// Get debug indicator for current user session
	var zz, result;
	var mysessionmode;
	var myaddress;
	var dbconnect=Application("hamptonsportsdb"); 
	var connobj, RS, SQL2;
	mysessionmode = new String("Normal").toString();
	myaddress = new String(Request.ServerVariables("REMOTE_ADDR")).toString();
	connobj=Server.CreateObject("ADODB.Connection");
	RS=Server.CreateObject("ADODB.Recordset");
	connobj.Open(dbconnect);
	// Get info from sign-on record
	SQL2 = new String("UPDATE user_sessions set sessionmode ='"+themode+"' where remoteaddress = '"+myaddress+"'").toString();
	v_memberid = "NONE";
	RS=connobj.Execute(SQL2);
	RS.Close();
	RS=null;
	connobj.Close();
	connobj=null;
	return(true);
}
//==================================================//==================================================
function current_debug_status() {
	// Get debug indicator for current user session
	var result = false;
	var mysessionmode = getSessionMode();
	if (mysessionmode == "Debug")
		result = true;
	return(result);
}
</script>