<%@language="JScript"%>
<!--#include file="functions.asp" -->
<!--#include file="strings.asp" -->
<!--#include file="memberfuncs.asp" -->
<!--#include file="coachfuncs.asp" -->
<%
// Now for any variables local to this page
var newone = false;   // default to update rather than create
var m_displayorder, m_url, m_hourlyrate, m_halfhourlyrate;
var m_voucher1, m_voucher2, m_voucher3, m_imageurl;
var m_id;   // uniqref of the coach being updated/created
var memberObj = new Object();
var RS, Conn, dbconnect;
var SQL1, SQLText;
var SQL2 = new String("").toString();

var debugging=current_debug_status();
var updating=true;

//
// debugging = true;

// Process form/querystring parameters
m_id = Trim(new String(Request.Form("id")));
if (m_id == "" || m_id =="null" || m_id == "undefined")
{
	m_id = Trim(new String(Request.QueryString("id")));
	if (m_id == "" || m_id =="null" || m_id == "undefined")
	{
		m_id = new String("0").toString();
	}
}
// Now get other form/querystring variables
m_url = Trim(new String(Request.Form("url")));
if (m_url == "" || m_url =="null" || m_url == "undefined")
{
	m_url = Trim(new String(Request.QueryString("id")));
	if (m_url == "" || m_url =="null" || m_url == "undefined")
	{
		m_url = new String("coaching_staff.html").toString();
	}
}
m_hourlyrate = Trim(new String(Request.Form("hourlyrate")));
if (m_hourlyrate == "" || m_hourlyrate =="null" || m_hourlyrate == "undefined")
{
	m_hourlyrate = Trim(new String(Request.QueryString("id")));
	if (m_hourlyrate == "" || m_hourlyrate =="null" || m_hourlyrate == "undefined")
	{
		m_hourlyrate = new String("10").toString();
	}
}
m_halfhourlyrate = Trim(new String(Request.Form("halfhourlyrate")));
if (m_halfhourlyrate == "" || m_halfhourlyrate =="null" || m_halfhourlyrate == "undefined")
{
	m_halfhourlyrate = Trim(new String(Request.QueryString("id")));
	if (m_halfhourlyrate == "" || m_halfhourlyrate =="null" || m_halfhourlyrate == "undefined")
	{
		m_halfhourlyrate = new String("5").toString();
	}
}
m_voucher1 = Trim(new String(Request.Form("voucher1")));
if (m_voucher1 == "" || m_voucher1 =="null" || m_voucher1 == "undefined")
{
	m_voucher1 = Trim(new String(Request.QueryString("id")));
	if (m_voucher1 == "" || m_voucher1 =="null" || m_voucher1 == "undefined")
	{
		m_voucher1 = new String("null").toString();
	}
}
m_voucher2 = Trim(new String(Request.Form("voucher2")));
if (m_voucher2 == "" || m_voucher2 =="null" || m_voucher2 == "undefined")
{
	m_voucher2 = Trim(new String(Request.QueryString("id")));
	if (m_voucher2 == "" || m_voucher2 =="null" || m_voucher2 == "undefined")
	{
		m_voucher2 = new String("null").toString();
	}
}
m_voucher3 = Trim(new String(Request.Form("voucher3")));
if (m_voucher3 == "" || m_voucher3 =="null" || m_voucher3 == "undefined")
{
	m_voucher3 = Trim(new String(Request.QueryString("id")));
	if (m_voucher3 == "" || m_voucher3 =="null" || m_voucher3 == "undefined")
	{
		m_voucher3 = new String("null").toString();
	}
}
m_imageurl = Trim(new String(Request.Form("imageurl")));
if (m_imageurl == "" || m_imageurl =="null" || m_imageurl == "undefined")
{
	m_imageurl = Trim(new String(Request.QueryString("id")));
	if (m_imageurl == "" || m_imageurl =="null" || m_imageurl == "undefined")
	{
		m_imageurl = new String("").toString();
	}
}

// Ok we know if it is a brand new coach by now

if (! (m_id == "0"))  {
	memberObj = getCoach(m_id);  // Retrieve the existing coach record
}  else {
	memberObj = getMember(m_id);  // Retrieve the existing member record
	// ... and augment with defauts for additional coach fields
	memberObj.displayorder = 12;  // default to high number to be last
	memberObj.url = new String("").toString();
	memberObj.hourlyrate = 10;
	memberObj.halfhourlyrate = 5;
	memberObj.voucher1 = "null";
	memberObj.voucher2 = "null";
	memberObj.voucher3 = "null";
	memberObj.imageurl = new String("").toString();
}  

// Update coach Object with data from submitting form

memberObj.url = new String(m_url).toString();
memberObj.hourlyrate = m_hourlyrate;
memberObj.halfhourlyrate = m_halfhourlyrate;
memberObj.voucher1 = m_voucher1;
memberObj.voucher2 = m_voucher2;
memberObj.voucher3 = m_voucher3;
memberObj.imageurl = new String(m_imageurl).toString();












// Processing finished - now return to the list of coaches

Response.Redirect("managecoaches.html");

%>




// Check if this member is already in the list of coaches
SQL1 = "SELECT uniqref FROM coaches WHERE uniqref = "

mNew = new String(Request.Form("NewMember")).toString();
if (mNew == "" || mNew == "null" || mNew == "undefined") {
	mNew = new String ("N");
}
mNew = mNew.toUpperCase();
// End of page start up coding

// Initialise update variables from prior form
SQLText = new String("").toString();
SQLAudit = new String("").toString();
SQL2 = new String("").toString();
memberObj.initials=new String("").toString();
memberObj.paid=new String("Y").toString();
zz3 = Trim(new String(Request.Form("MemberId")));
memberObj.id = Trim(new String(Request.Form("MemberId"))).toString();
memberObj.uniqueref = Trim(new String(Request.Form("Uniqueref"))).toString();
memberObj.surname = StringToDB(Request.Form("MemberSurname"));
memberObj.forename1 = Trim(new String(Request.Form("MemberForename"))).toString();
memberObj.gender = Trim(new String(Request.Form("MemberGender"))).toString();
memberObj.grade = Trim(new String(Request.Form("MemberGrade"))).toString();
memberObj.address1 = Trim(new String(Request.Form("MemberAddress1"))).toString();
memberObj.address2 = Trim(new String(Request.Form("MemberAddress2"))).toString();
memberObj.address3 = Trim(new String(Request.Form("MemberAddress3"))).toString();
memberObj.address4 = Trim(new String(Request.Form("MemberAddress4"))).toString();
memberObj.postcode = Trim(new String(Request.Form("MemberPostcode"))).toString();
memberObj.mobile = Trim(new String(Request.Form("MemberMobilePhone"))).toString();
memberObj.homephone = Trim(new String(Request.Form("MemberHomePhone"))).toString();
memberObj.email = Trim(new String(Request.Form("MemberEmail"))).toString();
memberObj.webpassword = Trim(new String(Request.Form("MemberPassword"))).toString();
memberObj.webaccess = Trim(new String(Request.Form("MemberWebAccess"))).toString();
memberObj.mailing = Trim(new String(Request.Form("MemberMailing"))).toString();
memberObj.internalleague = Trim(new String(Request.Form("InternalLeague"))).toString();
memberObj.onlinebookingid = Trim(new String(Request.Form("OnlineBookingID"))).toString();
memberObj.onlinebookingpin = Trim(new String(Request.Form("OnlineBookingPIN"))).toString();
memberObj.joined = Trim(new String(Request.Form("MemberJoined"))).toString();
memberObj.left = Trim(new String(Request.Form("MemberLeft"))).toString();
memberObj.dob = Trim(new String(Request.Form("MemberDOB"))).toString();
memberObj.pool = Trim(new String(Request.Form("MemberPool"))).toString();
memberObj.maxitennis = Trim(new String(Request.Form("MemberMaxiTennis"))).toString();
memberObj.britishtennisno = Trim(new String(Request.Form("BritishTennisNumber"))).toString();
memberObj.photoconsent = Trim(new String(Request.Form("PhotoConsent"))).toString();
memberObj.adultcontact = Trim(new String(Request.Form("AdultContact"))).toString();
memberObj.adultrelationship = Trim(new String(Request.Form("AdultRelationship"))).toString();
memberObj.adultphone = Trim(new String(Request.Form("AdultPhone"))).toString();
memberObj.adultmobile = Trim(new String(Request.Form("AdultMobile"))).toString();
memberObj.specialcare = Trim(new String(Request.Form("SpecialCare"))).toString();
memberObj.iscoach = Trim(new String(Request.Form("IsCoach"))).toString();

// MUST set surname to upper case otherwise searching will NOT work
memberObj.surname = memberObj.surname.toUpperCase();
// Set forename into proper (title) case
memberObj.forename1 = toProperCase(memberObj.forename1);

// Calculate current age 
mTextAge = new String("").toString();
mAge = 0;
if (memberObj.dob == "" || memberObj.dob == "null" || memberObj.dob == "undefined") {
	mTextAge = new String("Not Known").toString();
	mAge = -1;
} else {
	mAge = currentAge(memberObj.dob);
	mTextAge = new String(mAge).toString();	
}

// Response.Write("onlinebookingid before setMember is = ["+memberObj.onlinebookingid+"]<br />")
// Do DB update
if (updating)
{
	result = setMember(memberObj);
	resulttext += "<br />"+result.result+"<br />";
	resulttext += "<br />"+result.errno+"<br />";
	resulttext += "<br />"+result.description+"<br />";
	resulttext += "<br /><hr>"+result.sql+"<br />";
}	
pagetitle += " for " + memberObj.forename1 + " " + memberObj.surname;
// set real unique id
realmemberid = zz3;
dbconnect=Application("hamptonsportsdb");
Conn = Server.CreateObject("ADODB.Connection");
RS = Server.CreateObject("ADODB.RecordSet");
Conn.Open(dbconnect);
if (mNew == "Y") 
{
	realmemberid = zz3.substr(2);
	SQL2 = "update members set memberid = '"+realmemberid+"' where uniqueref = "+memberObj.uniqueref;
	if (updating) 
	{
		try {
			RS = Conn.Execute(SQL2);
		}
		catch(e) {
			if (!(e.number &0xFFFF) == 0) 
			{
				Response.Write("Error updating member ID to ["+realmemberid+"]<br /><br />");
				Response.Write("Facility code:        "+e.number>>16 & 0x1FFF);
				Response.Write("<br />");
				Response.Write("Error code:           "+e.number & 0xFFFF);
				Response.Write("<br />");
				Response.Write("Description:          "+e.description+"<br />");
			}  // end if		
		}  // end catch
	}  // end if updating 
}  // end mNew

// Now write audit record
SQLaudit = new String("insert into member_audits([memberid],[action]) values ('"+v_memberid+"','MEMBER UPDATE ("+memberObj.uniqueref+" - "+memberObj.forename1+" "+memberObj.surname+")')");
if (debugging)
	Response.Write("SQLaudit = ["+SQLaudit+"]</br />");
RS = Conn.Execute(SQLaudit);
//
RS=null;
Conn.Close();
Conn=null;
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<META NAME="ROBOTS" CONTENT="NOINDEX, NOFOLLOW, NOARCHIVE">
	<title>Hampton-In-Arden Tennis Club Home Page</title>
	<link rel="stylesheet" media="screen" type="text/css" href="layout.css" />
	<link rel="stylesheet" media="screen" type="text/css" href="colours.css" />
	<link rel="stylesheet" media="screen" type="text/css" href="typography.css" /> 
	<link rel="stylesheet" media="print" type="text/css" href="print3col.css" /> 
	<!-- Comment out the next style sheet if runniung in production -->
<%
if (debugging)
{
%>
<link rel="stylesheet" media="screen" type="text/css" href="borders.css" />
<%
}
%>
<style type="text/css">
<!--
#welcomediv {
				 position: relative;
				 margin: 20px auto;
}
#welcomebutton {
					 position: relative;
					 float: left;
					 width: 20%;
					 text-align: center;
}
#welcometext {
			 position: relative;
			 float: right;
			 width: 78%;
			 text-align: left;
}
#welcometext p {
			 padding-left: 30px;
}						

#welcomebutton a, #welcomebutton a:visited {
	background: #c9785a; 
	display: inline-block; 
  font-size: 140%;
  font-weight: bold;
	padding: 5px 10px 6px; 
	color: #fff; 
	text-decoration: none;
	-moz-border-radius: 5px; 
	-webkit-border-radius: 5px;
	-moz-box-shadow: 0 1px 3px rgba(0,0,0,0.5);
	-webkit-box-shadow: 0 1px 3px rgba(0,0,0,0.5);
	text-shadow: 0 -1px 1px rgba(0,0,0,0.25);
	border-bottom: 1px solid rgba(0,0,0,0.25);
	position: relative;
	cursor: pointer;
}

#welcomebutton a:hover { 
	background-color: #111; 
	color: #fff; 
}

-->
</style>
	<script language="Javascript" src="datetimepicker.js"></script>
</head>
<body>
<!--#include file="dateval.asp" -->
<!--   1.  Branding   -->
<div id="branding">
	<a href="index.asp" id="homelink"><img id="clublogo" src="images/logo.gif" alt="Hampton-In-Arden Sports Club logo" /></a>
	<h1>Hampton-In-Arden Sports Club</h1>
	<h2>Tennis Section</h2>

<!--   2.  Navigation   -->
	<div id="nav_main">
		<ul id="topmenu">
			<li id="nav_ourclub"><a href="aboutus.asp">Our Club</a></li>
			<li id="nav_coaching"><a href="juniors/coaching.html">Coaching</a></li>
			<li id="nav_playing"><a href="playing.asp">Playing</a></li>
			<li id="nav_links"><a href="juniors/index.html">Juniors</a></li>
			<li id="nav_contact"><a href="juniors/contact.html">Contact</a></li>
			<li id="nav_members"><a href="members.asp">Members</a></li>
		</ul>
		<p id="today">
			<%= displaydate %>&nbsp;<%= strtime %>
		</p>
	</div>
	
</div>

<div id="wrapper">

<!--   3. Content    -->
	<div id="content">
		<h1>Administration:&nbsp;<b>Updating details for <%= memberObj.forename1 %>&nbsp;<%= memberObj.surname %></b></h1>
<%
if (result.result)
{
	Response.Write("<p>Database updated.</p><br />");
	if (debugging) 
	{
		Response.Write("["+mNew+"] - [Y]<br />");
	}
	if (mNew == "Y")
	{
		Response.Write("<p>Sending reminder emails ... <br /><br />");
		memberObj.id = Trim(new String(realmemberid)).toString();

		// ========================================================================
		// Now send note to support so that entries onto online booking can be made.
		//=========================================================================

		// Read in contents of HTML email template to be used
		if (emaildir == "")
		{
			 sMessage = new String(GetEveryThing(supporttemplate));
			 fullwhichfile = new String(supporttemplate).toString();
		}
		else
		{
			 sMessage = new String(GetEveryThing(emaildir+"/"+supporttemplate));
			 fullwhichfile = new String(emaildir+"/"+supporttemplate).toString();
		}	 
		//
		/// Now do merge field replacements on support template
		var doctypeindex = sMessage.indexOf("<!DOCTYPE");
		var htmlindex = sMessage.indexOf("<html");
		var fullmessage = new String(sMessage.substr(doctypeindex)).toString();
		htmlindex = fullmessage.indexOf("<html");
		fullmessage = new String(fullmessage.substr(htmlindex)).toString();
		sMessage = new String(fullmessage.replace("ZZforenameZZ",memberObj.forename1)).toString();
		sMessage = new String(sMessage.replace("ZZsurnameZZ",memberObj.surname)).toString();
		sMessage = new String(sMessage.replace("ZZonlinebookingidZZ",memberObj.onlinebookingid)).toString();
		sMessage = new String(sMessage.replace("ZZonlinebookingpinZZ",memberObj.onlinebookingpin)).toString();
		sMessage = new String(sMessage.replace("ZZmembergradeZZ",memberObj.grade)).toString();
		sMessage = new String(sMessage.replace("ZZgenderZZ",memberObj.gender)).toString();
		sMessage = new String(sMessage.replace("ZZageZZ",mTextAge)).toString();

		objCDOMail = newMailObject();   // in emailfuncs
		
		objCDOMail.From=new String("secretary@hamptontennis.org.uk");
		objCDOMail.Subject=new String("New Member").toString();
		// objCDOMail.BodyFormat=0;
		// objCDOMail.MailFormat=0;
		// objCDOMail.Body=new String(tMessage).toString();
		objCDOMail.HTMLBody=new String(sMessage).toString();
		
		objCDOMail.To=new String("johncable@talk21.com").toString();
		// Response.Write("Recipient is johncable@talk21.com<br>");
		try {
				objCDOMail.Send();
				}
		catch(e) {
			if (!(e.number &0xFFFF) == 0) 
			{
				Response.Write("Error sending email<br /><br />");
				Response.Write("Facility code:        "+e.number>>16 & 0x1FFF);
				Response.Write("<br />");
				Response.Write("Error code:           "+e.number & 0xFFFF);
				Response.Write("<br />");
				Response.Write("Description:          "+e.description+"<br />");
			} else 
			{
				kount += 1;
				Response.Write("&nbsp;&nbsp;&nbsp;Online booking email sent OK<br />")
			}

		}
		objCDOMail=null;

		if (! debugging) 
		{
			// ========================================================================
			// Secondly, send note to Membership Secretary if member is over 17 so that 
			// a new clubhouse key card can be set up.
			//=========================================================================
			if (mAge > 17)
			{

				// Read in contents of HTML email template to be used
				if (emaildir == "")
				{
					 sMessage = new String(GetEveryThing(idtemplate));
					 fullwhichfile = new String(idtemplate).toString();
				}
				else
				{
					 sMessage = new String(GetEveryThing(emaildir+"/"+idtemplate));
					 fullwhichfile = new String(emaildir+"/"+idtemplate).toString();
				}	 
				//
				/// Now do merge field replacements on id card template
				var doctypeindex = sMessage.indexOf("<!DOCTYPE");
				var htmlindex = sMessage.indexOf("<html");
				var fullmessage = new String(sMessage.substr(doctypeindex)).toString();
				htmlindex = fullmessage.indexOf("<html");
				fullmessage = new String(fullmessage.substr(htmlindex)).toString();
				sMessage = new String(fullmessage.replace("ZZforenameZZ",memberObj.forename1)).toString();
				sMessage = new String(sMessage.replace("ZZsurnameZZ",memberObj.surname)).toString();
				sMessage = new String(sMessage.replace("ZZonlinebookingidZZ",memberObj.onlinebookingid)).toString();
				sMessage = new String(sMessage.replace("ZZonlinebookingpinZZ",memberObj.onlinebookingpin)).toString();
				sMessage = new String(sMessage.replace("ZZmembergradeZZ",memberObj.grade)).toString();
				sMessage = new String(sMessage.replace("ZZgenderZZ",memberObj.gender)).toString();


				objCDOMail = newMailObject();   // in emailfuncs
				
				objCDOMail.From=new String("secretary@hamptontennis.org.uk");
				objCDOMail.Subject=new String("New Tennis Club Member").toString();
				// objCDOMail.BodyFormat=0;
				// objCDOMail.MailFormat=0;
				// objCDOMail.Body=new String(tMessage).toString();
				objCDOMail.HTMLBody=new String(sMessage).toString();
				
				objCDOMail.To=new String("juliewilliamson45@hotmail.com").toString();
				// Response.Write("Recipient is juliewilliamson45@hotmail.com<br>");
				try {
					objCDOMail.Send();
				}
				catch(e) {
					if (!(e.number &0xFFFF) == 0) 
					{
						Response.Write("Error sending email<br /><br />");
						Response.Write("Facility code:        "+e.number>>16 & 0x1FFF);
						Response.Write("<br />");
						Response.Write("Error code:           "+e.number & 0xFFFF);
						Response.Write("<br />");
						Response.Write("Description:          "+e.description+"<br />");
					} 
					else 
					{
						kount += 1;
						Response.Write("&nbsp;&nbsp;&nbsp;Membership Secretary email sent OK<br />")
					}
				}  // end catch
			}  // end mAge > 17
			else
			{
				Response.Write("&nbsp;&nbsp;&nbsp;No email sent to membership secretary as no key card necessary (age is "+mTextAge+")<br />")
			}

			objCDOMail=null;

			// ========================================================================
			// Finally, send note to committee of new member.
			//=========================================================================

			// Read in contents of HTML email template to be used
			if (emaildir == "")
			{
				 sMessage = new String(GetEveryThing(committeetemplate));
				 fullwhichfile = new String(committeetemplate).toString();
			}
			else
			{
				 sMessage = new String(GetEveryThing(emaildir+"/"+committeetemplate));
				 fullwhichfile = new String(emaildir+"/"+committeetemplate).toString();
			}	 
			//
			/// Now do merge field replacements on id card template
			var doctypeindex = sMessage.indexOf("<!DOCTYPE");
			var htmlindex = sMessage.indexOf("<html");
			var fullmessage = new String(sMessage.substr(doctypeindex)).toString();
			htmlindex = fullmessage.indexOf("<html");
			fullmessage = new String(fullmessage.substr(htmlindex)).toString();
			sMessage = new String(fullmessage.replace("ZZforenameZZ",memberObj.forename1)).toString();
			sMessage = new String(sMessage.replace("ZZsurnameZZ",memberObj.surname)).toString();
			sMessage = new String(sMessage.replace("ZZonlinebookingidZZ",memberObj.onlinebookingid)).toString();
			sMessage = new String(sMessage.replace("ZZonlinebookingpinZZ",memberObj.onlinebookingpin)).toString();
			sMessage = new String(sMessage.replace("ZZmembergradeZZ",memberObj.grade)).toString();
			sMessage = new String(sMessage.replace("ZZgenderZZ",memberObj.gender)).toString();

			objCDOMail = newMailObject();   // in emailfuncs
			
			objCDOMail.From=new String("secretary@hamptontennis.org.uk");
			objCDOMail.Subject=new String("New Tennis Club Member").toString();
			// objCDOMail.BodyFormat=0;
			// objCDOMail.MailFormat=0;
			// objCDOMail.Body=new String(tMessage).toString();
			objCDOMail.HTMLBody=new String(sMessage).toString();
			
			objCDOMail.To=new String("committee@hamptontennis.org.uk").toString();
			// Response.Write("Recipient is committee@hamptontennis.org.uk<br>");
			try {
				objCDOMail.Send();
			}
			catch(e) {
				if (!(e.number &0xFFFF) == 0) 
				{
					Response.Write("Error sending email<br /><br />");
					Response.Write("Facility code:        "+e.number>>16 & 0x1FFF);
					Response.Write("<br />");
					Response.Write("Error code:           "+e.number & 0xFFFF);
					Response.Write("<br />");
					Response.Write("Description:          "+e.description+"<br />");
				} 
				else 
				{
					kount += 1;
					Response.Write("&nbsp;&nbsp;&nbsp;Committee email sent OK<br />")
				}

			}  // end catch
			objCDOMail=null;

		}   // end debugging
	}    // end mNew == Y
%>
			<table id="detailstable" width="100%" cellpadding="5">
				<thead>
					<tr style="height: 2em;">
						<th></th>
						<th></th>
					</tr>
				</thead>
				<tbody>
					<tr>
						<td class="tdnine" valign="CENTER" align="RIGHT">
							<b>Name:&nbsp;&nbsp;&nbsp;&nbsp;</b>
						</td>
						<td>
							<%= memberObj.forename1 %>&nbsp;<%= memberObj.initials %>&nbsp;<%= memberObj.surname %><br>
						</td>
					</tr>
					<tr>
						<td class="tdnine" valign="CENTER" align="RIGHT">
							<b>	Address Details:&nbsp;&nbsp;&nbsp;&nbsp;</b>
						</td>
						<td>
							<%= memberObj.address1 %><br>
							<%= memberObj.address2 %><br>
							<%= memberObj.address3 %><br>
							<%= memberObj.address4 %><br>
							<%= memberObj.postcode %><br>
						</td>
					</tr>
					<tr>
						<td class="tdnine" valign="CENTER" align="RIGHT">
							            <b>Phone Numbers:&nbsp;&nbsp;&nbsp;&nbsp;</b>
						</td>
						<td>
							<b>Home&nbsp;&nbsp;:</b> <%= memberObj.homephone %><br>
							<b>Mobile:</b> <%= memberObj.mobile %><br>
						</td>
					</tr>
					<tr>
						<td class="tdnine" valign="CENTER" align="RIGHT">
							<b>Email Address:&nbsp;&nbsp;&nbsp;&nbsp;</b>
						</td>
						<td>
							<%= memberObj.email %><br>
						</td>
					</tr>
					<tr>
						<td class="tdnine" valign="CENTER" align="RIGHT">
							<b>Player Pool:&nbsp;&nbsp;&nbsp;&nbsp;</b>
						</td>
						<td>
							<%= memberObj.pool %><br>
						</td>
					</tr>
<%
	mylevel = getUserAccessLevel();
	theirlevel = memberObj.webaccess;
	if (mylevel >= theirlevel)	
	{
%>	
					<tr>
						<td colspan="2" style="text-align:left;padding-left:10px;background:#ccc">
							<b>Online Court Booking</b>
						</td>
					</tr>
					<tr>
						<td class="tdnine" valign="CENTER" align="RIGHT">
							<b>User ID:&nbsp;&nbsp;&nbsp;&nbsp;</b>
						</td>
						<td>
							<%= memberObj.onlinebookingid %><br>
						</td>
					</tr>
					<tr>
						<td class="tdnine" valign="CENTER" align="RIGHT">
							<b>PIN:&nbsp;&nbsp;&nbsp;&nbsp;</b>
						</td>
						<td>
							<%= memberObj.onlinebookingpin %><br>
						</td>
					</tr>
<%
	}
%>					
				</tbody>
				<tfoot>
					<tr>
						<th></th>
						<th></th>
					</tr>
				</tfoot>
			</table>
<%	
}
else
{
	Response.Write("<p>Database NOT updated</p><br />");
	Response.Write("<p>"+resulttext+"</p>");
	Response.End();
}
if (debugging)
{
	Response.Write("<br />["+SQLText+"]<br />");
	Response.Write("<br />["+SQL2+"]<br />");
	Response.Write("<br />["+SQLAudit+"]<br />");
}
%>
		<h2>What would you like to do now?</h2>
		<div>
			<ul>
				<li><a href="editmemb2.asp?code=<%= memberObj.uniqueref %>">Edit <%= memberObj.forename1 %>'s details again</a><br /><br /></li>
				<li><a href="chk_memb.asp">Search members again</a><br /><br /></li>
<%
if (memberObj.email != "")
{
	if (isAdministrator())
	{
%>
				<div id="welcomediv" style="clear: both; padding: 20px 0; margin-top: 20px">
					 <div id="welcomebutton">
					 	<a href="sendwelcome.asp?uniqueref=<%= memberObj.uniqueref %>">Welcome Email</a>
					 </div>
					<div id="welcometext">
						<p>
							Clicking on the &quot;Welcome Email&quot; button will send the club standard welcome email to <%= memberObj.forename1 %>. The email will confirm to <%= memberObj.forename1 %> the User ID and PIN they can use to access this web site and to book courts.
						</p>
					</div>
				</div>
				<p style="clear: both; margin-bottom: 10px;">&nbsp;</p>
<%
	}
}
%>		
			</ul>
		</div>
	</div>
	
<!--     4.      Supplementary navigation    -->
	<div id="leftcolumn">
		<!--#include file="fixturepanel.asp" -->
		<!--#include file="memb_nav.asp" -->
	</div>

<!--    5.   Supplementary content     -->	
	<div id="rightcolumn">
		<!--#include file="searchpanel.asp" -->
		<!--#include file="addresspanel.asp" -->
		<!--#include file="newspanel.asp" -->
		<!--#include file="emailpanel.asp" -->
	</div>
</div>

<!--     6.    Site info     -->
<!--#include file="footer.asp" -->

</body>
</html>
<%
%>

