<%
//
//  EventObject
//
//  Holds details of an event on the club events calendar
//

function EventObject() {
	var dummy = new Date();
	this.eventdate = new Date();
	this.eventtime = new Now();
	this.eventyear = dummy.getFullYear();
	this.eventtype = new String("EVENT").toString();
	this.eventnote = new String("").toString();
	this.eventid = -1;   // Default to brand new event
	this.eventreport = new String("").toString();
	this.enddate = new Date();
	this.endtime = new Now();
	this.fixturelink = -1;
	this.tourlink = new String("tournaments.html").toString();
	this.holidaylink = new String("holidaycamps.html").toString();
	this.advert = new String("").toString();
};

/* 
EventObject.prototype.addSuccessMessage = function(themessage) {
	this.successmessages.push(themessage);
};
*/

%>
