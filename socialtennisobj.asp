<%
//
//  SocialTennisObject
//
//  Holds detail for a single social tennis session
//

function SocialTennisObject(thedate) {

	var mydate = new Date(thedate);

	this.socialdate = new String(thedate).toString() || new String("").toString();

	// Determine the day from the date
	var nyday = mydate.getDay();

	// Determine start end end times from the day
	// Tues (2) and Fri (5)  start at 20:00 and end at 22:00 (default)
	// Sun (0) starts at 10:00 and ends at 12:00
	// Sat (6) starts at 15:00 and ends at 

	switch(myday) {
    case 6:    // Saturday
        this.start_time = new String("15:00:00").toString();
        this.end_time = new String("17:00:00").toString();
        this.socialtext = new String("Saturday afternoon family social tennis")
        break;
    case 0:   // Sunday
        this.start_time = new String("10:00:00").toString();
        this.end_time = new String("12:00:00").toString();
        this.socialtext = new String("Sunday morning social tennis")
        break;
    case 2:   // Tuesday
        this.start_time = new String("20:00:00").toString();
        this.end_time = new String("22:00:00").toString();
        this.socialtext = new String("Tuesday evening social tennis")
        break;
    case 5:   // Friday
        this.start_time = new String("20:00:00").toString();
        this.end_time = new String("22:00:00").toString();
        this.socialtext = new String("Friday evening social tennis")
        break;
    default:    // Tuesday and Friday
        this.start_time = new String("20:00:00").toString();
        this.end_time = new String("22:00:00").toString();
        this.socialtext = new String("UNKNOWN SESSION")
	}

};

SocialTennisObject.prototype.setStartTime = function(start) {
	this.start_time = new String(start).toString();
};

SocialTennisObject.prototype.setEndTime = function(finish) {
	this.end_time = new String(finish).toString();
};

SocialTennisObject.prototype.setDate = function(thedate) {

	this.socialdate = new String(thedate).toString();
};


%>
