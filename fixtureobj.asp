<%
//
//  FixtureObject
//
//  Holds detail for a single Coventry League fixture 
//

function FixtureObject(id,fixturedate,teamname,season) {

	this.fixtureid = id || 0;
	this.fixturedate = new String(fixturedate).toString() || new String(Date()).toString();
	this.fixtureyear = season || 0;
	this.teamname = new String(teamname).toString() || new String("").toString();

	this.homeoraway = new String("H").toString();
	this.opponents = new String("NONE").toString();
	this.hamptonresult = 0;
	this.opponentresult = 0;
	this.fixturenote = new String("NONE").toString();
	this.matchreport = new String("NONE").toString();
	this.pair1 = new String("").toString();
	this.pair2 = new String("").toString();
};

FixtureObject.prototype.setVenue = function(venue) {
	this.homeoraway = new String(venue).toString();
};

FixtureObject.prototype.setOpponents = function(opponents) {
	this.opponents = new String(opponents).toString();
};

FixtureObject.prototype.setResult = function(us,them) {
	this.hamptonresult = us;
	this.opponentresult = them;
};

FixtureObject.prototype.setNote = function(notetext) {
	this.fixturenote = new String(notetext).toString();
};

FixtureObject.prototype.setMatchReport = function(reportext) {
	this.matchreport = new String(reporttext).toString();
};

FixtureObject.prototype.setPairs = function(pair1, pair2) {
	this.pair1 = new String(pair1).toString();
	this.pair2 = new String(pair2).toString();
};

FixtureObject.prototype.setPair1 = function(pair1) {
	this.pair1 = new String(pair1).toString();
};

FixtureObject.prototype.setPair2 = function(pair2) {
	this.pair2 = new String(pair2).toString();
};

%>
