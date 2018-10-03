//
//  ResultObject
//
//  Holds result of site operations and can be returned as JSON to requester
//

function ResultObject() {
	this.result = true;
	this.successmessages = new Array();
	this.errormessages = new Array();
	this.debugmessages = new Array();
	this.err = new Object();
	this.sql = new String("").toString();
	this.data = new Array();
};

ResultObject.prototype.addSuccessMessage = function(themessage) {
	this.successmessages.push(themessage);
};
ResultObject.prototype.addErrorMessage = function(themessage) {
	this.errormessages.push(themessage);
};
ResultObject.prototype.addDebugMessage = function(themessage) {
	this.debugmessages.push(themessage);
};

