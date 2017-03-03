
var signedin = new String("N").toString();  // define globally so can use the $.Deferred() jQuery construct
var debugthis = false;      // set to true to get console output

//
// Define utility functions first
//
// ------------------------------------------------------------------
// Function to set the signed-in status of the current user

function setSignedIn() {

	var jsonstring = new String("");
	var url = "https://hamptontennis.org.uk/fetchJSON.asp?id=7";
	
	// var ip = new String(Request.UserHostAddress);
	// var hostname = new String(Request.UserHostName);

	// Check gobal myip
	if (! myip) {
		signedin = new String("N").toString();
	} else {
		url += "&p1="+myip;
	/*
		alert("url is "+url);
	*/
		$.getJSON(url,function(allData) {

			// create a deferred object
			var s = $.Deferred();

			signedin = (allData.length > 0) ? "Y" : "N";
			if (debugthis)
				console.log("Value of signedin inside setSignedIn is ["+signedin+"]");

			showMenus();    // Make sure and set the bottom menus

		});  // end of function(data)
	}

}

// ------------------------------------------------------------------
// Function to make sure the correct menu is displayed in the footer

function showMenus() {

	if (debugthis) {
		console.log("Value of signedin is ["+signedin+"]");
		console.log("Value of myip is ["+myip+"]");
	}

	if (signedin == "N") {
		if (debugthis)
			console.log("I have NOT been found to be signed-in");
		$('.signedin').addClass('noshow');   // hide signed-in menu at first
		$('.notsignedin').removeClass('noshow');  // hide the Sign-in menu choice
	} 
	if (signedin == "Y") {
		if (debugthis)
			console.log("I AM signed-in");
		$('.notsignedin').addClass('noshow');
		$('.signedin').removeClass('noshow');
	}

}

