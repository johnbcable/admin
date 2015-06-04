
// Set up global variables to be available to each page

var signedin = new String("N").toString();  // Default to N - updated by setSignedIn()
var debugthis = false;      // set to true to get console output
var myMemberID = new String("").toString();         // set to memberid of petrson who has signed in
var mySignIn = new Object();
var myName = new String("").toString();
var myPin = new String("").toString(); 
var myUserId = new String("").toString(); 

//
// Define utility functions first
//
// ------------------------------------------------------------------
// Function to set the signed-in status of the current user

function setSignedIn() {

	// Converted to use localStorage as has been set by logMeIn 

	var jsonstring = new String("");
	var tmpsignin;

	jsonstring = localStorage.getItem("tennissignin");

	if (debugthis) {
		console.log("Retrieved signin string is:    "+jsonstring);
	}

	if (jsonstring) {
		// We have retrieved some local storage data - set up summary info into globals
		mySignIn = $.parseJSON(jsonstring); 
		// mySignin = tmpsignin[0];
		signedin = new String("Y").toString();    // For semi-legacy compatibility
		myMemberID = new String(mySignIn[0].memberid).toString();    // For semi-legacy compatibility
		myName = new String(mySignIn[0].forename1).toString();    // Set my name in global variable
		myUserId = new String(mySignIn[0].onlinebookingid).toString();
		myPin = new String(mySignIn[0].onlinebookingpin).toString();

		if (debugthis) {
			console.log("Value of signedin inside setSignedIn is ["+signedin+"]");
			console.log("Value of mySignIn inside setSignedIn is ["+JSON.stringify(mySignIn)+"]");
			console.log("Values of global variables inside setSignedIn");
			console.log("myMemberID is ["+myMemberID+"]");
			console.log("myName is ["+myName+"]");
			console.log("myUserId is ["+myUserId+"]");
			console.log("myPin is ["+myPin+"]");
		}
				
	}    

	showMenus();    // Make sure and set the bottom menus

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

// ------------------------------------------------------------------
// Function to try and log me in
function logMeIn() {

	// Take sinfo from the loginModal and signs you in

    var userid = $('#frmUserId').val();
    var pin = $('#frmPassword').val();
    var refurl = $('#refurl').attr("value");
    var signinData = new Object();
    var signinurl = new String("").toString();

	var url = "http://hamptontennis.org.uk/fetchJSON.asp?id=11&p1="+userid+"&p2="+pin;
	var goodlogin = "N";

	// Re-initialise global variables on new login
	signedin = new String("N").toString();  // define globally so can use the $.Deferred() jQuery construct
	myMemberID = new String("").toString();         // set to memberid of petrson who has signed in
	mySignIn = new Object();      // Holds details from the localStorage object
	myName = new String("").toString();
	myPin = new String("").toString(); 
	myUserId = new String("").toString(); 

	if (debugthis) {
		console.log(url);
	}

	$.getJSON(url,function(allData) {

		goodlogin = (allData.length > 0) ? "Y" : "N";

		if (debugthis) {
			console.log("Value of goodlogin is ["+goodlogin+"]");
	    	console.log("userid = ["+userid+"] and PIN = ["+pin+"], refurl = ["+refurl+"]");
		}

		if ( allData.length > 0 ) {

			// Successful sign-in 

			myUserId = new String(userid).toString();
			myPin = new String(pin).toString();
			myName = new String(allData[0].forename1).toString();
			myMemberID = new String(allData[0].memberid);

			if (debugthis) {
				console.log("Value of signedin inside logMeIn is ["+signedin+"]");
				console.log("Value of mySignIn inside logMeIn is ["+JSON.stringify(mySignIn)+"]");
				console.log("Values of global variables inside logMeIn");
				console.log("myMemberID is ["+myMemberID+"]");
				console.log("myName is ["+myName+"]");
				console.log("myUserId is ["+myUserId+"]");
				console.log("myPin is ["+myPin+"]");
			}

			var jsonstring = JSON.stringify(allData);

			// jsonstring = new String("{thisMember:"+jsonstring+"}");

			// var eventdata = $.parseJSON(jsonstring);
			var memberdata = eval("(" + jsonstring + ")");
			signinData = new Object(memberdata);

			// if (debugthis) {
			//	console.log("Value of memberdata object:   "+jsonstring);
			// }

			// Now, set localStorage up
			// if (debugthis) {
			//	console.log("About to set up data into mySignIn object");
			// }

			// mySignIn.uniqueref = new String(memberdata.uniqueref).toString();
			// mySignIn.onlinebookingid = new String(memberdata.onlinebookingid).toString();
			// mySignIn.onlinebookingpin = new String(memberdata.onlinebookingpin).toString();
			// mySignIn.memberid = new String(memberdata.memberid).toString();
			// mySignIn.forename = new String(memberdata.forename1).toString();
			// mySignIn.acclevel = new String(memberdata.webaccess).toString();

			if (debugthis) {
				// alert("Value of forename1, onlinebookingid and onlinebookingpin = ["+myName+","+myUserId+","+myPin+"]");
				console.log("Value of local signinData object after updates: "+JSON.stringify(signinData));
			}
		
			localStorage.setItem("tennissignin", JSON.stringify(signinData));

			// Complete signin process using server-side ASP page
			signinurl = new String("http://hamptontennis.org.uk/signin.asp?u="+userid+"&p="+pin+"&d="+refurl).toString();

			if (debugthis) {
				console.log("Signin server-side URL = "+signinurl);
			}
			
			// N.B.  Server side will do redirect afterwards

			// window.location.href = signinurl;
			$('#signin').submit();   // Try this way - refurl not being picked up
			
		}

	});  // end of function(allData)	  	

}    // end of logMeIn

// ------------------------------------------------------------------
// Function to log me out

function logMeOut() {
	
	var signouturl = new String("").toString();

	// Remove the data from localstorage
	localStorage.removeItem("tennissignin");

	// Reset global variables
	myMemberID = new String("").toString();    // For semi-legacy compatibility
	myName = new String("").toString();    // Set my name in global variable

	// Complete signout process using server-side ASP page
	signouturl = new String("http://hamptontennis.org.uk/signout.asp").toString();

	if (debugthis) {
		console.log("Sign-out server-side URL = "+signouturl);
	}
	
	// N.B.  Server side will do redirect afterwards

	window.location.href = signouturl;

}    // end of logMeOut()

// ------------------------------------------------------------------
// Function to route me appropriately if signed-in

function routeMember(destinationurl) {

	var myurl = destinationurl || "/members.html";

	if (debugthis) {
		console.log("Value of signedin in routeMember is ["+signedin+"]");
		console.log("Value of myip in routeMember is ["+myip+"]");
		console.log("Value of destinationurl in routeMember is ["+destinationurl+"]");
		console.log("Value of myurl in routeMember is ["+myurl+"]");
	}

	if (signedin == "N") {
		if (debugthis)
			console.log("Routing to login page ...");
		window.location.href = "/login.html";
	} 
	if (signedin == "Y") {
		if (debugthis)
			console.log("Routing to "+myurl);
		window.location.href = myurl;
	}
}    // end of routeMember(url) 

// 
// Now kick all this off
//


$(document).ready(function() {

	// Blank out welcome block to start off with
	$('.welcomeblock').css('display','none');

	setSignedIn();   // This should reset the signed0-in status and global variables
    
	// Set option text on home page banner image
	if (signedin == "Y") {
		optiontext = new String("Members").toString();
	} else {
		optiontext = new String("Sign-In").toString();
	}

	if (debugthis) {
		console.log("Signed-in status = "+signedin);
		console.log("optiontext = "+optiontext);
	}

	$('#memberslink').text(optiontext);

	// Now construct welcome message, set HTML and show the welcome block

	if (signedin == "Y") {
		var welcomemessage = "Hi "+myName+" - welcome back";
		$('#welcometext').html(welcomemessage);
		$('#notme').html("Not "+myName+"?");
		$('.welcomeblock').css('display','');
	}

	$( '#loginmenu' ).click(function() {
		$('#frmUserId').focus();
	});

	$('#loginSubmit').click(function(){
		logMeIn();
	});

	$('#logout').click(function(){
		logMeOut();
	});
/*
	$('#mylogout').click(function(){
		logMeOut();
	});

	$('#notme').click(function(){
		logMeOut();
	});
*/
	$(document).on('opened', '[data-reveal]', function () {
	    $("#frmUserId:visible").focus();
	});	


 });    // document.ready



