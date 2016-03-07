
// Set up global variables to be available to each page

var adminsignedin = new String("N").toString();  // Default to N - updated by setSignedIn()
var debugthis = true;      // set to true to get console output
var myMemberID = new String("").toString();         // set to memberid of petrson who has signed in
var myAdminSignIn = new Object();
var myName = new String("").toString();
var myPin = new String("").toString(); 
var myUserId = new String("").toString();
var myAccess = new String("20").toString();

//
// Define utility functions first
//
// ------------------------------------------------------------------
// Function to set the signed-in status of the current admin user

function setAdminSignedIn() {

	// Converted to use sessionStorage as has been set by logMeIn 

	var jsonstring = new String("");

	adminsignedin = new String("N").toString();  // default to not signed in

	jsonstring = sessionStorage.getItem("tennisadmin");

	if (debugthis) {
		console.log("Retrieved admin signin string is:    "+jsonstring);
	}

	if (jsonstring) {
		// We have retrieved some local storage data - set up summary info into globals
		myAdminSignIn = $.parseJSON(jsonstring); 
		// mySignin = tmpsignin[0];
		// 
		adminsignedin = new String("Y").toString();    // For semi-legacy compatibility
		myMemberID = new String(myAdminSignIn[0].memberid).toString();    // For semi-legacy compatibility
		myName = new String(myAdminSignIn[0].forename1).toString();    // Set my name in global variable
		myUserId = new String(myAdminSignIn[0].onlinebookingid).toString();
		myPin = new String(myAdminSignIn[0].onlinebookingpin).toString();
		myAccess = new String(myAdminSignIn[0].webaccess).toString();

		// Now need to check the access level to see if they are admins
		var dummy = new Number(myAccess).value();

		if (debugthis) {
			console.log("Value of adminsignedin inside setAdminSignedIn is ["+adminsignedin+"]");
			console.log("Value of myAdminSignIn inside setAdminSignedIn is ["+JSON.stringify(myAdminSignIn)+"]");
			console.log("Values of global variables inside setAdminSignedIn");
			console.log("myMemberID is ["+myMemberID+"]");
			console.log("myName is ["+myName+"]");
			console.log("myUserId is ["+myUserId+"]");
			console.log("myPin is ["+myPin+"]");
			console.log("myAccess is ["+myAccess+"]");
		}
				
	}  
	else {

	} 

	// showAdminMenus();    // Make sure and set the bottom menus

}


// ------------------------------------------------------------------
// Function to make sure the correct menu is displayed in the footer

function showAdminMenus() {

	if (debugthis) {
		console.log("Value of adminsignedin is ["+adminsignedin+"]");
		console.log("Value of myip is ["+myip+"]");
	}

	if (adminsignedin == "N") {
		if (debugthis)
			console.log("I have NOT been found to be signed-in");
		$('.adminsignedin').addClass('noshow');   // hide signed-in menu at first
		$('.notadminsignedin').removeClass('noshow');  // hide the Sign-in menu choice
	} 
	if (adminsignedin == "Y") {
		if (debugthis)
			console.log("I AM signed-in");
		$('.notadminsignedin').addClass('noshow');
		$('.adminsignedin').removeClass('noshow');
	}

}

// ------------------------------------------------------------------
// Function to try and log me in to admin side
function logAdminIn() {

	// Take sinfo from the loginModal and signs you in

    var userid = $('#frmUserId').val();
    var pin = $('#frmPassword').val();
    var refurl = $('#refurl').attr("value");
    var adminSigninData = new Object();
    var signinurl = new String("").toString();

	var url = "http://hamptontennis.org.uk/fetchJSON.asp?id=11&p1="+userid+"&p2="+pin;
	var goodlogin = "N";

	// Re-initialise global variables on new login
	adminsignedin = new String("N").toString();  // define globally so can use the $.Deferred() jQuery construct
	myMemberID = new String("").toString();         // set to memberid of petrson who has signed in
	myAdminSignIn = new Object();      // Holds details from the localStorage object
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

			// Successful sign-in - now see if they are administrators
			// the webaccess field MUST be greater than ?
			// 
			myUserId = new String(userid).toString();
			myPin = new String(pin).toString();
			myName = new String(allData[0].forename1).toString();
			myMemberID = new String(allData[0].memberid);
			myAccess = new String(allData[0].webaccess);

			if (debugthis) {
				console.log("Value of adminsignedin inside logMeIn is ["+adminsignedin+"]");
				console.log("Value of myAdminSignIn inside logMeIn is ["+JSON.stringify(myAdminSignIn)+"]");
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
			adminSigninData = {
				name: myName,
				pin: myPin,
				userid: myUserId,
				memberid: myMemberID,
				accesslevel: myAccess
			};

			// if (debugthis) {
			//	console.log("Value of memberdata object:   "+jsonstring);
			// }

			// Now, set localStorage up
			// if (debugthis) {
			//	console.log("About to set up data into myAdminSignIn object");
			// }

			// myAdminSignIn.uniqueref = new String(memberdata.uniqueref).toString();
			// myAdminSignIn.onlinebookingid = new String(memberdata.onlinebookingid).toString();
			// myAdminSignIn.onlinebookingpin = new String(memberdata.onlinebookingpin).toString();
			// myAdminSignIn.memberid = new String(memberdata.memberid).toString();
			// myAdminSignIn.forename = new String(memberdata.forename1).toString();
			// myAdminSignIn.acclevel = new String(memberdata.webaccess).toString();

			if (debugthis) {
				// alert("Value of forename1, onlinebookingid and onlinebookingpin = ["+myName+","+myUserId+","+myPin+"]");
				console.log("Value of local adminSigninData object after updates: "+JSON.stringify(adminSigninData));
			}
		
			sessionStorage.setItem("adminsignin", JSON.stringify(adminSigninData));

			// Complete signin process using server-side ASP page
			// signinurl = new String("http://hamptontennis.org.uk/adminsignin.asp?u="+userid+"&p="+pin+"&d="+refurl).toString();

			// if (debugthis) {
			//	console.log("Signin server-side URL = "+signinurl);
			// }
			
			window.location.href = "/admin/index.html";
			// $('#adminsignin').submit();   // Try this way - refurl not being picked up
			
		}

	});  // end of function(allData)	  	

}    // end of logMeIn

// ------------------------------------------------------------------
// Function to log me out

function logAdminOut() {
	
	var signouturl = new String("").toString();

	// Remove the data from localstorage
	sessionStorage.removeItem("adminsignin");

	// Reset global variables
	myMemberID = new String("").toString();    // For semi-legacy compatibility
	myName = new String("").toString();    // Set my name in global variable

	// Complete signout process using server-side ASP page
	signouturl = new String("http://hamptontennis.org.uk/adminsignout.asp").toString();

	if (debugthis) {
		console.log("Sign-out server-side admin URL = "+signouturl);
	}
	
	// N.B.  Server side will do redirect afterwards

	window.location.href = signouturl;

}    // end of logMeOut()

// ------------------------------------------------------------------
// Function to route me appropriately if signed-in

function routeAdmin(destinationurl) {

	var myurl = destinationurl || "/admin/";

	if (debugthis) {
		console.log("Value of adminsignedin in routeAdmin is ["+adminsignedin+"]");
		console.log("Value of myip in routeAdmin is ["+myip+"]");
		console.log("Value of destinationurl in routeAdmin is ["+destinationurl+"]");
		console.log("Value of myurl in routeAdmin is ["+myurl+"]");
	}

	if (adminsignedin == "N") {
		if (debugthis)
			console.log("Routing to login page ...");
		window.location.href = "/admin/adminlogin.html";
	} 
	if (adminsignedin == "Y") {
		if (debugthis)
			console.log("Routing to "+myurl);
		window.location.href = myurl;
	}
}    // end of routeAdmin(url) 

// 
// Now kick all this off
//


$(document).ready(function() {

	// Blank out welcome block to start off with
	// $('.welcomeblock').css('display','none');

	setAdminSignedIn();   // This should reset the admin signed-in status and global variables

	if (! adminsignedin) {
		// Go back to admin login screen - cant do admin unless signed in OK
		window.location.href = "./adminlogin.html";
	}
    
	// Set option text on home page banner image
/*
	if (adminsignedin == "Y") {
		optiontext = new String("Members").toString();
	} else {
		optiontext = new String("Sign-In").toString();
	}
*/
	if (debugthis) {
		console.log("Admin signed-in status = "+adminsignedin);
		// console.log("optiontext = "+optiontext);
	}

	// $('#memberslink').text(optiontext);

	// Now construct welcome message, set HTML and show the welcome block
/*
	if (adminsignedin == "Y") {
		var welcomemessage = "Hi "+myName+" - welcome back";
		$('#welcometext').html(welcomemessage);
		$('#notme').html("Not "+myName+"?");
		$('.welcomeblock').css('display','');
	}
*/
	$( '#loginmenu' ).click(function() {
		$('#frmUserId').focus();
	});

	$('#loginSubmit').click(function(){
		logAdminIn();
	});

	$('#logout').click(function(){
		logAdminOut();
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



