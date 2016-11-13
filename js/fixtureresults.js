//
//    fixtureresults.js
//

var debugthis = false;    // Production = false

// Utility functions

// Register Handlebars helpers

Handlebars.registerHelper('equalsTo', function(v1, v2, options) { 
    if(v1 == v2) { return options.fn(this); } 
    else { return options.inverse(this); } 
});

// Send debug message to the console log
function debugWrite(message) {
	if (debugthis) {
		console.log(message);
	}
}

// Display this seasons fixtures
function displayFixtures(gender,team) {

	var url = "http://hamptontennis.org.uk/admin/fetchJSON.asp?id=18";
	var offset = (gender == "Ladies" ? 0 : 3);
	var myindex = team + offset;
	var teamnames = ["",
		"Ladies 1st Team",
		"Ladies 2nd Team",
		"Ladies 3rd Team",
		"Mens 1st Team",
		"Mens 2nd Team",
		"Mens 3rd Team",
		"Mens 4th Team",
		"Mens 5th Team"];
	var today = new Date();
	var year = today.getFullYear();
	var myteam = new String(teamnames[myindex]).toString();

	debugWrite("gender = "+gender+", team="+team+", myindex="+myindex+", teamname="+myteam);

	if (debugthis) {
		year = "2016";
	}
	
	url += "&p1="+year+"&p2="+myteam;
	
	debugWrite("URL = "+url);

	// Now, set display text at top of the screen area
	var displaytext = year+" fixtures for the "+myteam;
	$('.fixturetitle').html(displaytext);

	// var eventsfound = false;
	$.getJSON(url,function(data){

		// console.log(url);

		var jsonstring = JSON.stringify(data);

		// Add name on front if missing
		jsonstring = new String("{allFixtures:"+jsonstring+"}");

		// var eventdata = $.parseJSON(jsonstring);
		var fixturedata = eval("(" + jsonstring + ")");

		// Now, need to make sure that we have 14 items in allFixtures

		var lengthactual =  fixturedata.allFixtures.length;
		if (lengthactual < 14) {   // 14 is the default number of matches in each division
			// Define skeleton default zero content object to put in fixturedata

			var dummy = {
				"fixturedate":null,
				"fixtureyear":year,
				"teamname":myteam,
				"homeoraway":"H",
				"opponents":"NONE",
				"fixturenote":"",
				"hamptonresult":0,
				"opponentresult":0,
				"matchreport":"",
				"pair1":"",
				"pair2":"",
				"fixtureid": 0
			}
			for (var i=lengthactual; i<14; i++) {
				dummy.fixtureid = (9999-i);
				fixturedata.allFixtures.push(dummy);
			}
			lengthactual = fixturedata.allFixtures.length;
			for (var i=0; i<14; i++) {
				fixturedata.allFixtures[i].teamname = myteam;
				fixturedata.allFixtures[i].fixtureyear = year;
			}

		}
		// Set the boolean if we have data
		// if (eventdata.length > 1)
		//	eventsfound = true;

		//Get the HTML from the template   in the script tag
	    var theTemplateScript = $("#fixturesetup-template").html(); 

	   //Compile the template
	    var theTemplate = Handlebars.compile (theTemplateScript); 
		// Handlebars.registerPartial("description", $("#shoe-description").html());    
		$("#main").append (theTemplate(fixturedata)); 

		// Output raw JSON back to page
		$("#receivedjson").html(jsonstring);

	});  // end of function(data)

}


// Main Sammy area
(function($) {

	// Set element main as where the action will be
	var app = $.sammy('#main', function() {

	// this.element_selector = '#main';

	// Define all the required routes

	// Home or start page   ----------------------------

	this.get('#/', function(context) { 
		// context.app.swap('');   // clears HTML content
		// Redisplay admin home page - blank with buttons
		context.app.swap('');

	});   // end get

	// Ladies fixture setup area   -----------------------------

	this.get('#/fixtures/ladies/1', function(context) { 
		context.app.swap('');   // clears HTML content
		// Redisplay coaches home page
		displayFixtures("Ladies",1);


	});   // end get

	this.get('#/fixtures/ladies/2', function(context) { 
		context.app.swap('');   // clears HTML content
		// Redisplay coaches home page
		displayFixtures("Ladies",2);


	});   // end get

	this.get('#/fixtures/ladies/3', function(context) { 
		context.app.swap('');   // clears HTML content
		// Redisplay coaches home page
		displayFixtures("Ladies",3);


	});   // end get

	// end of Ladies fixture setup area   -----------------------------

	// Mens fixture setup area   -----------------------------

	this.get('#/fixtures/mens/1', function(context) { 
		context.app.swap('');   // clears HTML content
		// Redisplay coaches home page
		displayFixtures("Mens",1);


	});   // end get

	this.get('#/fixtures/mens/2', function(context) { 
		context.app.swap('');   // clears HTML content
		// Redisplay coaches home page
		displayFixtures("Mens",2);


	});   // end get

	this.get('#/fixtures/mens/3', function(context) { 
		context.app.swap('');   // clears HTML content
		// Redisplay coaches home page
		displayFixtures("Mens",3);


	});   // end get

	this.get('#/fixtures/mens/4', function(context) { 
		context.app.swap('');   // clears HTML content
		// Redisplay coaches home page
		displayFixtures("Mens",4);


	});   // end get

	this.get('#/fixtures/mens/5', function(context) { 
		context.app.swap('');   // clears HTML content
		// Redisplay coaches home page
		displayFixtures("Mens",5);


	});   // end get


});


	// End of route definition

$(function() { 

	// Now run the main Sammy route
	app.run('#/');
}); 

})(jQuery);



