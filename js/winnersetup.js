//
//    fixturesetup.js
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
function displayWinnerSetup() {

	var url = "http://hamptontennis.org.uk/admin/fetchJSON.asp?id=19";

	debugWrite("URL = "+url);

	// var eventsfound = false;
	$.getJSON(url,function(data){

		// console.log(url);

		var jsonstring = JSON.stringify(data);

		// Add name on front if missing
		jsonstring = new String("{allWinners:"+jsonstring+"}");

		// var eventdata = $.parseJSON(jsonstring);
		var winnersdata = eval("(" + jsonstring + ")");

		// Now, need to make sure that we have 30 items in allWinners

		var lengthactual =  winnersdata.allWinners.length;

		// Set the boolean if we have data
		// if (eventdata.length > 1)
		//	eventsfound = true;

		//Get the HTML from the template   in the script tag
	    var theTemplateScript = $("#winnersetup-template").html(); 

	   //Compile the template
	    var theTemplate = Handlebars.compile (theTemplateScript); 
		// Handlebars.registerPartial("description", $("#shoe-description").html());    
		$("#main").append (theTemplate(winnersdata)); 

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



