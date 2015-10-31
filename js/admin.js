//
//		admin.js
//
//		Using, Sammy, Handlebars and plugin code 
//

// Utility functions

// Register Handlebars helpers

Handlebars.registerHelper('equalsTo', function(v1, v2, options) { 
    if(v1 == v2) { return options.fn(this); } 
    else { return options.inverse(this); } 
});

// Display coach list
function displayCoachList() {

	var url = "http://hamptontennis.org.uk/admin/fetchJSON.asp?id=0";
	// var eventsfound = false;
	$.getJSON(url,function(data){

		// console.log(url);

		var jsonstring = JSON.stringify(data);

		jsonstring = new String("{allCoaches:"+jsonstring+"}");

		// var eventdata = $.parseJSON(jsonstring);
		var coachdata = eval("(" + jsonstring + ")");

		// Set the boolean if we have data
		// if (eventdata.length > 1)
		//	eventsfound = true;

		//Get the HTML from the template   in the script tag
	    var theTemplateScript = $("#coachlist-template").html(); 

	   //Compile the template
	    var theTemplate = Handlebars.compile (theTemplateScript); 
		// Handlebars.registerPartial("description", $("#shoe-description").html());    
		$("#main").append (theTemplate(coachdata)); 

	});  // end of function(data)

}

// Edit individual coach details
function editCoach(id) {

	$("#main").append ("<p>Editing coach with an ID of "+id+"</p>"); 

	var url = "http://hamptontennis.org.uk/admin/fetchJSON.asp?id=1&p1="+id;

	// var eventsfound = false;
	$.getJSON(url,function(data){

		// console.log(url);

		var jsonstring = JSON.stringify(data);

		jsonstring = new String("{allCoaches:"+jsonstring+"}");

		// var eventdata = $.parseJSON(jsonstring);
		var coachdata = eval("(" + jsonstring + ")");

		// Set the boolean if we have data
		// if (eventdata.length > 1)
		//	eventsfound = true;

		//Get the HTML from the template   in the script tag
	    var theTemplateScript = $("#coachedit-template").html(); 

	   //Compile the template
	    var theTemplate = Handlebars.compile (theTemplateScript); 
		// Handlebars.registerPartial("description", $("#shoe-description").html());    
		$("#main").append (theTemplate(coachdata)); 

	});  // end of function(data)

}

// Display future event list
function displayEventList() {

	var url = "http://hamptontennis.org.uk/admin/fetchJSON.asp?id=2";

	console.log("Inside displayEventList");

	// var eventsfound = false;
	$.getJSON(url,function(data){

		// console.log(url);

		var jsonstring = JSON.stringify(data);

		jsonstring = new String("{allEvents:"+jsonstring+"}");

		// var eventdata = $.parseJSON(jsonstring);
		var eventdata = eval("(" + jsonstring + ")");

		// Set the boolean if we have data
		// if (eventdata.length > 1)
		//	eventsfound = true;

		//Get the HTML from the template   in the script tag
	    var theTemplateScript = $("#eventlist-template").html(); 

	   //Compile the template
	    var theTemplate = Handlebars.compile (theTemplateScript); 
		// Handlebars.registerPartial("description", $("#shoe-description").html());    
		$("#main").append (theTemplate(eventdata)); 

	});  // end of function(data)

}

// Edit selected event
function editEvent(id) {

	var url = "http://hamptontennis.org.uk/admin/fetchJSON.asp?id=3&p1="+id;

	console.log("Inside editEvent");

	// var eventsfound = false;
	$.getJSON(url,function(data){

		// console.log(url);

		var jsonstring = JSON.stringify(data);

		jsonstring = new String("{allEvents:"+jsonstring+"}");

		// var eventdata = $.parseJSON(jsonstring);
		var eventdata = eval("(" + jsonstring + ")");

		// Set the boolean if we have data
		// if (eventdata.length > 1)
		//	eventsfound = true;

		//Get the HTML from the template   in the script tag
	    var theTemplateScript = $("#eventedit-template").html(); 

	   //Compile the template
	    var theTemplate = Handlebars.compile (theTemplateScript); 
		// Handlebars.registerPartial("description", $("#shoe-description").html());    
		$("#main").append (theTemplate(eventdata)); 

		$('.datepicker').pickadate();

	});  // end of function(data)

}

// Display future tournament list
function displayTournamentList() {

	var url = "http://hamptontennis.org.uk/admin/fetchJSON.asp?id=8";

	console.log("Inside displayTournamentList");

	// var eventsfound = false;
	$.getJSON(url,function(data){

		// console.log(url);

		var jsonstring = JSON.stringify(data);

		jsonstring = new String("{allTournaments:"+jsonstring+"}");

		// var eventdata = $.parseJSON(jsonstring);
		var tournamentdata = eval("(" + jsonstring + ")");

		// Set the boolean if we have data
		// if (eventdata.length > 1)
		//	eventsfound = true;

		//Get the HTML from the template   in the script tag
	    var theTemplateScript = $("#tournamentlist-template").html(); 

	   //Compile the template
	    var theTemplate = Handlebars.compile (theTemplateScript); 
		// Handlebars.registerPartial("description", $("#shoe-description").html());    
		$("#main").append (theTemplate(tournamentdata)); 

	});  // end of function(data)

}

function editTournament(id) {

	// $("#main").append ("<p>Editing tournament with an ID of "+id+"</p>"); 

	var url = "http://hamptontennis.org.uk/admin/fetchJSON.asp?id=5&p1="+id;

	// var eventsfound = false;
	$.getJSON(url,function(data){

		// console.log(url);

		var jsonstring = JSON.stringify(data);

		jsonstring = new String("{allTournaments:"+jsonstring+"}");

		// var eventdata = $.parseJSON(jsonstring);
		var tournamentdata = eval("(" + jsonstring + ")");

		// Set the boolean if we have data
		// if (eventdata.length > 1)
		//	eventsfound = true;

		//Get the HTML from the template   in the script tag
	    var theTemplateScript = $("#tournamentedit-template").html(); 

	   //Compile the template
	    var theTemplate = Handlebars.compile (theTemplateScript); 
		// Handlebars.registerPartial("description", $("#shoe-description").html());    
					$("#main").append (theTemplate(tournamentdata)); 

		$('.datepicker').pickadate( {
			selectMonths: true,
			selectYears: true,
			format: 'dd/mm/yyyy',
			formatSubmit: 'dd/mm/yyyy',
			editable: true
		});

		$('.timepicker').pickatime( {
			format: 'HHi',
			formatSubmit: 'HHi',
			editable: true,
		    min: [9,0],
    		max: [19,0]
		});

	});  // end of function(data)

}

// --------------  End of helper/utility functions ---------------------------

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

	// Coaches home page   -----------------------------

	this.get('#/coaches', function(context) { 
		context.app.swap('');   // clears HTML content
		// Redisplay coaches home page
		displayCoachList();


	});   // end get

	// Display single coach page for edit
	this.get('#/coaches/:id', function(context) {
		context.app.swap('');   // clears HTML content
		// this.item = this.items[context.params['id']];

		// Now edit this coach details
		editCoach(context.params['id']);

	});    // end get

	// -------------  End of Coaches  -------------------

	// Events home page  -------------------------------

	this.get('#/events', function(context) { 
		context.app.swap('');   // clears HTML content
		// Redisplay coaches home page
		displayEventList();


	});   // end get

	// Display single event page for edit
	this.get('#/events/:id', function(context) {
		context.app.swap('');   // clears HTML content
		// this.item = this.items[context.params['id']];

		// Now edit this coach details
		editEvent(context.params['id']);

	});    // end get

	// -------------  End of Events  -------------------

	// Tournaments home page   -------------------------

	this.get('#/tournaments', function(context) { 

		context.app.swap('');   // clears HTML content

		// Redisplay coaches home page
		displayTournamentList();

	});   // end get

	// Display single tournament page for edit
	this.get('#/tournaments/:id', function(context) {

		context.app.swap('');   // clears HTML content
		// this.item = this.items[context.params['id']];

		// Now edit this coach details
		editTournament(context.params['id']);

	});    // end get

	// -------------  End of Tournaments  -------------------

	// -------------  Quick Message  ------------------------

	this.get('#/quickmessage', function(context) {

		context.app.swap('<p>Inside quick message area</p>');   // clears HTML content
		// this.item = this.items[context.params['id']];

		// Now edit this coach details
		//editTournament(context.params['id']);

	});    // end get

	// -------------  End of Quick Message  -----------------


});


	// End of route definition

$(function() { 

	// Set signed-in status
	setSignedIn();

	// Now run the main Sammy route
	app.run('#/');
}); 

})(jQuery);


