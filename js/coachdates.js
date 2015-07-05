

function displayCoachSetupDates() {

	var url = "http://hamptontennis.org.uk/fetchJSON.asp?id=18";
	// var eventsfound = false;
	$.getJSON(url,function(data){

		// console.log(url);

		var jsonstring = JSON.stringify(data);

		jsonstring = new String("{allCoachDates:"+jsonstring+"}");

		// var eventdata = $.parseJSON(jsonstring);
		var coachdatesdata = eval("(" + jsonstring + ")");

		// Set the boolean if we have data
		// if (eventdata.length > 1)
		//	eventsfound = true;

		//Get the HTML from the template   in the script tag
	    var theTemplateScript = $("#coachdates-template").html(); 

	   //Compile the template
	    var theTemplate = Handlebars.compile (theTemplateScript); 
		// Handlebars.registerPartial("description", $("#shoe-description").html());    
		$("#main").append (theTemplate(coachdatesdata)); 

		// Print raw received JSON on the page  
		$("#receivedjson").append(jsonstring);

	});  // end of function(data)

}


$(document).ready(function() {
	
	displayCoachSetupDates();

})  // end of document.ready

