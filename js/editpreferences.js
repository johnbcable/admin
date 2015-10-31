
// Register Handlebars helpers

Handlebars.registerHelper('equalsTo', function(v1, v2, options) {
    if(v1 == v2) { return options.fn(this); }
    else { return options.inverse(this); }
});


$(document).ready(function() {
	
	var debugthis = true;    // Output debug messages to console

	var jsonstring = new String("");

	jsonstring = localStorage.getItem("tennissignin");

	if (debugthis) {
		console.log("Raw stringified JSON = ["+jsonstring+"]");
	}

	jsonstring = new String("{myPreferences:"+jsonstring+"}");

	// var eventdata = $.parseJSON(jsonstring);
	var preferencedata = eval("(" + jsonstring + ")");

	// Set the boolean if we have data
	// if (eventdata.length > 1)
	//	eventsfound = true;

	//Get the HTML from the template   in the script tag
    var theTemplateScript = $("#mypreferences-template").html(); 

    //Compile the template
    var theTemplate = Handlebars.compile (theTemplateScript); 
	// Handlebars.registerPartial("description", $("#shoe-description").html());    
	$("#mypreferences").append (theTemplate(preferencedata)); 
	// jsonstring = JSON.stringify(data);
	if (debugthis) {
		$("#receivedjson").html('JSON received back from fetchJSON.asp?id=2 is <br /><br />'+jsonstring+'<br /><hr />');
	}

	// Now, if the section has been marked as completed, disable inputs
	// and show closure message

	$(".completed input").attr("disabled","true");
	$(".completed p.closedmessage").removeClass("noshow");

})  // end of document.ready

