$(document).ready(function() {
	
	var debugthis = true;    // Ouptu debug messages to console

	var jsonstring = new String("");

	jsonstring = localStorage.getItem("tennissignin");

	if (debugthis) {
		console.log("Raw stringified JSON = ["+jsonstring+"]");
	}

	jsonstring = new String("{myProfile:"+jsonstring+"}");

	// var eventdata = $.parseJSON(jsonstring);
	var profiledata = eval("(" + jsonstring + ")");

	// Set the boolean if we have data
	// if (eventdata.length > 1)
	//	eventsfound = true;

	//Get the HTML from the template   in the script tag
    var theTemplateScript = $("#myprofile-template").html(); 

    //Compile the template
    var theTemplate = Handlebars.compile (theTemplateScript); 
	// Handlebars.registerPartial("description", $("#shoe-description").html());    
	$("#mydetails").append (theTemplate(profiledata)); 
	// jsonstring = JSON.stringify(data);
	if (debugthis) {
		$("#receivedjson").html('JSON received back from fetchJSON.asp?id=2 is <br /><br />'+jsonstring+'<br /><hr />');
	}

})  // end of document.ready

