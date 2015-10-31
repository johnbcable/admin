$(document).ready(function() {
	
	var jsonstring = new String("");
	var url = "http://hamptontennis.org.uk/juniors/fetchJSON.asp?id=6";
	// var eventsfound = false;
	$.getJSON(url,function(data){

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
		$("#coachlist").append (theTemplate(coachdata)); 
		// jsonstring = JSON.stringify(data);
		// $("#receivedjson").html('JSON received back from fetchJSON.asp?id=2 is <br /><br />'+jsonstring+'<br /><hr />');
	});  // end of function(data)
})  // end of document.ready

