//
//  Display a given day's set of coaching lessons and allow editing
//
// Utility functions

// Register Handlebars helpers

Handlebars.registerHelper('equalsTo', function(v1, v2, options) { 
    if(v1 == v2) { return options.fn(this); } 
    else { return options.inverse(this); } 
});

var debugthis = true;   // Set to false for production

function daytext(day) {
	var texts = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];
	var dummy = new String("").toString();
	var retval = new String("").toString();
	for (var i=0; i<texts.length; i++) {
		dummy = new String(texts[i]).toString();
		dummy = dummy.substring(0,3).toUpperCase();
		if (debugthis) {
			console.log("comparing dummy ["+dummy+"] with day ["+day+"]");
		}
		if (dummy == day) {
			if (debugthis) {
				console.log("Found - retval = "+retval);
			}
			retval = new String(texts[i]).toString();
		}
	}
	if (debugthis) {
		console.log("returning "+retval);
	}
	return retval;
}

function displayCoachingLessons(day) {

	var url = "http://hamptontennis.org.uk/fetchJSON.asp?id=19";

	var theday = day || "MON";   // default to Monday

	url += "&p1="+theday

	// var eventsfound = false;
	$.getJSON(url,function(data){

		// console.log(url);

		var jsonstring = JSON.stringify(data);

		jsonstring = new String("{allCoachingLessons:"+jsonstring+"}");

		// var eventdata = $.parseJSON(jsonstring);
		var lessondata = eval("(" + jsonstring + ")");

		if (debugthis) {
			console.log("lessondata = "+JSON.stringify(lessondata));
			console.log("length of lessondata = "+lessondata.allCoachingLessons.length);
		}

		// Now, need to make sure that we have 10 items in allCoachingLessons

		var lengthactual =  lessondata.allCoachingLessons.length;
		if (lengthactual < 10) {
			// Define skeleton default zero content object to put in lessondata
			var dummy = {
				"schedule_role":"NEXT",
				"lesson_day":theday,
				"start_time":"00:00:00",
				"end_time":"00:00:00",
				"lesson_text":"",
				"lesson_coach":null,
				"lesson_id":0,
				"group_code":"-",
				"group_display_order":99
			}
			for (var i=lengthactual; i<10; i++) {
				dummy.lesson_id = (99-i);
				lessondata.allCoachingLessons.push(dummy);
			}
		}

		if (debugthis) {
			console.log("After adjustment, length of lessondata = "+lessondata.allCoachingLessons.length);
			console.log("lessondata = "+JSON.stringify(lessondata));
		}

		// Set the boolean if we have data
		// if (eventdata.length > 1)
		// 	eventsfound = true;

		//Get the HTML from the template   in the script tag
	    var theTemplateScript = $("#coachinglessons-template").html(); 

	   //Compile the template
	    var theTemplate = Handlebars.compile (theTemplateScript); 
		// Handlebars.registerPartial("description", $("#shoe-description").html());    
		$("#main").append (theTemplate(lessondata)); 

		// Print raw received JSON on the page  
		// $("#receivedjson").append(jsonstring);

	});  // end of function(data)

}


$(document).ready(function() {
	
	var day = new String(Request.QueryString("day")).toString();
	day = day.toUpperCase() || "MON";
	var mydaytext = new String("").toString();

	if (debugthis) {
		console.log("day = "+day);
	}

	mydaytext = daytext(day);
	if (debugthis) {
		console.log("mydaytext = "+mydaytext);
	}

	$('.daytitle').html(mydaytext+"&apos;s Coaching Lessons");

	displayCoachingLessons(day);

})  // end of document.ready

