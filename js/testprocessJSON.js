
// testprocessJSON.js

var worked = false;

$(function() {
	var thing = new String("berserk");
	alert('Just before $.ajax call and thing = ['+thing+']');
	$.ajax({
	        type: "POST",
	        url: 'http://hamptontennis.org.uk/admin/processJSON.asp',
	        data: "{'userName':'" + thing + "'}",
	        contentType: "application/json; charset=utf-8",
	        dataType: "json",
	        cache: false,
	      	async: false,
	        success: function() {
				worked = true;
			},
			error: function() {
				worked = false;
			}
	});    // end of $.ajax processing
	alert('After $.ajax call to processJSON.asp');
	if (worked)
		alert("... and it worked!");
	else
		alert("... but it signalled it didnt work :(");

});        

