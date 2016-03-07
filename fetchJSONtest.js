
$(document).ready(function() {
	
	var jsonstring = new String("");

	$('#mysubmit').click( function () {

		var $url = new String("http://hamptontennis.org.uk/fetchJSON.asp?id=&thequery=");
		var $str = $('#thequery').val();

		$url += encodeURIComponent($str);
		
		alert($url);

		$.getJSON($url,function(data){

			jsonstring = JSON.stringify(data);
			$("#receivedjson").html();
			$("#receivedjson").html('JSON received back from fetchJSON.asp is <br /><br />'+jsonstring+'<br /><hr />');
		});

	});

})
