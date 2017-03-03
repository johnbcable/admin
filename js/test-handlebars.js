//
//		test-sammy.js
//
//		Sammy.js and plugin code 
//


(function($) {

	var app = $.sammy('#main', function() {

	// this.element_selector = '#main';

	// Use Sammy.template
	// this.use('Template');
	// Use Handlebars
	this.use('Handlebars','hb');
	// Use Sammy sessions
	this.use('Session');

	// Utility functions
	this.around(function(callback) {
		var context = this;
		// Reload coach list data
   		this.load('https://hamptontennis.org.uk/admin/fetchJSON.asp?id=0')
  			.then(function(items) {
        		$.each(this.json(items), function(i, item) {
          			context.log(item);
        		});  // end $.each

    		});  // end function(items)
			.then(callback);
		});

	// Define all the required routes


	// Home or start page
	this.get('#/', function(context) { 
		context.app.swap('');   // clears HTML content
		// Redisplay admin home page - blank with buttons


	});   // end get

	// Coaches home page
	this.get('#/coaches', function(context) { 
		context.app.swap('');   // clears HTML content
		// Redisplay coaches home page
		displayCoachList();


	});   // end get

	// Display single coach page for edit
	this.get('#/coaches/:id', function(context) {
		this.item = this.items[context.params['id']];
		if (!this.item) {
			return this.notFound();
		}


		// this.partial('templates/item-detail.template');

	});    // end get


});


	// End of route definition

$(function() { 

	// Set signed-in status
	setSignedIn();

	// Now run the main Sammy route
	app.run('#/');
}); 

})(jQuery);


