//
//		admin.js
//
//		Sammy.js and Handlebars code to manage aspects of maintaining the list of coaches
//


		var app = $.sammy(function() {

			// include the plugin and alias handlebars() to hb()
			this.use('Handlebars','hb');

		    this.element_selector = '#thedetails';
		     
		    // Define Sammy routes 
		    
			this.get('#/hello/:name/to/:friend', function(context) {

				//fetch handlebars partial first
				this.load('mypartial.hb')
					.then( function(partial) {
						//set local vars
						context.partials = {hello_friend:partial};
						context.name = context.params.name;
						context.friend = context.params.friend;

						// render the template and pass it through Handlebars
						context.partial('mytemplate.hb');
					} );

			});
		 

			//  End of Sammy routes

		}); 


// ------------------------------------------------------------------
// Now use the utility functions in the document.ready area 

$(document).ready(function() {

	// Call signed-in check function and use .done() for next steps
	setSignedIn();

	app.run();

})    <!--    /document.ready    -->
	
