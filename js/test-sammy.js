//
//		test-sammy.js
//
//		Sammy.js and plugin code to manage aspects of maintaining the list of coaches
//


(function($) {

	var app = $.sammy('#main', function() {

	// this.element_selector = '#main';

	// Use Sammy.template
	this.use('Template');
	// Use Handlebars
	// this.use('Handlebars','hb');
	// Use Sammy sessions
	this.use('Session');

	// Utility functions
	this.around(function(callback) {
		var context = this;
		// Reload coach list data
   		this.load('http://hamptontennis.org.uk/admin/fetchJSON.asp?id=0')
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
		$.each(this.items, function(i, item) {
			context.render('templates/item.template', {id: i, item: item}) 
				.appendTo(context.$element());
		}); 
	});   // end get

	// Itemdetail page
	this.get('#/item/:id', function(context) {
		this.item = this.items[context.params['id']];
		if (!this.item) {
			return this.notFound();
		}
		this.partial('templates/item-detail.template');

	});    // end get

	// Add to shopping cart (POST)
	this.post('#/cart', function(context) { 
		// context.log("I'm in a post route. Add me to your cart");
		var item_id = context.params['item_id'];
		// fetch the current cart
		var cart = this.session('cart', function() {
			return {};
		});
		if (! cart[item_id]) {
			// This item is not yet in the cart
			// Initialize its quantity with zero
			cart[item_id] = 0;
		}
		cart[item_id] += parseInt(context.params['quantity'], 10);
		// store the cart
		this.session('cart', cart);
		context.log("The current cart: ", cart);
		this.trigger('update-cart');
	});

	// Add in a bind to update cart totals
	this.bind('update-cart', function() {
		var sum = 0;
		$.each(this.session('cart') || {}, function(id, quantity) { 
			sum += quantity;
		}); 
		$('.cart-info')
			.find('.cart-items').text(sum).end() 
			.animate({paddingTop: '30px'}) 
			.animate({paddingTop: '10px'});
	});

	// Update the cart on initial run
	this.bind('run', function() {
		// initialize the cart display
		this.trigger('update-cart'); 
	});


});


	// End of route definition

$(function() { 

	// Set signed-in status
	setSignedIn();

	// Now run the main Sammy route
	app.run('#/');
}); 

})(jQuery);


