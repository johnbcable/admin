

Handlebars.getTemplate = function(name) {
	if (Handlebars.templates === undefined || Handlebars.templates[name] === undefined) { 
		$.ajax({
			url : 'handlebars/' + name + '.handlebars',
			success: function(data) {
				if (Handlebars.templates === undefined) {
					Handlebars.templates = {}; 
				}
				Handlebars.templates[name] = Handlebars.compile(data);
			},
			async: false
		}); 
	}
	return Handlebars.templates[name];
};

(function() {
  var template = Handlebars.template, templates = Handlebars.templates = Handlebars.templates || {};
templates['coachlist'] = template({"1":function(depth0,helpers,partials,data) {
  var helper, functionType="function", escapeExpression=this.escapeExpression;
  return "\n          <tr>\n            <td>"
    + escapeExpression(((helper = helpers.forename1 || (depth0 && depth0.forename1)),(typeof helper === functionType ? helper.call(depth0, {"name":"forename1","hash":{},"data":data}) : helper)))
    + " "
    + escapeExpression(((helper = helpers.surname || (depth0 && depth0.surname)),(typeof helper === functionType ? helper.call(depth0, {"name":"surname","hash":{},"data":data}) : helper)))
    + "</td>\n            <td>\n              Hourly: &pound;"
    + escapeExpression(((helper = helpers.hourlyrate || (depth0 && depth0.hourlyrate)),(typeof helper === functionType ? helper.call(depth0, {"name":"hourlyrate","hash":{},"data":data}) : helper)))
    + "<br />\n              Half-hourly: &pound;"
    + escapeExpression(((helper = helpers.halfhourlyrate || (depth0 && depth0.halfhourlyrate)),(typeof helper === functionType ? helper.call(depth0, {"name":"halfhourlyrate","hash":{},"data":data}) : helper)))
    + "\n            </td>\n            <td>\n              <a href=\"/admin/#/coaches/"
    + escapeExpression(((helper = helpers.uniqueref || (depth0 && depth0.uniqueref)),(typeof helper === functionType ? helper.call(depth0, {"name":"uniqueref","hash":{},"data":data}) : helper)))
    + "\" class=\"small button coachedit\" data-name=\""
    + escapeExpression(((helper = helpers.surname || (depth0 && depth0.surname)),(typeof helper === functionType ? helper.call(depth0, {"name":"surname","hash":{},"data":data}) : helper)))
    + "\">Edit</a>\n            </td>\n          </tr>\n          ";
},"compiler":[5,">= 2.0.0"],"main":function(depth0,helpers,partials,data) {
  var stack1, buffer = "      <h4>Current Coaching Staff</h4>\n      <table width=\"100%\">\n        <thead>\n          <tr>\n            <th>Name</th>\n            <th>Charge rates</th>\n            <th>Action</th>\n          </tr>\n        </thead>\n        <tbody>\n          ";
  stack1 = helpers.each.call(depth0, (depth0 && depth0.allCoaches), {"name":"each","hash":{},"fn":this.program(1, data),"inverse":this.noop,"data":data});
  if(stack1 || stack1 === 0) { buffer += stack1; }
  return buffer + "\n        </tbody>\n      </table>\n\n";
},"useData":true});
})();
