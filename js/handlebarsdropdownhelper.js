Handlebars.registerHelper('equalsTo', function(v1, v2, options) { 
    if(v1 == v2) { return options.fn(this); } 
    else { return options.inverse(this); } 
});

var priorityType = "medium";
var data = {priorityType: priorityType}; 

var source= " <select class=\"selectPriortyCl\"> " +
                "<option value=\"High\" {{#equalsTo priorityType \"high\"}}selected{{/equalsTo}}>High</option>"+
                "<option value=\"Medium\" {{#equalsTo priorityType \"medium\"}}selected{{/equalsTo}}>Medium</option>"+
                "<option value=\"Low\" {{#equalsTo priorityType \"low\"}}selected{{/equalsTo}}>Low</option>"+
                "<option value=\"None\" {{#equalsTo priorityType \"none\"}}selected{{/equalsTo}}>None</option>"+
              "</select>";

var template = Handlebars.compile(source);
alert(template(data));
