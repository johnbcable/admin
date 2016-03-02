<%
//
//  CoachingLessonObject
//
//  Holds detail for a single coaching lesson 
//

function CoachingLessonObject(lessonid,day,start,finish,group) {

	this.schedule_role = new String("NEXT").toString();
	this.lesson_day = new String(day).toString() || new String("").toString();
	this.start_time = new String(start).toString() || new String("").toString();
	this.end_time = new String(finish).toString() || new String("").toString();
	this.group_code = new String(group).toString() || new String("").toString();
	this.lesson_text = new String("").toString();
	this.lesson_id = lesson_id || 1;
	this.group_display_order = 99;
};

CoachingLessonObject.prototype.setLessonText = function(lessontext) {
	this.lesson_text = lessontext;
};

CoachingLessonObject.prototype.setDisplayOrder = function() {

	var groups = ["","TOTS","RED","ORANGE","GREEN","FULL BALL","UNDER 14","UNDER 16","UNDER 18"]
	for (var i=0; i<groups.length; i++) {
		if (groups[i] == this.group_code) {
			this.group_display_order = i;
		}
	}
};


%>
