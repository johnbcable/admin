<%@language="JScript" CODEPAGE="65001"%>
<!--#include file="json2.js.asp" -->
<!--#include file="functions.asp" -->
<!--#include file="strings.asp" -->
<!--#include file="coachinglessonobj.asp" -->
<!--#include file="emailfuncs.asp" -->
<!--#include file="datefuncs.asp" -->
<%

// Set up cache control on this page
Response.AddHeader("Cache-Control", "no-cache,no-store,must-revalidate");
Response.AddHeader("Pragma", "no-cache");
Response.AddHeader("Expires", 0);

var strtime, strdate, title2;
var clubname = new String("Hampton-In-Arden Sports Club");
var pagetitle = new String("Updating Coaching Lesson Schedule");
// Now for any variables local to this page
var RS, Conn, SQL1, SQL2;
var dbconnect=Application("hamptonsportsdb"); 
var SQLstart, SQLmiddle, SQLend;
var resultObj = new Object();
var uniqueref;
var today = new Date();
var lessondetail = new Object(); 
var start_time, end_time, lesson_text, group_code;
var schedule_role, lesson_day, lesson_id;
var lessons = new Array();

var debugging = false;  // Production = false

// Set up DB connections and unchanging bits of SQL
Conn = Server.CreateObject("ADODB.Connection");
RS = Server.CreateObject("ADODB.RecordSet");
Conn.Open(dbconnect);
SQLstart = new String("UPDATE coaching_lessons ")

// Retrieve POST'ed data

// One-offs first

schedule_role = Trim(new String(Request.Form("schedule_role")));
if (schedule_role == "" || schedule_role =="null" || schedule_role == "undefined")
{
	schedule_role = new String("NEXT").toString();
} 
lesson_day = Trim(new String(Request.Form("lesson_day")));
if (lesson_day == "" || lesson_day =="null" || lesson_day == "undefined")
{
	lesson_day = new String("").toString();
} 

// Then a line at a time
// Starts at line 0 - ignore lines with no group code

// Line 0

group_code = Trim(new String(Request.Form("group_code_0")));
if (group_code == "" || group_code =="null" || group_code == "undefined")
{
	group_code = new String("").toString();
} 

start_time = Trim(new String(Request.Form("start_time_0")));
if (start_time == "" || start_time =="null" || start_time == "undefined")
{
	start_time = new String("00:00:00").toString();
} 

end_time = Trim(new String(Request.Form("end_time_0")));
if (end_time == "" || end_time =="null" || end_time == "undefined")
{
	end_time = new String("00:00:00").toString();
} 

lesson_text = Trim(new String(Request.Form("lesson_text_0")));
if (lesson_text == "" || lesson_text =="null" || lesson_text == "undefined")
{
	lesson_text = new String("").toString();
} 

lesson_id = Trim(new String(Request.Form("lesson_id_0")));
if (lesson_id == "" || lesson_id =="null" || lesson_id == "undefined")
{
	lesson_id = new String("0").toString();
} 

lessondetail = new CoachingLessonObject(lesson_id,lesson_day,start_time,end_time,group_code);
lessondetail.setLessonText(lesson_text);
lessondetail.setDisplayOrder();
lessons.push(lessondetail);

// Line 1

group_code = Trim(new String(Request.Form("group_code_1")));
if (group_code == "" || group_code =="null" || group_code == "undefined")
{
	group_code = new String("").toString();
} 

start_time = Trim(new String(Request.Form("start_time_1")));
if (start_time == "" || start_time =="null" || start_time == "undefined")
{
	start_time = new String("00:00:00").toString();
} 

end_time = Trim(new String(Request.Form("end_time_1")));
if (end_time == "" || end_time =="null" || end_time == "undefined")
{
	end_time = new String("00:00:00").toString();
} 

lesson_text = Trim(new String(Request.Form("lesson_text_1")));
if (lesson_text == "" || lesson_text =="null" || lesson_text == "undefined")
{
	lesson_text = new String("").toString();
} 

lesson_id = Trim(new String(Request.Form("lesson_id_1")));
if (lesson_id == "" || lesson_id =="null" || lesson_id == "undefined")
{
	lesson_id = new String("0").toString();
} 

lessondetail = new CoachingLessonObject(lesson_id,lesson_day,start_time,end_time,group_code);
lessondetail.setLessonText(lesson_text);
lessondetail.setDisplayOrder();
lessons.push(lessondetail);

// Line 2

group_code = Trim(new String(Request.Form("group_code_2")));
if (group_code == "" || group_code =="null" || group_code == "undefined")
{
	group_code = new String("").toString();
} 

start_time = Trim(new String(Request.Form("start_time_2")));
if (start_time == "" || start_time =="null" || start_time == "undefined")
{
	start_time = new String("00:00:00").toString();
} 

end_time = Trim(new String(Request.Form("end_time_2")));
if (end_time == "" || end_time =="null" || end_time == "undefined")
{
	end_time = new String("00:00:00").toString();
} 

lesson_text = Trim(new String(Request.Form("lesson_text_2")));
if (lesson_text == "" || lesson_text =="null" || lesson_text == "undefined")
{
	lesson_text = new String("").toString();
} 

lesson_id = Trim(new String(Request.Form("lesson_id_2")));
if (lesson_id == "" || lesson_id =="null" || lesson_id == "undefined")
{
	lesson_id = new String("0").toString();
} 

lessondetail = new CoachingLessonObject(lesson_id,lesson_day,start_time,end_time,group_code);
lessondetail.setLessonText(lesson_text);
lessondetail.setDisplayOrder();
lessons.push(lessondetail);

// Line 3

group_code = Trim(new String(Request.Form("group_code_3")));
if (group_code == "" || group_code =="null" || group_code == "undefined")
{
	group_code = new String("").toString();
} 

start_time = Trim(new String(Request.Form("start_time_3")));
if (start_time == "" || start_time =="null" || start_time == "undefined")
{
	start_time = new String("00:00:00").toString();
} 

end_time = Trim(new String(Request.Form("end_time_3")));
if (end_time == "" || end_time =="null" || end_time == "undefined")
{
	end_time = new String("00:00:00").toString();
} 

lesson_text = Trim(new String(Request.Form("lesson_text_3")));
if (lesson_text == "" || lesson_text =="null" || lesson_text == "undefined")
{
	lesson_text = new String("").toString();
} 

lesson_id = Trim(new String(Request.Form("lesson_id_3")));
if (lesson_id == "" || lesson_id =="null" || lesson_id == "undefined")
{
	lesson_id = new String("0").toString();
} 

lessondetail = new CoachingLessonObject(lesson_id,lesson_day,start_time,end_time,group_code);
lessondetail.setLessonText(lesson_text);
lessondetail.setDisplayOrder();
lessons.push(lessondetail);

// Line 4

group_code = Trim(new String(Request.Form("group_code_4")));
if (group_code == "" || group_code =="null" || group_code == "undefined")
{
	group_code = new String("").toString();
} 

start_time = Trim(new String(Request.Form("start_time_4")));
if (start_time == "" || start_time =="null" || start_time == "undefined")
{
	start_time = new String("00:00:00").toString();
} 

end_time = Trim(new String(Request.Form("end_time_4")));
if (end_time == "" || end_time =="null" || end_time == "undefined")
{
	end_time = new String("00:00:00").toString();
} 

lesson_text = Trim(new String(Request.Form("lesson_text_4")));
if (lesson_text == "" || lesson_text =="null" || lesson_text == "undefined")
{
	lesson_text = new String("").toString();
} 

lesson_id = Trim(new String(Request.Form("lesson_id_4")));
if (lesson_id == "" || lesson_id =="null" || lesson_id == "undefined")
{
	lesson_id = new String("0").toString();
} 

lessondetail = new CoachingLessonObject(lesson_id,lesson_day,start_time,end_time,group_code);
lessondetail.setLessonText(lesson_text);
lessondetail.setDisplayOrder();
lessons.push(lessondetail);

// Line 5

group_code = Trim(new String(Request.Form("group_code_5")));
if (group_code == "" || group_code =="null" || group_code == "undefined")
{
	group_code = new String("").toString();
} 

start_time = Trim(new String(Request.Form("start_time_5")));
if (start_time == "" || start_time =="null" || start_time == "undefined")
{
	start_time = new String("00:00:00").toString();
} 

end_time = Trim(new String(Request.Form("end_time_5")));
if (end_time == "" || end_time =="null" || end_time == "undefined")
{
	end_time = new String("00:00:00").toString();
} 

lesson_text = Trim(new String(Request.Form("lesson_text_5")));
if (lesson_text == "" || lesson_text =="null" || lesson_text == "undefined")
{
	lesson_text = new String("").toString();
} 

lesson_id = Trim(new String(Request.Form("lesson_id_5")));
if (lesson_id == "" || lesson_id =="null" || lesson_id == "undefined")
{
	lesson_id = new String("0").toString();
} 

lessondetail = new CoachingLessonObject(lesson_id,lesson_day,start_time,end_time,group_code);
lessondetail.setLessonText(lesson_text);
lessondetail.setDisplayOrder();
lessons.push(lessondetail);

// Line 6

group_code = Trim(new String(Request.Form("group_code_6")));
if (group_code == "" || group_code =="null" || group_code == "undefined")
{
	group_code = new String("").toString();
} 

start_time = Trim(new String(Request.Form("start_time_6")));
if (start_time == "" || start_time =="null" || start_time == "undefined")
{
	start_time = new String("00:00:00").toString();
} 

end_time = Trim(new String(Request.Form("end_time_6")));
if (end_time == "" || end_time =="null" || end_time == "undefined")
{
	end_time = new String("00:00:00").toString();
} 

lesson_text = Trim(new String(Request.Form("lesson_text_6")));
if (lesson_text == "" || lesson_text =="null" || lesson_text == "undefined")
{
	lesson_text = new String("").toString();
} 

lesson_id = Trim(new String(Request.Form("lesson_id_6")));
if (lesson_id == "" || lesson_id =="null" || lesson_id == "undefined")
{
	lesson_id = new String("0").toString();
} 

lessondetail = new CoachingLessonObject(lesson_id,lesson_day,start_time,end_time,group_code);
lessondetail.setLessonText(lesson_text);
lessondetail.setDisplayOrder();
lessons.push(lessondetail);

// Line 7

group_code = Trim(new String(Request.Form("group_code_7")));
if (group_code == "" || group_code =="null" || group_code == "undefined")
{
	group_code = new String("").toString();
} 

start_time = Trim(new String(Request.Form("start_time_7")));
if (start_time == "" || start_time =="null" || start_time == "undefined")
{
	start_time = new String("00:00:00").toString();
} 

end_time = Trim(new String(Request.Form("end_time_7")));
if (end_time == "" || end_time =="null" || end_time == "undefined")
{
	end_time = new String("00:00:00").toString();
} 

lesson_text = Trim(new String(Request.Form("lesson_text_7")));
if (lesson_text == "" || lesson_text =="null" || lesson_text == "undefined")
{
	lesson_text = new String("").toString();
} 

lesson_id = Trim(new String(Request.Form("lesson_id_7")));
if (lesson_id == "" || lesson_id =="null" || lesson_id == "undefined")
{
	lesson_id = new String("0").toString();
} 

lessondetail = new CoachingLessonObject(lesson_id,lesson_day,start_time,end_time,group_code);
lessondetail.setLessonText(lesson_text);
lessondetail.setDisplayOrder();
lessons.push(lessondetail);

// Line 8

group_code = Trim(new String(Request.Form("group_code_8")));
if (group_code == "" || group_code =="null" || group_code == "undefined")
{
	group_code = new String("").toString();
} 

start_time = Trim(new String(Request.Form("start_time_8")));
if (start_time == "" || start_time =="null" || start_time == "undefined")
{
	start_time = new String("00:00:00").toString();
} 

end_time = Trim(new String(Request.Form("end_time_8")));
if (end_time == "" || end_time =="null" || end_time == "undefined")
{
	end_time = new String("00:00:00").toString();
} 

lesson_text = Trim(new String(Request.Form("lesson_text_8")));
if (lesson_text == "" || lesson_text =="null" || lesson_text == "undefined")
{
	lesson_text = new String("").toString();
} 

lesson_id = Trim(new String(Request.Form("lesson_id_8")));
if (lesson_id == "" || lesson_id =="null" || lesson_id == "undefined")
{
	lesson_id = new String("0").toString();
} 

lessondetail = new CoachingLessonObject(lesson_id,lesson_day,start_time,end_time,group_code);
lessondetail.setLessonText(lesson_text);
lessondetail.setDisplayOrder();
lessons.push(lessondetail);

// Line 9

group_code = Trim(new String(Request.Form("group_code_9")));
if (group_code == "" || group_code =="null" || group_code == "undefined")
{
	group_code = new String("").toString();
} 

start_time = Trim(new String(Request.Form("start_time_9")));
if (start_time == "" || start_time =="null" || start_time == "undefined")
{
	start_time = new String("00:00:00").toString();
} 

end_time = Trim(new String(Request.Form("end_time_9")));
if (end_time == "" || end_time =="null" || end_time == "undefined")
{
	end_time = new String("00:00:00").toString();
} 

lesson_text = Trim(new String(Request.Form("lesson_text_9")));
if (lesson_text == "" || lesson_text =="null" || lesson_text == "undefined")
{
	lesson_text = new String("").toString();
} 

lesson_id = Trim(new String(Request.Form("lesson_id_9")));
if (lesson_id == "" || lesson_id =="null" || lesson_id == "undefined")
{
	lesson_id = new String("0").toString();
} 

lessondetail = new CoachingLessonObject(lesson_id,lesson_day,start_time,end_time,group_code);
lessondetail.setLessonText(lesson_text);
lessondetail.setDisplayOrder();
lessons.push(lessondetail);

// Line 10

group_code = Trim(new String(Request.Form("group_code_10")));
if (group_code == "" || group_code =="null" || group_code == "undefined")
{
	group_code = new String("").toString();
} 

start_time = Trim(new String(Request.Form("start_time_10")));
if (start_time == "" || start_time =="null" || start_time == "undefined")
{
	start_time = new String("00:00:00").toString();
} 

end_time = Trim(new String(Request.Form("end_time_10")));
if (end_time == "" || end_time =="null" || end_time == "undefined")
{
	end_time = new String("00:00:00").toString();
} 

lesson_text = Trim(new String(Request.Form("lesson_text_10")));
if (lesson_text == "" || lesson_text =="null" || lesson_text == "undefined")
{
	lesson_text = new String("").toString();
} 

lesson_id = Trim(new String(Request.Form("lesson_id_10")));
if (lesson_id == "" || lesson_id =="null" || lesson_id == "undefined")
{
	lesson_id = new String("0").toString();
} 

lessondetail = new CoachingLessonObject(lesson_id,lesson_day,start_time,end_time,group_code);
lessondetail.setLessonText(lesson_text);
lessondetail.setDisplayOrder();
lessons.push(lessondetail);

if (debugging) {
	Response.Write("lessons = "+JSON.stringify(lessons)+"<br /><hr />");
}

// Update coaching dates details from POST'ed data
// Loop through all the lessons, update via lesson_id as unique row identifier
// Need to make sure and have exactly 10 lessons for each day, incl blanks

for (var j=0; j<10; j++) {

	lessondetail = lessons[j];

	SQLend = new String(" WHERE lesson_id = "+lessondetail.lesson_id).toString();

	SQLmiddle = new String("SET ").toString();
	SQLmiddle += " start_time='"+lessondetail.start_time+"',";
	SQLmiddle += " end_time='"+lessondetail.end_time+"',";
	SQLmiddle += " lesson_text='"+lessondetail.lesson_text+"',";
	SQLmiddle += " group_code='"+lessondetail.group_code+"', ";
	SQLmiddle += " group_display_order="+lessondetail.group_display_order+" ";

	// Now do date fields. If null dont insert them as part of the update clause
	//  Access doesnt like setting date fields to ''

	// No date fields to do for this table

	// Default values into result object
	resultObj.result = true;
	resultObj.errno = 0;
	resultObj.description = new String("").toString();

	SQL1 = new String(SQLstart+SQLmiddle+SQLend).toString();
	resultObj.sql = new String(SQL1).toString();

	if (! debugging) {
		try {
			RS = Conn.Execute(SQL1);
		}
		catch(e) {
			resultObj.result = false;
			resultObj.errno = (e.number & 0xFFFF);
			resultObj.description += e.description;
			resultObj.sql = new String(SQL1).toString();
		}
	}

	if (debugging) {
		Response.Write("SQL used for update: <br />"+resultObj.sql+"<br /><br />");

	}
}

// On completion, redirect back to the same lesson day

if (! debugging) {
	Response.Redirect("/coachingday.html?day="+lesson_day);
}

Response.End();


%>


