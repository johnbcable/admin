<script language=JScript runat=Server>
// List of function prototypes
//
// timestring()
// datestring()
// timegreeting()
// dateasstring(adate)
// timeasstring(atime)
// today()
// sortabletoday()
// ddmmyyyy(Date)
// getMonthNumber(strMonth)
// getMonthAsString(monthno)
// currentYear()
// currentMonth()
// currentMonthAsString()
// currentDate()
// currentSeason()
// asSortableDate(Date)
// currentAge(strDOB)
// seasonAge(strDOB)
// daysBetween(startdate, enddate)
// nextBirthday(dob)
//
//==================================================
function timestring() {
var d, hh, mm, ss, s = "";
var c = ":";
d = new Date();
hh=d.getHours();
mm=d.getMinutes();
ss=d.getSeconds();
if (hh < 10)
s += "0";
s += hh + c;
if (mm < 10)
s += "0";
s += mm + c;
if (ss < 10)
s += "0";
s += ss;
return (s);
}
//==================================================
function datestring(){
   var d, day, x, s = "";
   var x = new Array("Sunday", "Monday", "Tuesday");
   var x = x.concat("Wednesday","Thursday", "Friday");
   var x = x.concat("Saturday");
   d = new Date();
   day = d.getDay();
   s += x[day] + " "
   s += d.getDate() + "/";
   s += (d.getMonth() + 1) + "/";
   s += d.getYear();
   return(s);
}
//==================================================
function timegreeting() {
var d, h, s = "Good "
d = new Date();
h = d.getHours();
if (h < 12)
s += "morning";
if (h > 11 & h < 18)
s += "afternoon";
if (h > 17)
s += "evening";
return (s);
}
//==================================================
function dateasstring(d){
   var thedate = new Date(d);
   var day, month, x, s = "";
   var x = new Array("Sunday", "Monday", "Tuesday","Wednesday","Thursday", "Friday","Saturday");
    var months = new Array("January","February","March","April","May","June","July","August","September","October","November","December");
   day = thedate.getDay();
    month = thedate.getMonth();
   s += thedate.getDate() + " " + months[month] + ", " + thedate.getYear();
   return(s);
}
//==================================================
function fulltextdate(d){
   var thedate = new Date(d);
   var day, month, x, s = "";
   var x = new Array("Sunday", "Monday", "Tuesday","Wednesday","Thursday", "Friday","Saturday");
	 var months = new Array("January","February","March","April","May","June","July","August","September","October","November","December");
   day = thedate.getDay();
	 month = thedate.getMonth();
   s += x[day] + " " + thedate.getDate() + " " + months[month] + ", " + thedate.getYear();
   return(s);
}
//==================================================
function timeasstring(t){
   var day, x, s = "", minutes, hours;
    hours = t.getHours();
   if (hours < 10)
	s += "0";
   s += hours+":";
   minutes = t.getMinutes();
   if (minutes < 10)
	s += "0";
   s += minutes;
   return(s);
}
//==================================================
function today(){
	var d = new Date();   
	var day, month, x, s = "";
	day = d.getDate();
	if (day < 10)
		s += "0";
	s += day+"/";
	month = (d.getMonth() + 1)
	if (month < 10)
		s+= "0";
	s += month + "/";
   s += d.getFullYear();
   return(s);
}
//==================================================
function sortabletoday(){
	var d = new Date();  
	var s = asSortableDate(d); 
   return(s);
}
//==================================================
function ddmmyyyy(d){
   var day, month, x;
   var s = new String("").toString();
   day = d.getDate();
	if (day < 10)
		s += "0";
	s += day+"/";
   month = (d.getMonth() + 1)
	if (month < 10)
		s+= "0";
	s += month + "/";
   s += d.getFullYear();
   return(new String(s).toString());
}
//==================================================
function getMonthNumber(strMonth){
var months = new Array("January","February","March","April","May","June","July","August","September","October","November","December");
var i, retval;
retval = -1;
for (i = 0; i < 12; i++)
{
  if (strMonth == months[i])
  	retval = i+1;
}
return(retval);
}
//==================================================
function getMonthAsString(monthno){
var months = new Array("","January","February","March","April","May","June","July","August","September","October","November","December");
var i, retval;
retval = new String("").toString();
for (i = 1; i < 13; i++)
{
  if (monthno == i)
  	retval += months[i];
}
return(retval);
}
//==================================================
function currentYear()
{
	var today = new Date();
	var s = new String("").toString();
	s += today.getFullYear();
	return(s);
}
//==================================================
function currentMonth()
{
	var today = new Date();
	var month;
	var s = 0;
	month = today.getMonth()+1;
	if (month < 10)
		s+= "0";
	s += month;
	return(s);
}
//==================================================
function currentMonthAsString()
{
	var today = new Date();
	var month;
	var s = new String("").toString();
	month = today.getMonth()+1;
	if (month < 10)
		s+= "0";
	s += month;
	return(s);
}
//==================================================
function currentDate()
{
	var today = new Date();
	var day;
	var s = new String("").toString();
	day = today.getDate();
	if (day < 10)
		s += "0";
	s += day;
	return(s);
}
//==================================================
function currentSeason()
{
	var thisyear = currentYear();
	var thismonth = currentMonth();
	if (thismonth < 5)
		thisyear = thisyear - 1;
	return(thisyear);
}
//==================================================
function asSortableDate(thedate){
	var d = new Date(thedate);   
	var day, month, x, s = "";
  s += d.getFullYear();
	month = (d.getMonth() + 1)
	if (month < 10)
		s+= "0";
	s += month;
	day = d.getDate();
	if (day < 10)
		s += "0";
	s += day;
   return(s);
}
//==================================================
function currentAge(birthDate) {
	// birthDate must be in DD/MM/YYYY format
	var age;
	var dateArr = birthDate.split("/");
	var theyear = dateArr[2];
	var themonth = new Number(dateArr[1])-1;  // Month numbers start at zero in Javascript
	var theday = dateArr[0];
	var d = new Date(theyear,themonth,theday,0,0,0);
	var age = 0;
	var today = new Date();
	age = Math.floor ( (today - d) / 31556952000 );
	return ( age );
}
//==================================================
function seasonAge(birthdate) {
   var seasonStart;
   var age = 0;
   var dateArr = birthDate.split("/");
   var theyear = dateArr[2];
   var themonth = new Number(dateArr[1])-1;  // Month numbers start at zero in Javascript
   var theday = dateArr[0];
   var d = new Date(theyear,themonth,theday,0,0,0);

   seasonStart = new Date(2013, 3, 1, 0, 0, 0);  // So, April = 3 in JS
   
   age = Math.floor ( (seasonStart - d) / 31556952000 );
   return ( age );
}
//==================================================
function daysBetween(startdate, enddate) {
   var resultObj = new Object();
   var interval;
   var msecondsPerMinute, msecondsPerHour, msecondsPerDay;

   resultObj.days = 0;
   resultObj.hours = 0;
   resultObj.minutes = 0;
   resultObj.seconds = 0;
   resultObj.milliseconds = 0;

   // Make sure time elements of both dates are zeroed.
   startdate.setHours(0);
   startdate.setMinutes(0);
   startdate.setSeconds(0);
   enddate.setHours(0);
   enddate.setMinutes(0);
   enddate.setSeconds(0);

   interval = enddate.getTime() - startdate.getTime();

   // Establish larger units based on milliseconds.
   msecondsPerMinute = 1000 * 60;
   msecondsPerHour = msecondsPerMinute * 60;
   msecondsPerDay = msecondsPerHour * 24;

   // Calculate how many days the interval contains, then subtract that
   // many days from the interval to come up with a remainder.
   resultObj.days = Math.floor( interval / msecondsPerDay );
   interval = interval - (resultObj.days * msecondsPerDay );

   // Repeat the previous calculation on the remainder using hours,
   // then subtract the hours from the remainder.
   resultObj.hours = Math.floor( interval / msecondsPerHour );
   interval = interval - (resultObj.hours * msecondsPerHour );

   resultObj.minutes = Math.floor( interval / msecondsPerMinute );
   interval = interval - (resultObj.minutes * msecondsPerMinute );

   resultObj.seconds = Math.floor( interval / 1000 );

   return(resultObj);

}
//==================================================
function nextBirthday(dob) {

   var resultDate = new Date();
   var interval;
   var todaysdate = new Date();
   var dobDay, dobMonth, thisMonth, thisYear, thisDay;

   dobDay = dob.getDate();
   dobMonth = dob.getMonth();
   thisMonth = new Number(currentMonth()) - 1;
   thisDay = todaysdate.getDate(); 
   thisYear = todaysdate.getFullYear();

   if (thisDay == dobDay && thisMonth == dobMonth)
   {
		resultDate = new Date();   
   } 	
   if (dobMonth > thisMonth) {
   		resultDate = new Date(thisYear, dobMonth, dobDay, 0, 0, 0);
   }
   if (dobMonth < thisMonth) {
   		resultDate = new Date(thisYear+1, dobMonth, dobDay, 0, 0, 0);
   }
   if (dobMonth == thisMonth) {
   		if (! (thisDay == dobDay)) {
	   		resultDate = new Date(thisYear, dobMonth, dobDay, 0, 0, 0);
   		}
   }
   return ( resultDate );
}


</script>