<script language="JScript" CODEPAGE="65001" runat=Server>
//
//  setNews( newsobj)                                                 returns boolean
//  getNews( newsid)                                                  returns news object
//  printNews(newsobj)																								returns string
//  getNewsSummaries()				  																			returns string
//  getNewsHeaders( )                                                 returns string
// 	getNewsByCategory(category)                              					returns string
// 	getNewsExcludingCategory(category)                 								returns string
//
function setNews(newsobj)
{
	// Establish local variables
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	var onews = new Object();
	var mDateObj, dummy1;
	var resultObj = new Object();
	var e;
	var SQLstart, SQLend, SQLmiddle;
	//
	resultObj.result = true;
	resultObj.errno = 0;
	resultObj.description = new String("").toString();
	resultObj.sql = new String("").toString();
// Initialise the news object
	onews.newsid = newsobj.newsid;
	onews.title = new String(newsobj.title).toString();
	onews.author = new String(newsobj.author).toString();
	onews.webpage = new String(newsobj.webpage).toString();
	onews.visible = new String(newsobj.visible).toString();
	onews.newsfrom = new String(newsobj.newsfrom).toString();
	onews.newsuntil = new String(newsobj.newsuntil).toString();
	onews.newspriority = new String(newsobj.newspriority).toString();
	onews.category = new String(newsobj.category).toString();
	onews.blurb = new String(newsobj.blurb).toString();
	onews.thumbnail = new String(newsobj.thumbnail).toString();
	
	if (onews.title == "null" || onews.title == "undefined")
		onews.title = new String("").toString();
	if (onews.author == "null" || onews.author == "undefined")
		onews.author = new String("").toString();
	if (onews.webpage == "null" || onews.webpage == "undefined")
		onews.webpage = new String("").toString();
	if (onews.visible == "null" || onews.visible == "undefined")
		onews.visible = new String("").toString();
	if (onews.category == "null" || onews.category == "undefined")
		onews.category = new String("").toString();
	if (onews.blurb == "null" || onews.blurb == "undefined")
		onews.blurb = new String("").toString();
	if (onews.thumbnail == "null" || onews.thumbnail == "undefined")
		onews.thumbnail = new String("").toString();
	
	// Adjust strings
	onews.title = onews.title.replace(/'/g,"''");
	onews.blurb = onews.blurb.replace(/'/g,"''");

	// Set up database connection
	dbconnect=Application("hamptonsportsdb");
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	
	SQLstart = new String("UPDATE news ")
	SQLend = new String(" WHERE [newsid] ="+onews.newsid).toString();
	SQLmiddle = new String("SET ").toString();
	SQLmiddle += " [title] = '"+onews.title+"', ";
	SQLmiddle += " [author] = '"+onews.author+"', ";
	SQLmiddle += " [webpage] = '"+onews.webpage+"', ";
	SQLmiddle += " [visible] = '"+onews.visible+"', ";
	SQLmiddle += " [newsfrom] = '"+onews.newsfrom+"', ";
	SQLmiddle += " [newsuntil] = '"+onews.newsuntil+"', ";
	SQLmiddle += " [newspriority] = "+onews.newspriority+", ";
	SQLmiddle += " [category] = '"+onews.category+"', ";
	SQLmiddle += " [blurb] = '"+onews.blurb+"', ";
	SQLmiddle += " [thumbnail] = '"+onews.thumbnail+"' ";

	SQL1 = new String(SQLstart+SQLmiddle+SQLend);
	
	try {
		RS = Conn.Execute(SQL1);
	}
	catch(e) {
		resultObj.result = false;
		resultObj.errno = (e.number & 0xFFFF);
		resultObj.description += e.description;
		resultObj.sql = new String(SQL1).toString();
	}
	return(resultObj);
}
//=======================================================================================================
function getNews(newsid)
{
	// Establish local variables
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	var localid = newsid;
	var onews = new Object();
	var mDateObj, dummy1;
	// Initialise the news object
	onews.newsid = localid;
	onews.title = new String("").toString();
	onews.author = new String("").toString();
	onews.webpage = new String("").toString();
	onews.visible = new String("").toString();
	onews.newsfrom = new String("").toString();
	onews.newsuntil = new String("").toString();
	onews.newspriority = new String("").toString();
	onews.category = new String("").toString();
	onews.blurb = new String("").toString();
	onews.thumbnail = new String("").toString();
	// Set up database connection
	dbconnect=Application("hamptonsportsdb");
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	SQL1 = new String("SELECT * from news where newsid = "+localid).toString();
	RS = Conn.Execute(SQL1);
	if (! RS.EOF)
	{
		onews.title = new String(RS("title")).toString();
		onews.author = new String(RS("author")).toString();
		onews.webpage = new String(RS("webpage")).toString();
		onews.thumbnail = new String(RS("thumbnail")).toString();
		onews.visible = new String(RS("visible")).toString();
		mDateObj=new Date(RS("newsfrom"));
		dummy1 = mDateObj.valueOf();
		if (dummy1 == 0) // no date in database
			onews.newsfrom = new String("").toString();
		else
			onews.newsfrom = new String(ddmmyyyy(mDateObj)).toString();
		mDateObj=new Date(RS("newsuntil"));
		dummy1 = mDateObj.valueOf();
		if (dummy1 == 0) // no date in database
			onews.newsuntil = new String("").toString();
		else
			onews.newsuntil = new String(ddmmyyyy(mDateObj)).toString();
		onews.newspriority = new String(RS("newspriority")).toString();
		onews.category = new String(RS("category")).toString();
		onews.blurb = new String(RS("blurb")).toString();
		// now deal with nulls
		if (onews.title == "null" || onews.title == "undefined")
			onews.title = new String("").toString();
		if (onews.author == "null" || onews.author == "undefined")
			onews.author = new String("").toString();
		if (onews.webpage == "null" || onews.webpage == "undefined")
			onews.webpage = new String("").toString();
		if (onews.visible == "null" || onews.visible == "undefined")
			onews.visible = new String("").toString();
		if (onews.category == "null" || onews.category == "undefined")
			onews.category = new String("").toString();
		if (onews.blurb == "null" || onews.blurb == "undefined")
			onews.blurb = new String("").toString();
		if (onews.thumbnail == "null" || onews.thumbnail == "undefined")
			onews.thumbnail = new String("").toString();
	}
	RS.Close();
	Conn.Close();
	return (onews);
}
// ================================================================
function printNews(newsobj)
{
	// Establish local variables
	var sReport = new String("").toString();
	
	sReport += "<table><caption>Contents of newsobj</caption>";
	sReport += "<thead><tr><th>Variable</th><th>Contents</th></tr></thead><tbody>";
	sReport += "<tr><td>ID</td><td>"+newsobj.newsid +"</td></tr>";
	sReport += "<tr><td>Title</td><td>"+newsobj.title +"</td></tr>";
  	sReport += "<tr><td>Author</td><td>"+newsobj.author +"</td></tr>";
	sReport += "<tr><td>Web Page</td><td>"+newsobj.webpage +"</td></tr>";
  	sReport += "<tr><td>Visible?</td><td>"+newsobj.visible +"</td></tr>";
	sReport += "<tr><td>News From</td><td>"+newsobj.newsfrom +"</td></tr>";
	sReport += "<tr><td>News Until</td><td>"+newsobj.newuntil +"</td></tr>";
	sReport += "<tr><td>Priority (low is higher priority)</td><td>"+newsobj.newspriority +"</td></tr>";
	sReport += "<tr><td>Category</td><td>"+newsobj.category +"</td></tr>";
	sReport += "<tr><td>Blurb</td><td>"+newsobj.blurb +"</td></tr>";
	sReport += "<tr><td>Thumbnail image (optional)</td><td>"+newsobj.thumbnail +"</td></tr>";
	sReport += '</tbody><tfoot><tr><td colspan="2">&nbsp;</td></tr></tfoot>';

	sReport += "</table>"
	
	return(sReport);
}
// ================================================================
function getNewsSummaries()
{
	// Establish local variables
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	var debugging = false;
	var editline, newssubject;
	var newslist = new String("").toString();
	var newsknt = 0;
	var newsmax = 5;
	// Set up database connections
	dbconnect=Application("hamptonsportsdb");
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	// Need to query last 5 newss sent and return as string
	// title<hr><br>event1<br>event2<br> .... eventN<hr>
	SQL1 = new String("SELECT * from news where visible = 'Y' order by newsid").toString();
	newslist += '<div class="seealso"><ul>';
	RS = Conn.Execute(SQL1);
	newsknt = 0;
	while (! RS.EOF && newsknt < newsmax)
	{
		newslist += '<li><a href="'+RS("webpage")+'">'+RS("title")+'</a><br /><p>'+RS("blurb")+'</p><br /></li>';
		newsknt++;
		RS.MoveNext();
	}
	newslist += "</ul></div>";
	RS.Close();
	Conn.Close();
	return (newslist);
}
//=======================================================================================================
function getNewsHeaders()
{
	// Run to get display items for the right-hand column normally
	// Establish local variables
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	var debugging = false;
	var dummy;
	var editline, newssubject;
	var newslist = new String("").toString();
	var newsknt = 0;
	var newsmax = 3;
	var noimage = new String("NONE").toString();
	// Set up database connections
	var dbconnect=Application("hamptonsportsdb"); 
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	// Need to query last 5 newss sent and return as string
	// title<hr><br>event1<br>event2<br> .... eventN<hr>
	SQL1 = new String("SELECT * from news where visible = 'Y' order by newspriority asc, newsid desc").toString();
	newslist += '<ol id="news"><div id="newsheader"><h3><a href="news.asp">Club News</a></h3></div>';
	RS = Conn.Execute(SQL1);
	newsknt = 0;
	while (! RS.EOF && newsknt < newsmax)
	{
		newslist += "<li>";
		dummy = new String(RS("thumbnail")).toString();
		if (dummy == "" || dummy == "null" || dummy == "undefined")
			dummy = new String("NONE").toString();
		if (! (dummy == noimage) )    // if there is a thumbnail then it must form part of the display
			newslist += '<a href="'+RS("webpage")+'"><img src="images/'+RS("thumbnail")+'" alt="'+RS("title")+'" /></a>';
		newslist += '<h4><a href="'+RS("webpage")+'">'+RS("title")+'</a></h4>';
		newslist += '<p>'+RS("blurb")+'</p>';
		newslist += "</li>"
		newsknt++;
		RS.MoveNext();
	}
	newslist += "</ol>";
	RS.Close();
	Conn.Close();
	return (newslist);
}

//=======================================================================================================
function getNewsByCategory(thecategory)
{
	// Establish local variables
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	var debugging = false;
	var editline, newssubject;
	var newslist = new String("").toString();
	var newsknt = 0;
	var newsmax = 10;
	var mycategory = new String(thecategory).toString();
	var divid = new String("").toString();
	//
	mycategory = mycategory.toLowerCase();
	divid += "mynewslinks"
	// Set up database connections
	var dbconnect=Application("hamptonsportsdb"); 
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	// Need to query last 5 newss sent and return as string
	// title<hr><br>event1<br>event2<br> .... eventN<hr>
	SQL1 = new String("SELECT * from news where category = '"+thecategory+"' and visible = 'Y' order by newsid").toString();
	newslist += '<div id="'+divid+'"><h5 class="navheader"><span style="margin-left:10px;">Club News</span></h5><ul>';
	RS = Conn.Execute(SQL1);
	newsknt = 0;
	while (! RS.EOF && newsknt < newsmax)
	{
		newslist += '<li><a href="'+RS("webpage")+'">'+RS("title")+'</a><br /><p class="blurb">'+RS("blurb")+'</p></li>';
		newsknt++;
		RS.MoveNext();
	}
	newslist += "</ul></div>";
	RS.Close();
	Conn.Close();
	return (newslist);
}
//=======================================================================================================
function getNewsExcludingCategory(thecategory)
{
	// Establish local variables
	var RS, RS2, Conn, SQL1, SQL2, dbconnect, uniqref;
	var debugging = false;
	var editline, newssubject;
	var newslist = new String("").toString();
	var newsknt = 0;
	var newsmax = 10;
	var mycategory = new String(thecategory).toString();
	var divid = new String("").toString();
	//
	mycategory = mycategory.toLowerCase();
	divid += "othernewslinks"
	// Set up database connections
	var dbconnect=Application("hamptonsportsdb"); 
	Conn = Server.CreateObject("ADODB.Connection");
	RS = Server.CreateObject("ADODB.RecordSet");
	Conn.Open(dbconnect);
	// Need to query last 5 newss sent and return as string
	// title<hr><br>event1<br>event2<br> .... eventN<hr>
	SQL1 = new String("SELECT * from news where category not like '"+thecategory+"' and visible = 'Y' order by newsid").toString();
	newslist += '<div id="'+divid+'"><h5 class="navheader"><span style="margin-left:10px;">Club News</span></h5><ul>';
	RS = Conn.Execute(SQL1);
	newsknt = 0;
	while (! RS.EOF && newsknt < newsmax)
	{
		newslist += '<li><a href="'+RS("webpage")+'">'+RS("title")+'</a><br /><p class="blurb">'+RS("blurb")+'</p></li>';
		newsknt++;
		RS.MoveNext();
	}
	newslist += "</ul></div>";
	RS.Close();
	Conn.Close();
	return (newslist);
}
</script>
