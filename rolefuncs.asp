<script language="JScript" CODEPAGE="65001" runat=Server>
// List of function prototypes
//
// signedin()
// isAdministrator()
// isSupport()
// isMembershipEditor()
// is200ClubEditor()
// isFixtureEditor()
// isPublicationEditor
// isCommittee()
//
// Role access is based on numeric access level:
// 99 		Support (top level)
// 90 - 98	Administrator
// 50 - 60  Committee member
// 20		Ordinary member
//==================================================
function signedin() {
var result=false;
var zz, result, thevalue;
zz=getUserID();
if(zz == "NONE")
	result=false;
else
	result=true;
return(result);
}
//==================================================
function isSupport() {
var result=false;
var thelevel=getUserAccessLevel();
if (thelevel > 95)    					// is main tech support person
	result=true;
return(result);
}
//==================================================
function isAdministrator() {
var result=false;
var thelevel=getUserAccessLevel();
if (thelevel > 89)    	   				// Can alter member details
	result=true;
return(result);
}
//==================================================
function isCommittee() {
var result=false;
var thelevel=getUserAccessLevel();
if (thelevel == 50 || thelevel > 89)    // Read-only access to all member details
	result=true;
return(result);
}
//==================================================
function isMembershipEditor() {
var result=false;
var thelevel=getUserAccessLevel();
if (thelevel == 85 || thelevel > 94)    // can edit membership details
	result=true;
return(result);
}
//==================================================
function is200ClubEditor() {
var result=false;
var thelevel=getUserAccessLevel();
if (thelevel == 55 || thelevel > 95)    // can edit 200 club details
	result=true;
return(result);
}
//==================================================
function isFixtureEditor() {
// var result=false;
// var thelevel=getUserAccessLevel();
// if (thelevel == 65 || thelevel > 95)    // can edit fixture details
// 	result=true;
// return(result);
	return(true);
}
//==================================================
function isPublicationEditor() {
var result=false;
var thelevel=getUserAccessLevel();
if (thelevel == 75 || thelevel > 90)    // can edit article details
	result=true;
return(result);
}
</script>
