<script language=JScript CODEPAGE="65001" runat=Server>;
/*

LTrim(string) : Returns a copy of a string without leading spaces.
RTrim(string) : Returns a copy of a string without trailing spaces.
Trim(string) : Returns a copy of a string without leading or trailing spaces
StringFromDB(string):  Returns a copy of a string where "null" and "undefined" are set to ""
replaceAll(string,to_replcae,with_this):  Returns a copy of string with all occurrences of to_replace changed to with_this
toPropercase(string) : Returns a copy of the string with a leading capital letter and all other characters in lower case
asUTF16(string)  :  Returns supplied string as Unicode-escaped string
Left(string, n) : Returns n characters from the left-hand end of string
Right(string, n) : Returns n characters from the right-hand end of string
Lpad(string, padchar, numchars) : Returns a string left-padded with up to numchars padchar
Rpadsstring, padchar, numchars)) : Returns a string right-padded with up to numchars padchar

*/

/*
==================================================================
LTrim(string) : Returns a copy of a string without leading spaces.
==================================================================
*/
function LTrim(str)
/*
   PURPOSE: Remove leading blanks from our string.
   IN: str - the string we want to LTrim
*/
{
   var whitespace = new String(" \t\n\r");

   var s = new String(str);

   if (whitespace.indexOf(s.charAt(0)) != -1) {
      // We have a string with leading blank(s)...

      var j=0, i = s.length;

      // Iterate from the far left of string until we
      // don't have any more whitespace...
      while (j < i & whitespace.indexOf(s.charAt(j)) != -1)
         j++;

      // Get the substring from the first non-whitespace
      // character to the end of the string...
      s = s.substring(j, i);
   }
   return s;
}

/*
==================================================================
RTrim(string) : Returns a copy of a string without trailing spaces.
==================================================================
*/
function RTrim(str)
/*
   PURPOSE: Remove trailing blanks from our string.
   IN: str - the string we want to RTrim

*/
{
   // We don't want to trim JUST spaces, but also tabs,
   // line feeds, etc.  Add anything else you want to
   // "trim" here in Whitespace
   var whitespace = new String(" \t\n\r");

   var s = new String(str);

   if (whitespace.indexOf(s.charAt(s.length-1)) != -1) {
      // We have a string with trailing blank(s)...

      var i = s.length - 1;       // Get length of string

      // Iterate from the far right of string until we
      // don't have any more whitespace...
      while (i >= 0 & whitespace.indexOf(s.charAt(i)) != -1)
         i--;

      // Get the substring from the front of the string to
      // where the last non-whitespace character is...
      s = s.substring(0, i+1);
   }

   return s;
}

/*
=============================================================
Trim(string) : Returns a copy of a string without leading or 
trailing spaces
=============================================================
*/
function Trim(str)
/*
   PURPOSE: Remove trailing and leading blanks from our string.
   IN: str - the string we want to Trim

*/
{
	   return RTrim(LTrim(str));
}
/*
=============================================================
StringFromDB(string) : Returns a copy of a Trim'med string 
where "null" and "undefined" are set to ""
=============================================================
*/
function StringFromDB(str)
/*
   PURPOSE: Return a string in usable from from  the database
   IN: str - the string we need to return sensibly

   RETVAL: A usable string from the database
*/
{
	var retval = new String(RTrim(LTrim(str))).toString();
	if (retval == "null" || retval == "undefined")
		retval = new String("").toString();
	return new String(retval).toString();
}
/*
=============================================================
StringToDB(string) : Returns a copy of a string where apostrophes
e.g. O'Donnell can be sent on to the database correctly
=============================================================
*/
function StringToDB(str)
/*
   PURPOSE: Returns a copy of a string where apostrophes
e.g. O'Donnell can be sent on to the database correctly
   IN: str - the string we need to return sensibly

   RETVAL: A usable string for updating the database
*/
{
	var retval = new String(RTrim(LTrim(str))).toString();
	if (retval == "null" || retval == "undefined")
		retval = new String("").toString();
	else
		retval = new String(retval.replace("'","''")).toString();
	return new String(retval).toString();
}

/*
=============================================================
replaceAll(string) : Returns a copy of a string where all
occurrences of to_replace have been replaced by with_this
=============================================================
*/
function replaceAll(txt, to_replace, with_this) {
  return txt.replace(new RegExp(to_replace, 'g'),with_this);
}

/*
==================================================================
toProperCase(string) : Returns a copy of a string with leading
capital letter
==================================================================
*/
function toProperCase(str)
{
 
   var s = new String(str).toString();
   var head, tail;

   s = s.toLowerCase();
   head = s.substr(0,1);
   tail = s.substr(1);
   head = head.toUpperCase();

   s = new String(head+tail).toString();

   return s;
}

/*
==================================================================
asUTF16(string) : Returns a string with all letters in 
input string presented as Unicode values
==================================================================
*/
function asUTF16(input) 
{
    function pad_four(input) {
        var l = input.length;
        if (l == 0) return '0000';
        if (l == 1) return '000' + input;
        if (l == 2) return '00' + input;
        if (l == 3) return '0' + input;
        return input;
    }
    var output = '';
    for (var i = 0, l = input.length; i < l; i++)
        output += '\\u' + pad_four(input.charCodeAt(i).toString(16));
    return output;
}

/*
==================================================================
asASCII(string) : Returns a string with all letters in 
input string presented as ASCII values
==================================================================
*/
function asASCII(input) 
{
    var output = '';
    for (var i = 0, l = input.length; i < l; i++)
        output += input.charCodeAt(i)+" ".toString();
    return output;
}
/*
==================================================================
Left(string) : Returns n characters from the left-hand 
end of the string
==================================================================
*/
function Left(str, n){
  if (n <= 0)
      return "";
  else if (n > String(str).length)
      return str;
  else
      return String(str).substring(0,n);
}
/*
==================================================================
Right(string) : Returns n characters from the right-hand 
end of the string
==================================================================
*/
function Right(str, n){
    if (n <= 0)
       return "";
    else if (n > String(str).length)
       return str;
    else {
       var iLen = String(str).length;
       return String(str).substring(iLen, iLen - n);
    }
}

/*
==================================================================
Lpad(string, padchar, numchars) : Returns a string left-padded 
with up to numchars padchar
==================================================================
*/
// left padding s with c to a total of n chars
function Lpad(s, c, n) {
  if (! s || ! c || s.length >= n) {
    return s;
  }
  var max = (n - s.length)/c.length;
  for (var i = 0; i < max; i++) {
    s = c + s;
  }
  return s;
}
 
/*
==================================================================
Rpadsstring, padchar, numchars)) : Returns a string right-padded 
with up to numchars padchar
==================================================================
*/
// right padding s with c to a total of n chars
function Rpad(s, c, n) {
  if (! s || ! c || s.length >= n) {
    return s;
  }
  var max = (n - s.length)/c.length;
  for (var i = 0; i < max; i++) {
    s += c;
  }
  return s;
}
</SCRIPT>
