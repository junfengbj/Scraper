// Define execute result
var CONST_CHARS_IN_LINE 		= 78;
var CONST_SEPARATE_LINE_1		= "  ----------------------------------------------------------------------------";
var CONST_SEPARATE_LINE_2		= "    --------------------------------------------------------------------------";
var CONST_SEPARATE_LINE_3		= "      ------------------------------------------------------------------------";
var CONST_SEPARATE_LINE_4		= "        ----------------------------------------------------------------------";
var CONST_LIST_LEVEL_1			= "  ";
var CONST_LIST_LEVEL_2			= "    ";
var CONST_LIST_LEVEL_3			= "      ";
var CONST_LIST_LEVEL_4			= "        ";

var CONST_RESULT_OK 			= " [  OK  ]";
var CONST_RESULT_ERROR 			= " [ ERROR]";
var CONST_RESULT_FAILED 		= " [ FAIL ]";

// To include the common module
var objFSO = WScript.CreateObject("Scripting.FileSystemObject");

var objFolder = objFSO.GetFolder(".");

var strName, strDate;

WScript.Echo(objFolder.Path);
WScript.Echo(CONST_SEPARATE_LINE_1);

var colFolder = new Enumerator(objFolder.SubFolders);

for (; !colFolder.atEnd(); colFolder.moveNext()) {
	strName = colFolder.item().Name;
	strDate = colFolder.item().DateCreated;
	strDate = FormatDate(strDate);

	if (strName.indexOf(strDate) > -1) {
		continue;
	}

	try {
		objFSO.moveFolder(strName, strDate + "." + strName);

		DisplayResult(CONST_LIST_LEVEL_1 + strDate + "." + strName.substring(0, 40) + " ", CONST_RESULT_OK, false);
	} catch (e) {
		DisplayResult(CONST_LIST_LEVEL_1 + strDate + "." + strName.substring(0, 40) + " ", CONST_RESULT_FAILED, false);
	}
}

WScript.Echo(CONST_SEPARATE_LINE_1);

//********************************************************************
//* Function: FormatDateTime
//*
//* Purpose: Return the time format in YYYY-MM-DDTHH:mm:ss.
//*
//* Input:
//*  [in]    lngTime	The value of time, 0 for now.
//* Output:
//*  [out]	 string		Formatted time.
//*
//********************************************************************
function FormatDate(strTime) {
	var myDate;
	var strFmt;

	myDate=new Date(strTime);

	strFmt  = myDate.getYear() + ".";
	strFmt += StringAlign(myDate.getMonth() + 1, -2, "0") + ".";
	strFmt += StringAlign(myDate.getDate(), -2, "0");

	return strFmt;
}

//********************************************************************
//* Function: StringPadding
//*
//* Purpose: Padding a string.
//*
//* Input:
//*  [in]    strString		Input string.
//*  [in]    intLength		Length of string after padding, positive or negative.
//*  [in]    strPadding 	Padding string.
//* Output:
//*  [out]	 string 		Padding result.
//*
//********************************************************************
function StringPadding(strString, intLength, strPadding) {
	var blnNegative = false;
	var strAssemble = "";
	var intCharCnts = 0;

	if (intLength < 0) {
		blnNegative = true;
	}

	intLength = Math.abs(intLength);
	intCharCnts = GetCharCount(strString);

	if (intLength < intCharCnts) {
		return strString;
	}

	while (!(strAssemble.length > (intLength - intCharCnts))) {
		strAssemble += strPadding;
	}

	strAssemble = strAssemble.substring(0, intLength - intCharCnts);

	if (blnNegative == false) {
		return strString + strAssemble;
	} else {
		return strAssemble + strString;
	}
}


//********************************************************************
//* Function: StringAlign
//*
//* Purpose: Output string in left, center, right.
//*
//* Input:
//*  [in]    strString		Input string.
//*  [in]    intLength		Length of string after padding, positive or negative.
//*  [in]    strPadding 	Padding string.
//* Output:
//*  [out]	 string 		Aligned string.
//*
//********************************************************************
function StringAlign(strString, intLength, strPadding) {
	var blnNegative = false;
	var strAssemble = "";
	var intCharCnts = 0;

	if (intLength > 0) {
		blnNegative = false;
	} else {
		blnNegative = true;
	}

	intLength = Math.abs(intLength);

	if (intLength <= strString.length) {
		return strString;
	}

	while (strAssemble.length < intLength) {
		strAssemble += strPadding;
	}

	intCharCnts = GetCharCount(strString);

	strAssemble = strAssemble.substring(0, intLength - intCharCnts);

	if (blnNegative == false) {
		return strString + strAssemble;
	} else {
		return strAssemble + strString;
	}
}


//********************************************************************
//* Function: GetCharCount
//*
//* Purpose: Calculate the chars of a string.
//*
//* Input:
//*  [in]    strString		The string.
//* Output:
//*  [out]	 integer		The count of chars.
//*
//********************************************************************
function GetCharCount(strString) {
	var j = 0;
	var k = 0;

	strString = strString.toString();

	k = strString.length;

	for (i = 0; i < k; i++) {
		j++;

		if (strString.substr(i, 1).charCodeAt() > 256) {
			j++;
		}
	}

	return j;
}


//********************************************************************
//* Function: DisplayResult
//*
//* Purpose: Display a result message in entire row.
//*
//* Input:
//*  [in]    strMessage		The message which will be displayed.
//*  [in]    strResult		The result which will be displayed.
//*  [in]    blnStdOut		True = WScript.Std.Out, False = WScript.Echo
//*  [in]    strSperator	The character of sperator.
//* Output:
//*  [out]	 none.
//*
//********************************************************************
function DisplayResult(strMessage, strResult, blnStdOut, strSperator) {
	var j = 0;
	var k = CONST_CHARS_IN_LINE - strResult.length;
	var m = 0;

	var strSperator = arguments[3] ? arguments[3] : ".";

	strMessage = StringPadding(strMessage, CONST_CHARS_IN_LINE, strSperator);

	m = strMessage.length

	for (var i = 0; i < m; i++) {
		j++;

		if (strMessage.substr(i, 1).charCodeAt() > 256) {
			j++;
		}

		if (j > k) {
			strMessage = strMessage.substr(0, i);
			i = m;
		}
	}

	strMessage = strMessage + strResult;

	if (blnStdOut == true) {
		WScript.StdOut.Write(strMessage + "\r");
	} else {
		WScript.Echo(strMessage);
	}
}
