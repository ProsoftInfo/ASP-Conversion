// Removes all whitespace characters from start and end of a string
function trim(sString)
{
	sTrimmedString = "";
	if (sString != "")
	{
		var iStart = 0;
		var iEnd = sString.length - 1;
		var sWhitespace = " \t\f\n\r\v";
		
		while (sWhitespace.indexOf(sString.charAt(iStart)) != -1)
		{
			iStart++;
			if (iStart > iEnd)
				break;
		}
		
		// If the string not just whitespace
		if (iStart <= iEnd)
		{
			while (sWhitespace.indexOf(sString.charAt(iEnd)) != -1)
				iEnd--;
			sTrimmedString = sString.substring(iStart,++iEnd);
		}
	}
	return sTrimmedString;
}

function checkmailid(mailid) 
{
	var exclude=/[^@\-\.\w]{2}|[@\.]{2}|(@)[^@]*\1/;
	var check=/@[\w\-]+\./;
	var checkend=/\.[a-zA-Z]{2,3}$/;

	if(((mailid.search(exclude) != -1)||(mailid.search(check))==-1)||(mailid.search(checkend) == -1)) {
		alert("Enter Valid Emailid");
		return false;
	}
	else
		return true;
}

