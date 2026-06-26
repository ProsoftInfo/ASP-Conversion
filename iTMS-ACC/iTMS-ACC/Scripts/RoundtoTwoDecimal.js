function getDec2(totVal)
{	
	var totStr,index;
	totStr = totVal.toString();
	index = totStr.indexOf(".");
	if(index == -1)
		return totVal;
	else
	{
		var totRound,decRound,strLen;

		strLen = totVal.toString().length;

		totRound = parseInt(totStr.substr(0,index),10);

		decRound = totStr.substr(index+1,strLen);
		var decFirst,decLast;
		decFirst = decRound.substr(0,2);
		decLast = Math.round("."+decRound.substr(2,decRound.length));
		var decVal;	
		if(isNaN(decLast))
			decVal = (new Number("."+decFirst));
		else
			decVal = (new Number("."+decFirst)) + (new Number(".0"+decLast));
		return 	(totRound+decVal);
	}
}

function getDec3(totVal)
{	
	var totStr,index;
	totStr = totVal.toString();
	index = totStr.indexOf(".");
	if(index == -1)
		return totVal;
	else
	{
		var totRound,decRound,strLen;

		strLen = totVal.toString().length;

		totRound = parseInt(totStr.substr(0,index),10);

		decRound = totStr.substr(index+1,strLen);
		var decFirst,decLast;
		decFirst = decRound.substr(0,3);
		decLast = Math.round("."+decRound.substr(3,decRound.length));
		var decVal;	
		if(isNaN(decLast))
			decVal = (new Number("."+decFirst));
		else
			decVal = (new Number("."+decFirst)) + (new Number(".0"+decLast));
		return 	(totRound+decVal);
	}
}

function getDec4(totVal)
{
	return Math.round(Number(totVal) * 10000) / 10000;
}

function getDec6(totVal)
{
	return Math.round(Number(totVal) * 1000000) / 1000000;
}
