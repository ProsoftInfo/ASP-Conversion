function vd(val,todaysdate) 
{

	//init  
	var month;
	var day;
	var year;   
	var delim = new Array("/","");
	var monthArray = new Array(0,31,29,31,30,31,30,31,31,30,31,30,31);

	dtString = val
	while ((dtString.charAt(0) == " ") && (dtString.length != 0))
	dtString = dtString.substring(1,dtString.length - 1)
	while ((dtString.charAt(dtString.length - 1) == " ") && (dtString.length != 0))
	dtString = dtString.substring(0,dtString.length - 1)
	//get date components
	i = 0; startPos = 0; pos = 0;
	do 
	{
		pos = dtString.indexOf(delim[i], startPos);
		i++
	}
	while ((pos == -1) && (i < delim.length));
	if (pos == -1)return false;

	//get day
	day  = parseInt(dtString.substring(startPos,pos),10);
	startPos = pos + 1;
	i = 0;
	//get month
	do 
	{
		pos = dtString.indexOf(delim[i], startPos);
		i++
	}
	while ((pos == -1) && (i < delim.length));
	if (pos == -1) return false;
	month  = parseInt(dtString.substring(startPos,pos),10);
	startPos = pos + 1;

	//get year
	year = parseInt(dtString.substring(startPos,dtString.length),10);
	// valid dateformat check
	if (isNaN(day) || isNaN(month) || isNaN(year)){
		return false;
	}

	// valid month check
	if ((month < 1) || (month > 12)) {
		return false;
	}

	// valid date check
	if ((day < 1) || (day > monthArray[month])) {
		return false;
	}

	// valid year check
	if(year < 1900) {
		return false;
	}
	//check for leap year
	if ((month == 2) && (day == 29))
	if ((((year % 4) == 0) && ((year % 100) != 0)) == false)
	{ 
		return false; 
	}
	ValidDate = val
	var D1 = day;
	var M1 = month - 1;
	var Y1 = year;

	var D2 = parseInt(todaysdate.substring(0,2),10);
	var M2 = parseInt(todaysdate.substring(3,5),10) - 1;
	var Y2 = parseInt(todaysdate.substring(6,10),10);

	date1 = new Date(Y1,M1,D1);
	date2 = new Date(Y2,M2,D2);

	if (date1 < date2) {
		//alert("Date entered should be less than or Equal to today's date");
		return false;
	}
	else
		return true;
	
//if we've gotten this far, return true
return true;
} // end function vd


function checkValidDate(val,todaysdate,flag) 
{
	//init  
	var month;
	var day;
	var year;   
	var delim = new Array("/","");
	var monthArray = new Array(0,31,29,31,30,31,30,31,31,30,31,30,31);

	dtString = val
	while ((dtString.charAt(0) == " ") && (dtString.length != 0))
	dtString = dtString.substring(1,dtString.length - 1)
	while ((dtString.charAt(dtString.length - 1) == " ") && (dtString.length != 0))
	dtString = dtString.substring(0,dtString.length - 1)
	//get date components
	i = 0; startPos = 0; pos = 0;
	do 
	{
		pos = dtString.indexOf(delim[i], startPos);
		i++
	}
	while ((pos == -1) && (i < delim.length));
	if (pos == -1)return false;

	//get day
	day  = parseInt(dtString.substring(startPos,pos),10);
	startPos = pos + 1;
	i = 0;
	//get month
	do 
	{
		pos = dtString.indexOf(delim[i], startPos);
		i++
	}
	while ((pos == -1) && (i < delim.length));
	if (pos == -1) return false;
	month  = parseInt(dtString.substring(startPos,pos),10);
	startPos = pos + 1;

	//get year
	year = parseInt(dtString.substring(startPos,dtString.length),10);
	// valid dateformat check
	if (isNaN(day) || isNaN(month) || isNaN(year)){
		return false;
	}

	// valid month check
	if ((month < 1) || (month > 12)) {
		return false;
	}

	// valid date check
	if ((day < 1) || (day > monthArray[month])) {
		return false;
	}

	// valid year check
	if(year < 1900) {
		return false;
	}
	//check for leap year
	if ((month == 2) && (day == 29))
	if ((((year % 4) == 0) && ((year % 100) != 0)) == false)
	{ 
		return false; 
	}
	ValidDate = val
	var D1 = day;
	var M1 = month - 1;
	var Y1 = year;

	var D2 = parseInt(todaysdate.substring(0,2),10);
	var M2 = parseInt(todaysdate.substring(3,5),10) - 1;
	var Y2 = parseInt(todaysdate.substring(6,10),10);

	date1 = new Date(Y1,M1,D1);
	date2 = new Date(Y2,M2,D2);

	// Flag - 0 : to check for the given date to be less than todays date
	// Flag - 1 : to check for the given date to be greater than todays date
	if (flag == 0) {	
		if (date1 > date2) {
			alert("Date entered should be less than or equal to today's date");
			return false;
		}
		else
			return true;
	}
	if (flag == 1) {	
		if (date1 < date2) {
			alert("Date entered should be greater than or equal to today's date");
			return false;
		}
		else
			return true;
	}
	
} // end function checkValidDate


function resetValue() {
	document.forms[0].reset();
}



function checkDate(val,todaysdate,flag,sMessage) 
{
	//init  
	var month;
	var day;
	var year;   
	var delim = new Array("/","");
	var monthArray = new Array(0,31,29,31,30,31,30,31,31,30,31,30,31);

	dtString = val
	while ((dtString.charAt(0) == " ") && (dtString.length != 0))
	dtString = dtString.substring(1,dtString.length - 1)
	while ((dtString.charAt(dtString.length - 1) == " ") && (dtString.length != 0))
	dtString = dtString.substring(0,dtString.length - 1)
	//get date components
	i = 0; startPos = 0; pos = 0;
	do 
	{
		pos = dtString.indexOf(delim[i], startPos);
		i++
	}
	while ((pos == -1) && (i < delim.length));
	if (pos == -1)return false;

	//get day
	day  = parseInt(dtString.substring(startPos,pos),10);
	startPos = pos + 1;
	i = 0;
	//get month
	do 
	{
		pos = dtString.indexOf(delim[i], startPos);
		i++
	}
	while ((pos == -1) && (i < delim.length));
	if (pos == -1) return false;
	month  = parseInt(dtString.substring(startPos,pos),10);
	startPos = pos + 1;

	//get year
	year = parseInt(dtString.substring(startPos,dtString.length),10);
	// valid dateformat check
	if (isNaN(day) || isNaN(month) || isNaN(year)){
		return false;
	}

	// valid month check
	if ((month < 1) || (month > 12)) {
		return false;
	}

	// valid date check
	if ((day < 1) || (day > monthArray[month])) {
		return false;
	}

	// valid year check
	if(year < 1900) {
		return false;
	}
	//check for leap year
	if ((month == 2) && (day == 29))
	if ((((year % 4) == 0) && ((year % 100) != 0)) == false)
	{ 
		return false; 
	}
	ValidDate = val
	var D1 = day;
	var M1 = month - 1;
	var Y1 = year;

	var D2 = parseInt(todaysdate.substring(0,2),10);
	var M2 = parseInt(todaysdate.substring(3,5),10) - 1;
	var Y2 = parseInt(todaysdate.substring(6,10),10);

	date1 = new Date(Y1,M1,D1);
	date2 = new Date(Y2,M2,D2);

	// Flag - 0 : to check for the given date to be less than todays date
	// Flag - 1 : to check for the given date to be greater than todays date
	if (flag == 0) {	
		if (date1 > date2) {
			alert(sMessage);
			return false;
		}
		else
			return true;
	}
	if (flag == 1) {	
		if (date1 < date2) {
			alert(sMessage);
			return false;
		}
		else
			return true;
	}
	
} // end function checkValidDate



