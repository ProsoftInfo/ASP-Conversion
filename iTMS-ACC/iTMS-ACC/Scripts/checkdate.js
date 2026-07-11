function  validdateone(FromDt,ToDt,CurrentDt,ValDt,datename,message)
{
	var dd,mm,yy,dtval, form, fieldname;
	var Fyymmdd,Tyymmdd,Cyymmdd,Vyymmdd
	
	form = document.forms[0] || null;
	fieldname = form && form.elements ? form.elements[datename] || null : null;
	
	if (checkdt(FromDt,"Finacial Year From Date") && checkdt(ToDt,"Finacial Year To Date") && checkdt(CurrentDt,"Currnet Date") && checkdt(ValDt,message))
	{
			dtval=FromDt;
			dd = dtval.substring(0,2);
   			mm = dtval.substring(3,5);
   			yy = dtval.substring(6,10);
   			Fyymmdd = yy+mm+dd;
		
			dtval=ToDt;
			dd = dtval.substring(0,2);
   			mm = dtval.substring(3,5);
   			yy = dtval.substring(6,10);
   			Tyymmdd = yy+mm+dd;
   			
   			dtval=CurrentDt;
			dd = dtval.substring(0,2);
   			mm = dtval.substring(3,5);
   			yy = dtval.substring(6,10);
   			Cyymmdd = yy+mm+dd;
   			
   			
   			dtval=ValDt;
			dd = dtval.substring(0,2);
   			mm = dtval.substring(3,5);
   			yy = dtval.substring(6,10);
   			Vyymmdd = yy+mm+dd;
			
			
			if (Vyymmdd<Fyymmdd)
			{
				alert(message+" Cannot be Less than Finacial Year From date");
				return false;
			}
			if (Vyymmdd>Tyymmdd)
			{
				alert(message+" Cannot be Greater than Finacial Year To date");
				return false;
			}
			
			if (Vyymmdd>Cyymmdd)
			{
				alert(message+" Cannot be Greater than Current date");
				return false;
			}
			return true;
	}
	else
	{
		return false;
	}
} // end of  validdateone function

function checkdt(dateval,message)
{
	leapyear= new Array(31,29,31,30,31,30,31,31,30,31,30,31);
   	nonleapyear= new Array(31,28,31,30,31,30,31,31,30,31,30,31);

	temp = new Array();
	
	fyea = dateval.substring(6,10);
	fmon  = dateval.substring(3,5);
	fday = dateval.substring(0,2);

	dminus = fday.indexOf("-");
	mminus = fmon.indexOf("-");
	yminus	= fyea.indexOf("-");

	dplus = fday.indexOf("+");
	mplus = fmon.indexOf("+");
	yplus	= fyea.indexOf("+");

	ddot = fday.indexOf(".");
	mdot = fmon.indexOf(".");
	ydot = fyea.indexOf(".");


    //leap year check
	if(fyea % 4 == 0 && ( (fyea % 100) || (fyea % 400))==0) {
		temp = leapyear;
		}
	else {
		temp = nonleapyear;
	}

	if(dateval.length < 0 || dateval == " " || dateval == "") {
		alert("Enter "+message);
		return false;
		}
	else
	if(dateval.indexOf("/") != 2  &&  dateval.lastIndexOf("/") != 5) {
		alert("Enter the Date in dd/mm/yyyy format");
		return false;
		}
	else
	if(dateval.indexOf("/") != 2) {
		alert("Enter the Date in dd/mm/yyyy format");
		return false;
		}
	else
	if(dateval.lastIndexOf("/") != 5) {
		alert("Enter the Date in dd/mm/yyyy format");
		return false;
		}
	else
	if(fyea.length <= 0 || fyea == "0000" || fyea.length < 4) {
		alert("Enter the year");
		return false;
		}
	else
	if(fmon.length <= 0 || fmon == "00" || fmon.charAt(0) == " " || fmon.charAt(1) == " ") {
		alert("Enter the month");
		return false;
		}
	else
	if(fday.length <= 0 || fday == "00" || fday.charAt(0) == " " || fday.charAt(1) == " ") {
		alert("Enter the day");
		return false;
		}
	else
	if(isNaN(fyea) == true || isNaN(fday) == true || isNaN(fmon) == true) {
		alert("Enter Valid Numbers ");
		return false;
		}
	else
	if( dminus != -1 || mminus != -1 || yminus != -1 ) {
		alert("No negative values for Day/Month/Year ");
		return false;
	}
	else
	if( dplus != -1 || mplus != -1 || yplus != -1 ) {
		alert("No Positive values for Day/Month/Year ");
		return false;
	}
	else
	if( ddot != -1 || mdot != -1 || ydot != -1 ) {
		alert("No Decimals for Day/Month/Year Value");
		return false;
	}
	else
	if(fmon > 12 ) {
		alert("Enter a valid Month");
		return false;
		}
	else
	if ( fday > temp[fmon-1] ) {
		alert ("Enter a Valid Date");
		return false;
		}
	else
		return true;	
}//End of checkdt function

//Done by subbiah on 09.06.03 
function checkdtsal(dateval,message,Optval)
{
	leapyear= new Array(31,29,31,30,31,30,31,31,30,31,30,31);
   	nonleapyear= new Array(31,28,31,30,31,30,31,31,30,31,30,31);

	temp = new Array();
	
	fyea = dateval.substring(6,10);
	fmon  = dateval.substring(3,5);
	fday = dateval.substring(0,2);

	dminus = fday.indexOf("-");
	mminus = fmon.indexOf("-");
	yminus	= fyea.indexOf("-");

	dplus = fday.indexOf("+");
	mplus = fmon.indexOf("+");
	yplus	= fyea.indexOf("+");

	ddot = fday.indexOf(".");
	mdot = fmon.indexOf(".");
	ydot = fyea.indexOf(".");


    //leap year check
	if(fyea % 4 == 0 && ( (fyea % 100) || (fyea % 400))==0) {
		temp = leapyear;
		}
	else {
		temp = nonleapyear;
	}

	if(dateval.length < 0 || dateval == " " || dateval == "") {
		if (Optval == "S") {
		return false;
		}
		else
		{
		alert("Enter "+message);
		return false;
		}
	}	
	else
	if(dateval.indexOf("/") != 2  &&  dateval.lastIndexOf("/") != 5) {
		alert("Enter the Date in dd/mm/yyyy format");
		return false;
		}
	else
	if(dateval.indexOf("/") != 2) {
		alert("Enter the Date in dd/mm/yyyy format");
		return false;
		}
	else
	if(dateval.lastIndexOf("/") != 5) {
		alert("Enter the Date in dd/mm/yyyy format");
		return false;
		}
	else
	if(fyea.length <= 0 || fyea == "0000" || fyea.length < 4) {
		alert("Enter the year");
		return false;
		}
	else
	if(fmon.length <= 0 || fmon == "00" || fmon.charAt(0) == " " || fmon.charAt(1) == " ") {
		alert("Enter the month");
		return false;
		}
	else
	if(fday.length <= 0 || fday == "00" || fday.charAt(0) == " " || fday.charAt(1) == " ") {
		alert("Enter the day");
		return false;
		}
	else
	if(isNaN(fyea) == true || isNaN(fday) == true || isNaN(fmon) == true) {
		alert("Enter Valid Numbers ");
		return false;
		}
	else
	if( dminus != -1 || mminus != -1 || yminus != -1 ) {
		alert("No negative values for Day/Month/Year ");
		return false;
	}
	else
	if( dplus != -1 || mplus != -1 || yplus != -1 ) {
		alert("No Positive values for Day/Month/Year ");
		return false;
	}
	else
	if( ddot != -1 || mdot != -1 || ydot != -1 ) {
		alert("No Decimals for Day/Month/Year Value");
		return false;
	}
	else
	if(fmon > 12 ) {
		alert("Enter a valid Month");
		return false;
		}
	else
	if ( fday > temp[fmon-1] ) {
		alert ("Enter a Valid Date");
		return false;
		}
	else
		return true;	
}//End of checkdt function
