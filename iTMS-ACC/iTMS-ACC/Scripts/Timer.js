var sDate

function getDate() {
	return new Date().toLocaleDateString();
}

function onLoad() {
	sDate = getDate()
	Delay();
}

function ClockAndAssign() {
	time = new Date ();
	secs = time.getSeconds();
	mins = time.getMinutes();
	hr = time.getHours();
	if (new String(secs).length == "1")
		secs = "0"+secs;
	if (new String(mins).length == "1")
		mins = "0"+mins;
	if (new String(hr).length == "1")
		hr = "0"+hr;
	
	//timer.innerHTML = sDate+"  "+hr+":"+mins+":"+secs
	timer.innerHTML = hr+":"+mins+":"+secs
}

function Delay(){
	ClockAndAssign();
	setTimeout(Delay,120)
}
