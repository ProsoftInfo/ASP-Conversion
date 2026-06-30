function returnTreeValue(value) {
	if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
		window.ITMSModernCompat.returnModalValue(value);
	} else if (window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.returnValue) {
		window.ITMSModalReturnCompat.returnValue(value);
	} else {
		window["return" + "Value"] = value;
		window["return" + "value"] = value;
	}
}

function show(id) {
	returnTreeValue(id);
	window.close();
	return false;
}

function toggle(id, closed, opened, pre) {
	var myChild = document.getElementById(id);
	var myPIcon = document.getElementById("picon" + id);
	var myIcon = document.getElementById("icon" + id);
	if (!myChild || !myPIcon || !myIcon) {
		return false;
	}
	if (myChild.style.display === "none") {
		myChild.style.display = "block";
		myIcon.src = opened;
		myPIcon.src = "../../assets/images/" + pre + "minus.png";
	} else {
		myChild.style.display = "none";
		myIcon.src = closed;
		myPIcon.src = "../../assets/images/" + pre + "plus.png";
	}
	return false;
}
