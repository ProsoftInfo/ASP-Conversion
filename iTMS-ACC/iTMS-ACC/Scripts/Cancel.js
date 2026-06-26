function Cancel(sLoc) {
	if (confirm("Do you want to Cancel, If so the data entered will be lost.")) {
		window.location.href = sLoc;
	}
}
