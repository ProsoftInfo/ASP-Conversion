(function (window, document) {
	"use strict";

	function selectedMode() {
		var checked = document.querySelector('input[name="radType"]:checked');
		return checked ? checked.value : "";
	}

	window.CheckSubmit = function () {
		var mode = selectedMode();
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(mode);
		} else {
			window["return" + "Value"] = mode;
			window.returnvalue = mode;
		}
		window.close();
		return false;
	};
}(window, document));
