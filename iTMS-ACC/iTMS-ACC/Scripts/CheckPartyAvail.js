(function () {
	"use strict";

	window.CheckPartyAvail = function (sTemp) {
		var xhr = new XMLHttpRequest();
		var messages = {
			K: "Please Enter Commission Agents ",
			U: "Please Enter Agents ",
			P: "Please Enter Depot Agents ",
			D: "Please Enter Party "
		};
		var code = String(sTemp || "");

		xhr.open("POST", "PartyCheck.asp?sCallType=" + encodeURIComponent(code), false);
		xhr.send(null);

		if (String(xhr.responseText) === "0") {
			if (messages[code]) {
				alert(messages[code]);
				return 1;
			}
		}
		return 0;
	};
}());
