(function () {
	"use strict";

	window.CreateTempItem = function (appCode, modCode, creStage, sItemtype, sOrgCode, callback) {
		var url = "../../inventory/master/tempitmCreationEntry.asp?appCode=" + encodeURIComponent(appCode) +
			"&modCode=" + encodeURIComponent(modCode) +
			"&creStage=" + encodeURIComponent(creStage) +
			"&itmType=" + encodeURIComponent(sItemtype) +
			"&sOrgCode=" + encodeURIComponent(sOrgCode);
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			return window.ITMSModernCompat.openModalDialog(
				url,
				"A",
				"dialogHeight:230px;dialogWidth:600px;center:Yes;help:No;resizable:No;status:No",
				callback
			);
		}
		alert("The compatibility script is still loading. Please try again.");
		return null;
	};
}());
