(function (window, document) {
	"use strict";

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function openDialog(url, args, features, callback) {
		ensureCompat();
		if (!window.ITMSModernCompat || !window.ITMSModernCompat.openModalDialog) {
			alert("Modern browser compatibility script is still loading. Please try again.");
			return null;
		}
		return window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
	}

	window.DisplayPack = function (pickNo, itemCode, classCode, attId) {
		var url = "IssuePickedPackDetPop.asp?PickNo=" + encodeURIComponent(pickNo || "") +
			"&ItemCode=" + encodeURIComponent(itemCode || "") +
			"&ClassCode=" + encodeURIComponent(classCode || "") +
			"&AttID=" + encodeURIComponent(attId || "");
		openDialog(url, "Picked Pack Details", "dialogHeight:350px;dialogWidth:300px;Status:no;");
		return false;
	};

	window.CheckSubmit = function () {
		window.location.href = "mrsIssuePickListEntry.asp";
		return false;
	};
}(window, document));
