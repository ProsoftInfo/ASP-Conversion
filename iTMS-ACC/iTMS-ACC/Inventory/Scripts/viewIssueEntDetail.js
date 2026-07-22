(function (window, document) {
	"use strict";

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function submitTo(action) {
		var frm = form();
		if (frm) {
			frm.action = action;
			frm.submit();
		}
	}

	window.DeleteDetails = function (orgCode, issueNo) {
		if (confirm("Do you want to delete this issue Permanently?")) {
			submitTo("mrsIssueDelete.asp?ISSNO=" + encodeURIComponent(issueNo || ""));
		}
		return false;
	};

	window.PrintDetails = function (orgCode, issueNo) {
		if (typeof window.PrintWindow === "function") {
			window.PrintWindow("../reports/PRNDICreateDetails.asp?sTemp=" + encodeURIComponent((orgCode || "") + ":" + (issueNo || "")));
		}
		return false;
	};
}(window, document));
