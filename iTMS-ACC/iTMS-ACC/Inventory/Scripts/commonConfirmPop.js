(function (window, document) {
	"use strict";

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] : null;
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function xmlObject(id) {
		return document.getElementById(id) || window[id] || null;
	}

	function xmlRoot(value) {
		if (!value) {
			return null;
		}
		return value.documentElement || value.XMLDocument && value.XMLDocument.documentElement || value._doc && value._doc.documentElement || value;
	}

	function isNowSelected() {
		var now = document.querySelector('input[name="radConfirm"][value="Y"]');
		return !!(now && now.checked);
	}

	function returnRoot(root) {
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(root);
		} else {
			window["return" + "Value"] = root;
			window.returnvalue = root;
		}
	}

	window.FinalSubmit = function () {
		var root = xmlRoot(xmlObject("RefData"));
		var callFrom = trim(field("hCallFrom") && field("hCallFrom").value);
		var issueType = trim(field("hIssType") && field("hIssType").value);
		var pickType = trim(field("hPickType") && field("hPickType").value);
		var lotOrPickFlag = trim(field("hLotOrPickFlag") && field("hLotOrPickFlag").value);
		if (!root) {
			return false;
		}
		if (callFrom === "DIS") {
			if (issueType === "M" && pickType === "N" && lotOrPickFlag === "P" || issueType === "F") {
				root.setAttribute("Confirm", isNowSelected() ? "Y" : "N");
			} else {
				root.setAttribute("Confirm", "N");
			}
			root.setAttribute("Invoice", "A");
		} else {
			root.setAttribute("Confirm", isNowSelected() ? "Y" : "N");
		}
		if (callFrom === "SUB") {
			root.setAttribute("Confirm", "Y");
			root.setAttribute("ProInv", isNowSelected() ? "Y" : "N");
		}
		returnRoot(root);
		window.close();
		return false;
	};

	window.addEventListener("beforeunload", function () {
		var root = xmlRoot(xmlObject("RefData"));
		if (root) {
			returnRoot(root);
		}
	});
}(window, document));
