(function (window, document) {
	"use strict";

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] || document.getElementById(name) : document.getElementById(name);
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function parseDate(value) {
		var text = trim(value);
		var match = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/);
		if (match) {
			return new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1]));
		}
		match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
		if (match) {
			return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
		}
		return null;
	}

	function getDate(name) {
		var control = field(name);
		if (!control) {
			return "";
		}
		if (typeof control.getDate === "function") {
			return trim(control.getDate());
		}
		if (typeof control.GetDate === "function") {
			return trim(control.GetDate());
		}
		return trim(control.value);
	}

	function setDate(name, value) {
		var control = field(name);
		if (!control) {
			return;
		}
		if (typeof control.setDate === "function") {
			control.setDate(value);
		} else if (typeof control.SetDate === "function") {
			control.SetDate(value);
		} else {
			control.value = value;
		}
	}

	function selectedGatePassDate() {
		return getDate("ctlGatePassDate");
	}

	function validateGatePassDate() {
		var gatePassDate = parseDate(selectedGatePassDate());
		var invoiceDate = parseDate(field("hInvDate") ? field("hInvDate").value : "");
		if (gatePassDate && invoiceDate && gatePassDate < invoiceDate) {
			alert("The GatePass date must be equal or greater than Invoice Date");
			return false;
		}
		if (field("hGatePassDate")) {
			field("hGatePassDate").value = selectedGatePassDate();
		}
		return true;
	}

	function submitTo(action) {
		var frm = form();
		if (frm) {
			frm.action = action;
			frm.submit();
		}
	}

	function openDialog(url, features, callback) {
		ensureCompat();
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			return window.ITMSModernCompat.openModalDialog(url, "", features || "", callback || function () {});
		}
		return window.open(url, "_blank", "width=600,height=350,resizable=no,status=no,scrollbars=yes");
	}

	function xmlRoot(value) {
		if (!value) {
			return null;
		}
		return value.documentElement || value.XMLDocument && value.XMLDocument.documentElement || value._doc && value._doc.documentElement || value;
	}

	function hasChildren(node) {
		return !!(node && node.childNodes && node.childNodes.length);
	}

	function printUrl(base, includeDate) {
		var params = "?hDCno=" + encodeURIComponent(field("hDCno") ? field("hDCno").value : "") +
			"&Remarks=" + encodeURIComponent(field("txtRemarks") ? field("txtRemarks").value : "") +
			"&GatePassNo=" + encodeURIComponent(field("hGatePassNo") ? field("hGatePassNo").value : "") +
			"&InvoiceType=" + encodeURIComponent(field("hInvoiceType") ? field("hInvoiceType").value : "");
		if (includeDate) {
			params += "&hDCDate=" + encodeURIComponent(field("hGatepassDate") ? field("hGatepassDate").value : field("hGatePassDate") ? field("hGatePassDate").value : "");
		}
		return base + params;
	}

	window.CheckSubmit = function () {
		if (!validateGatePassDate()) {
			return false;
		}
		if (field("button1")) {
			field("button1").disabled = true;
		}
		submitTo("GatePassInsert.asp");
		return false;
	};

	window.CheckModifiy = function () {
		if (!validateGatePassDate()) {
			return false;
		}
		submitTo("GatePassUpdate.asp");
		return false;
	};

	window.Cancel = function (location) {
		window.location.href = location;
		return false;
	};

	window.Check = function (gatePassNo) {
		var unit = field("hOrg") ? field("hOrg").value : "";
		openDialog("GatePassServiceAccountPop.asp?iGPNo=" + encodeURIComponent(gatePassNo || "") + "&ForUnit=" + encodeURIComponent(unit), "dialogHeight:350px;dialogWidth:600px;help:no;status:no", function (result) {
			var root = xmlRoot(result);
			if (hasChildren(root)) {
				submitTo("./../../Purchase/Transaction/InvServiceBillEntry.asp?RefNo=" + encodeURIComponent(gatePassNo || "") + "&RefType=13");
			}
		});
		return false;
	};

	window.Print = function () {
		if (typeof window.PrintWindow === "function") {
			window.PrintWindow(printUrl("../Reports/PRNGatePass.asp", true));
		}
		return false;
	};

	window.FormJJPrint = function () {
		if (typeof window.PrintWindow === "function") {
			window.PrintWindow(printUrl("../Reports/PRNGatePassFormJJ.asp", false));
		}
		return false;
	};

	window.Init = function () {
		ensureCompat();
		if (field("hGatePassDate")) {
			setDate("ctlGatePassDate", field("hGatePassDate").value);
		}
		return false;
	};
}(window, document));
