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

	function submitTo(action) {
		var frm = form();
		if (frm) {
			frm.action = action;
			frm.submit();
		}
	}

	function selectedStatus() {
		var checked = document.querySelector('input[name="rStatus"]:checked');
		return checked ? checked.value : "N";
	}

	window.CreateNew = function () {
		submitTo("GatePassServiceEntry.asp");
		return false;
	};

	window.CheckEdit = function () {
		if (!trim(field("hInvoiceNo") && field("hInvoiceNo").value)) {
			alert("Select the Invoice No");
			return false;
		}
		submitTo("GATEPASSSERVICEENTRYAMD.ASP?GatePassNo=" + encodeURIComponent(field("hInvoiceNo").value) + "&InvoiceType=" + encodeURIComponent(field("hInvoiceType") ? field("hInvoiceType").value : ""));
		return false;
	};

	window.OptionClick = function (obj) {
		if (field("hInvoiceNo")) {
			field("hInvoiceNo").value = obj && obj.value || "";
		}
		return false;
	};

	window.CheckSubmit = function () {
		if (!trim(field("hInvoiceNo") && field("hInvoiceNo").value)) {
			alert("Select the Invoice No");
			return false;
		}
		if (field("B10")) {
			field("B10").disabled = true;
		}
		submitTo("GatePassEntry.asp?GatePassNo=" + encodeURIComponent(field("hInvoiceNo").value) + "&InvoiceType=" + encodeURIComponent(field("hInvoiceType") ? field("hInvoiceType").value : ""));
		return false;
	};

	window.Validate = function () {
		if (field("hInvNo") && field("selInvType")) {
			field("hInvNo").value = field("selInvType").selectedIndex;
		}
		if (field("hFromDate")) {
			field("hFromDate").value = getDate("ctlFromDate");
		}
		if (field("hToDate")) {
			field("hToDate").value = getDate("ctlToDate");
		}
		if (field("hInvType") && field("selInvType")) {
			field("hInvType").value = field("selInvType").value;
		}
		form().submit();
		return false;
	};

	window.Status = function () {
		if (field("hSentType")) {
			field("hSentType").value = selectedStatus();
		}
		form().submit();
		return false;
	};

	window.ChkReset = function () {
		if (field("selInvType")) {
			field("selInvType").selectedIndex = 0;
		}
		return false;
	};

	window.SetDefault = function () {
		ensureCompat();
		if (field("selInvType") && field("hInvType")) {
			field("selInvType").value = field("hInvType").value;
		}
		if (field("hFromDate")) {
			setDate("ctlFromDate", field("hFromDate").value);
		}
		if (field("hToDate")) {
			setDate("ctlToDate", field("hToDate").value);
		}
		return false;
	};
}(window, document));
