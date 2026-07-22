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
		var match;
		if (!text) {
			return null;
		}
		match = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/);
		if (match) {
			return new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1]));
		}
		match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
		if (match) {
			return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
		}
		return null;
	}

	function dateOnly(value) {
		return value ? new Date(value.getFullYear(), value.getMonth(), value.getDate()) : null;
	}

	function dateText(name) {
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

	function setControlDate(name, value) {
		var control = field(name);
		if (!control) {
			return;
		}
		if (typeof control.setDate === "function") {
			control.setDate(value);
			return;
		}
		if (typeof control.SetDate === "function") {
			control.SetDate(value);
			return;
		}
		control.value = value;
	}

	function setDateLimit(name, limit, value) {
		var control = field(name);
		var method = limit === "min" ? "setMinDate" : "setMaxDate";
		var oldMethod = limit === "min" ? "SetMinDate" : "SetMaxDate";
		if (!control) {
			return;
		}
		if (typeof control[method] === "function") {
			control[method](value);
		} else if (typeof control[oldMethod] === "function") {
			control[oldMethod](value);
		} else {
			control[limit] = window.ITMSModernCompat && window.ITMSModernCompat.toIsoDate ? window.ITMSModernCompat.toIsoDate(value) : value;
		}
	}

	function selectedText(select) {
		return select && select.options && select.selectedIndex >= 0 ? select.options[select.selectedIndex].text : "";
	}

	window.SetDate = function () {
		var fromDate = field("hFromDate") ? field("hFromDate").value : "";
		var toDate = field("hToDate") ? field("hToDate").value : "";
		var today = dateOnly(new Date());
		var toDateValue;
		ensureCompat();
		toDateValue = parseDate(toDate);
		setDateLimit("ctlClosingDate", "min", fromDate);
		if (toDateValue && toDateValue < today) {
			setDateLimit("ctlClosingDate", "max", toDate);
			setControlDate("ctlClosingDate", toDate);
		} else {
			setDateLimit("ctlClosingDate", "max", today);
			setControlDate("ctlClosingDate", today);
		}
		return false;
	};

	window.CheckSubmit = function (todaysDate) {
		var frm = form();
		var unit = field("selUnit");
		var closingFor = field("selFor");
		var closingDate = dateText("ctlClosingDate");
		var today = dateOnly(parseDate(todaysDate) || new Date());
		var closing = dateOnly(parseDate(closingDate));
		if (!unit || unit.selectedIndex === 0) {
			alert("Select Unit");
			if (unit) {
				unit.focus();
			}
			return false;
		}
		if (closing && closing > today) {
			alert("Closing as on should be less than or equal to Today's date");
			return false;
		}
		if (!closingFor || closingFor.selectedIndex === 0) {
			alert("Select Closing Stock For");
			if (closingFor) {
				closingFor.focus();
			}
			return false;
		}
		if (field("hClosingDate")) {
			field("hClosingDate").value = closingDate;
		}
		if (field("hUnitName")) {
			field("hUnitName").value = selectedText(unit);
		}
		if (field("hForName")) {
			field("hForName").value = selectedText(closingFor);
		}
		if (frm) {
			frm.action = "sendToAccountsDetails.asp";
			frm.submit();
		}
		return false;
	};
}(window, document));
