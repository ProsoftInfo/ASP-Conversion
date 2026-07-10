(function (window, document) {
	"use strict";

	var baseSetDate = window.SetDate;

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.formname || document.forms.formname || document.forms[0] || null;
	}

	function field(name) {
		var frm = form();
		var elements;
		var target;
		var index;
		if (!frm || !frm.elements) {
			return null;
		}
		if (frm.elements[name]) {
			return frm.elements[name];
		}
		elements = frm.elements;
		target = String(name).toLowerCase();
		for (index = 0; index < elements.length; index += 1) {
			if (String(elements[index].name || "").toLowerCase() === target) {
				return elements[index];
			}
		}
		return null;
	}

	function fields(name) {
		var item = field(name);
		var frm = form();
		var result = [];
		var target;
		var index;
		if (item && item.length != null && !item.tagName) {
			return Array.prototype.slice.call(item);
		}
		if (item) {
			return [item];
		}
		if (!frm || !frm.elements) {
			return result;
		}
		target = String(name).toLowerCase();
		for (index = 0; index < frm.elements.length; index += 1) {
			if (String(frm.elements[index].name || "").toLowerCase() === target) {
				result.push(frm.elements[index]);
			}
		}
		return result;
	}

	function valueOf(name, fallback) {
		var item = field(name);
		if (!item) {
			return fallback || "";
		}
		return item.value == null ? fallback || "" : item.value;
	}

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function parseDate(value) {
		var text = trim(value);
		var parts;
		if (!text) {
			return null;
		}
		if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
			parts = text.split("-");
			return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
		}
		parts = text.split(/[\/\-]/);
		if (parts.length >= 3) {
			return new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
		}
		return null;
	}

	function dateOnly(date) {
		return date ? new Date(date.getFullYear(), date.getMonth(), date.getDate()) : null;
	}

	function todayDisplay() {
		var date = new Date();
		var day = String(date.getDate()).padStart(2, "0");
		var month = String(date.getMonth() + 1).padStart(2, "0");
		return day + "/" + month + "/" + date.getFullYear();
	}

	function setDateControl(name, value) {
		var control = field(name) || document.getElementById(name);
		if (!control || value == null || value === "") {
			return;
		}
		if (typeof control.SetDate === "function") {
			control.SetDate(value);
		} else if (typeof control.setDate === "function") {
			control.setDate(value);
		} else if (window.ITMSModernCompat && window.ITMSModernCompat.toIsoDate) {
			control.value = window.ITMSModernCompat.toIsoDate(value);
		} else {
			control.value = value;
		}
	}

	function initChequeDates() {
		var fromDate = valueOf("hFromDate");
		var toDate = valueOf("hToDate");
		var parsedTo = dateOnly(parseDate(toDate));
		var today = dateOnly(new Date());
		ensureCompat();
		if (trim(fromDate)) {
			setDateControl("ctlVouFromDate", fromDate);
		}
		if (parsedTo && today && parsedTo < today) {
			setDateControl("ctlVouToDate", toDate);
		} else {
			setDateControl("ctlVouToDate", todayDisplay());
		}
	}

	function openDialog(url, args, features, callback) {
		ensureCompat();
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
			return;
		}
		window.open(url, "_blank", "height=480,width=420,resizable=no,status=no,scrollbars=yes");
	}

	function checkboxText(item) {
		return item ? item.getAttribute("text") || item.text || "" : "";
	}

	function selectedChequeRows() {
		var checks = fields("Chkbox");
		var count = trim(valueOf("hCnt", String(checks.length)));
		if (count === "1" && checks.length === 1) {
			return checks;
		}
		return checks.filter(function (item) {
			return item.checked;
		});
	}

	window.ShowVouch = function (transNo) {
		openDialog("BankVouchView_San.asp?TransNo=" + encodeURIComponent(transNo), "", "dialogHeight:500px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No", function () {});
		return false;
	};

	window.ShowInsDet = function (transNo) {
		openDialog("InstrumentDetView.asp?TransNo=" + encodeURIComponent(transNo), "", "dialogHeight:250px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No", function () {});
		return false;
	};

	window.PrintCheque = function () {
		var selected = selectedChequeRows();
		var transNos = [];
		var bookNos = [];
		var frm = form();
		selected.forEach(function (item) {
			var textParts = checkboxText(item).split("@");
			transNos.push(item.value);
			bookNos.push(textParts[4] || "");
		});
		if (!transNos.length) {
			alert("Select any one Instrument No For Printing");
			return false;
		}
		frm.action = "ChequeVoucherView.asp?TransNo=" + encodeURIComponent(transNos.join(",")) + "&BookNo=" + encodeURIComponent(bookNos.join(","));
		frm.submit();
		return false;
	};

	window.init = function () {
		initChequeDates();
		return true;
	};

	window.SetDate = function () {
		if (typeof baseSetDate === "function") {
			baseSetDate();
		}
		initChequeDates();
		return true;
	};
}(window, document));
