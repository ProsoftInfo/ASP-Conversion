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
		return frm && frm.elements ? frm.elements[name] : null;
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function dateControl(name) {
		var control = field(name) || document.getElementById(name) || window[name];
		if (control && window.ITMSModernCompat && window.ITMSModernCompat.decorateDateInput) {
			window.ITMSModernCompat.decorateDateInput(control);
		}
		return control;
	}

	function setDate(name, value) {
		var control = dateControl(name);
		if (!control) {
			return;
		}
		if (typeof control.SetDate === "function") {
			control.SetDate(value);
		} else {
			control.value = value;
		}
	}

	function getDate(name) {
		var control = dateControl(name);
		if (!control) {
			return "";
		}
		if (typeof control.GetDate === "function") {
			return control.GetDate();
		}
		if (typeof control.getDate === "function") {
			return control.getDate();
		}
		return control.value || "";
	}

	function parseLegacyDate(value) {
		var text = trim(value);
		var match;
		if (!text) {
			return null;
		}
		match = /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/.exec(text);
		if (match) {
			return new Date(Number(match[3].length === 2 ? "20" + match[3] : match[3]), Number(match[2]) - 1, Number(match[1]));
		}
		match = /^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/.exec(text);
		if (match) {
			return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
		}
		return null;
	}

	function inRange(value, minValue, maxValue) {
		var date = parseLegacyDate(value);
		var min = parseLegacyDate(minValue);
		var max = parseLegacyDate(maxValue);
		if (!date || !min || !max) {
			return true;
		}
		return date >= min && date <= max;
	}

	function selectedValue(name) {
		var select = field(name);
		if (!select || select.selectedIndex < 0) {
			return "";
		}
		return select.options[select.selectedIndex].value;
	}

	function submitTo(action) {
		var frm = form();
		if (frm) {
			frm.action = action;
			frm.submit();
		}
	}

	function xmlObject(id) {
		return document.getElementById(id) || window[id] || null;
	}

	function xmlDocument(value) {
		if (!value) {
			return null;
		}
		if (value.XMLDocument) {
			return value.XMLDocument;
		}
		if (value._doc) {
			return value._doc;
		}
		if (value.nodeType === 9) {
			return value;
		}
		return value.ownerDocument || null;
	}

	function xmlRoot(value) {
		if (!value) {
			return null;
		}
		return value.documentElement || value.XMLDocument && value.XMLDocument.documentElement || value._doc && value._doc.documentElement || value;
	}

	function getAttr(node, name) {
		return trim(node && node.getAttribute ? node.getAttribute(name) : "");
	}

	function serializeXml(value) {
		var doc = xmlDocument(value);
		var target = doc || xmlRoot(value);
		return target ? new XMLSerializer().serializeToString(target) : "";
	}

	function syncPost(url, body) {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", url, false);
		xhr.send(body || null);
		return xhr.responseText || "";
	}

	function openDialog(url, args, features, callback) {
		ensureCompat();
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			return window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
		}
		alert("Modern browser compatibility script is still loading. Please try again.");
		return null;
	}

	window.Init = function () {
		ensureCompat();
		setDate("ctlFromDate", field("hFrmDate") ? field("hFrmDate").value : "");
		setDate("ctlToDate", field("hToDate") ? field("hToDate").value : "");
	};

	window.checkMR = function () {
		submitTo("MRGENERATIONENTRY.ASP?sOrg=" + encodeURIComponent(field("hUnit") ? field("hUnit").value : ""));
		return false;
	};

	window.MinDate = function () {
		var minDate = field("hFrmDate") ? field("hFrmDate").value : "";
		var maxDate = field("hToDate") ? field("hToDate").value : "";
		var fromDate = getDate("ctlFromDate");
		var toDate = getDate("ctlToDate");
		if (!inRange(fromDate, minDate, maxDate)) {
			alert("Date Should be With in the Range " + minDate + " to " + maxDate);
			setDate("ctlFromDate", minDate);
			return false;
		}
		if (!inRange(toDate, minDate, maxDate)) {
			alert("Date Should be With in the Range " + minDate + " to " + maxDate);
			setDate("ctlToDate", maxDate);
			return false;
		}
		return true;
	};

	window.CheckSubmit = function () {
		if (!field("hCheck") || field("hCheck").value !== "Z") {
			if (field("hFrmDate")) {
				field("hFrmDate").value = getDate("ctlFromDate");
			}
			if (field("hToDate")) {
				field("hToDate").value = getDate("ctlToDate");
			}
		}
		submitTo("mrsMgmtList.asp?ISSTYPE=" + encodeURIComponent(selectedValue("cmbIssType")));
		return false;
	};

	window.DirectIssue = function () {
		var refData = xmlObject("RefData");
		var url = "IssueUsageSelPop.asp?OrgID=" + encodeURIComponent(field("hUnit") ? field("hUnit").value : "");
		openDialog(url, refData, "dialogHeight:340px;dialogWidth:500px;center:yes;help:no;resizable:no;status:no", function (outValue) {
			if (getAttr(xmlRoot(outValue), "Done") === "Y") {
				syncPost("XMLSave.asp?Name=UsageSelection&SessionFlag=true", serializeXml(refData));
				submitTo("DirectIssueItemEntry.asp");
			}
		});
		return false;
	};

	window.OpenMrsPopup = function (mrsNo) {
		openDialog("mrsPopup.asp?MRSNO=" + encodeURIComponent(mrsNo || ""), "", "dialogHeight:470px;dialogWidth:615px;center:yes;help:no;resizable:no;status:no");
		return false;
	};
}(window, document));
