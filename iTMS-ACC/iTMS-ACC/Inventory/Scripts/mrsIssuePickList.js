(function (window, document) {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		if (!frm) {
			return null;
		}
		return frm.elements[name] || frm.elements[name.toLowerCase()] || frm.elements[name.toUpperCase()] || null;
	}

	function setValue(name, value) {
		var element = field(name);
		if (element) {
			element.value = value == null ? "" : String(value);
		}
	}

	function submitTo(action) {
		var frm = form();
		if (!frm) {
			return;
		}
		frm.action = action;
		frm.submit();
	}

	function getDateValue(name) {
		var control = field(name);
		if (!control) {
			return "";
		}
		if (typeof control.getDate === "function") {
			return control.getDate();
		}
		if (typeof control.GetDate === "function") {
			return control.GetDate();
		}
		if (window.ITMSModernCompat && window.ITMSModernCompat.toDisplayDate) {
			return window.ITMSModernCompat.toDisplayDate(control.value);
		}
		return trim(control.value);
	}

	function setDateValue(name, value) {
		var control = field(name);
		if (!control) {
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

	function parseDate(value) {
		var parts = trim(value).split("/");
		var date;
		if (parts.length !== 3) {
			return null;
		}
		date = new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
		return isNaN(date.getTime()) ? null : date;
	}

	function selectedRadioValue(name) {
		var control = field(name);
		var i;
		if (!control) {
			return "";
		}
		if (typeof control.length === "number" && !control.tagName) {
			for (i = 0; i < control.length; i += 1) {
				if (control[i].checked) {
					return control[i].value;
				}
			}
			return "";
		}
		return control.checked ? control.value : "";
	}

	function selectIssueTo() {
		var select = field("selIssueTo");
		var type = trim(field("hIssueToType") && field("hIssueToType").value);
		var code = trim(field("hIssueToCode") && field("hIssueToCode").value);
		var wanted;
		var i;
		if (!select || !select.options) {
			return;
		}
		if (type && code) {
			wanted = (type + ":" + code).toLowerCase();
		} else if (type) {
			wanted = type.toLowerCase();
		} else {
			select.selectedIndex = 0;
			return;
		}
		for (i = 0; i < select.options.length; i += 1) {
			if (trim(select.options[i].value).toLowerCase() === wanted) {
				select.selectedIndex = i;
				break;
			}
		}
	}

	window.ResetData = function () {
		setValue("hFromDate", "");
		setValue("hToDate", "");
		setValue("hIssueToType", "");
		setValue("hIssueToCode", "");
		setValue("hIssueToSubCode", "");
		submitTo("mrsIssuePickList.asp");
	};

	window.Validate = function () {
		setValue("hFromDate", getDateValue("ctlFromDate"));
		setValue("hToDate", getDateValue("ctlToDate"));
		submitTo("mrsIssuePickList.asp");
	};

	window.PickBag = function (issueEntryNo) {
		submitTo("mrsIssuePickItemList.asp?IssueNo=" + encodeURIComponent(issueEntryNo));
	};

	window.PickBagSchedule = function (issueEntryNo, scheduleNo) {
		submitTo("mrsIssuePickItemList.asp?IssueNo=" + encodeURIComponent(issueEntryNo) + "&ScheduleNo=" + encodeURIComponent(scheduleNo));
	};

	window.CheckPick = function () {
		setValue("hCallFrom", selectedRadioValue("radPick"));
		form().submit();
	};

	window.Init = function () {
		setDateValue("ctlFromDate", field("hFromDate") && field("hFromDate").value);
		setDateValue("ctlToDate", field("hToDate") && field("hToDate").value);
		selectIssueTo();
	};

	window.MinDate = function () {
		var minText = field("hFrmDate") && field("hFrmDate").value;
		var maxText = field("hToDate") && field("hToDate").value;
		var fromText = getDateValue("ctlFromDate");
		var toText = getDateValue("ctlToDate");
		var minDate = parseDate(minText);
		var maxDate = parseDate(maxText);
		var fromDate = parseDate(fromText);
		var toDate = parseDate(toText);
		if (minDate && maxDate && fromDate && (fromDate < minDate || fromDate > maxDate)) {
			alert("Date Should be With in the Range " + minText + " to " + maxText);
			setDateValue("ctlFromDate", minText);
			return;
		}
		if (minDate && maxDate && toDate && (toDate < minDate || toDate > maxDate)) {
			alert("Date Should be With in the Range " + minText + " to " + maxText);
			setDateValue("ctlToDate", maxText);
		}
	};

	window.CheckSubmit = function () {
		var usage = "";
		var select = field("selUsage");
		var values = [];
		var i;
		var frm = form();
		setValue("hFromDate", getDateValue("ctlFromDate"));
		setValue("hToDate", getDateValue("ctlToDate"));
		if (select && select.options) {
			for (i = 0; i < select.options.length; i += 1) {
				if (select.options[i].selected) {
					values.push("'" + select.options[i].value + "'");
				}
			}
			usage = values.join(",");
		}
		if (frm) {
			frm.action = "mrsIssuePickList.asp";
			frm.submit();
		}
		return usage;
	};
})(window, document);
