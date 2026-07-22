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

	function dateControl(name) {
		var control = field(name) || window[name];
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
			control.value = value || "";
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

	function selectedOption(select) {
		if (!select || select.selectedIndex < 0) {
			return { value: "", text: "" };
		}
		return {
			value: select.options[select.selectedIndex].value,
			text: select.options[select.selectedIndex].text
		};
	}

	function selectByValue(name, value) {
		var select = field(name);
		var wanted = trim(value).toLowerCase();
		if (!select || !select.options) {
			return;
		}
		for (var i = 0; i < select.options.length; i += 1) {
			if (trim(select.options[i].value).toLowerCase() === wanted) {
				select.selectedIndex = i;
				return;
			}
		}
	}

	function checkedRadioValue(name) {
		var checked = document.querySelector('input[name="' + name + '"]:checked');
		return checked ? checked.value : "";
	}

	function setRadioByValue(name, value) {
		var radios = document.querySelectorAll('input[name="' + name + '"]');
		var wanted = trim(value);
		for (var i = 0; i < radios.length; i += 1) {
			if (radios[i].value === wanted || wanted === "A" && radios[i].value === "B") {
				radios[i].checked = true;
				return;
			}
		}
	}

	function submitTo(action) {
		var frm = form();
		if (frm) {
			frm.action = action;
			frm.submit();
		}
	}

	function selectedIssue() {
		var frm = form();
		var radios = document.querySelectorAll('input[name="rSelect"]');
		var checked = document.querySelector('input[name="rSelect"]:checked');
		var index = -1;
		if (!frm || Number(field("hCtr") && field("hCtr").value || 0) === 0) {
			return null;
		}
		if (!checked) {
			return null;
		}
		for (var i = 0; i < radios.length; i += 1) {
			if (radios[i] === checked) {
				index = i + 1;
				break;
			}
		}
		return {
			index: index,
			issueNo: checked.value,
			issueDate: field("hIssDate" + index) ? field("hIssDate" + index).value : "",
			issueTypeCode: field("hIssTypeCode" + index) ? field("hIssTypeCode" + index).value : "",
			tempValues: (field("hUnit") ? field("hUnit").value : "") + ":" + checked.value + ":" + (field("hIssDate" + index) ? field("hIssDate" + index).value : "")
		};
	}

	function requireIssue() {
		var issue = selectedIssue();
		if (!issue) {
			alert("Select any one Issue");
			return null;
		}
		return issue;
	}

	function openDialog(url, args, features, callback) {
		ensureCompat();
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			return window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
		}
		alert("Modern browser compatibility script is still loading. Please try again.");
		return null;
	}

	window.DirectIssue = function () {
		var issueType = selectedOption(field("cmbIssType")).value;
		if (trim(issueType).toUpperCase() === "SEL") {
			issueType = "GEN";
		}
		submitTo("MATERIALISSUEENTRY.ASP?TYPE=" + encodeURIComponent(issueType));
		return false;
	};

	window.checkSubmit = function () {
		var issue = requireIssue();
		var unit = field("hUnit") ? field("hUnit").value : "";
		if (!issue) {
			return false;
		}
		if (typeof window.PrintWindow === "function") {
			window.PrintWindow("../reports/PRNDICreateDetails.asp?sTemp=" + encodeURIComponent(unit + ":" + issue.issueNo));
		}
		return false;
	};
	window.CheckSubmit = window.checkSubmit;

	window.fninit = function () {
		var issueToType;
		var issueToCode;
		ensureCompat();
		setDate("ctlRcptFromDate", field("hFromDate") ? field("hFromDate").value : "");
		setDate("ctlRcptToDate", field("hToDate") ? field("hToDate").value : "");
		issueToType = field("hIssueToType") ? field("hIssueToType").value : "";
		issueToCode = field("hIssueToCode") ? field("hIssueToCode").value : "";
		if (trim(issueToType) && trim(issueToCode)) {
			selectByValue("selIssueTo", trim(issueToType) + ":" + trim(issueToCode));
		} else if (trim(issueToType)) {
			selectByValue("selIssueTo", trim(issueToType));
		} else if (field("selIssueTo")) {
			field("selIssueTo").selectedIndex = 0;
		}
		if (field("hUser") && trim(field("hUser").value)) {
			selectByValue("selUser", field("hUser").value);
		} else if (field("selUser")) {
			field("selUser").selectedIndex = 0;
		}
		if (field("hStatus") && trim(field("hStatus").value)) {
			setRadioByValue("rStatus", field("hStatus").value);
		}
		if (dateControl("ctlRcptFromDate")) {
			dateControl("ctlRcptFromDate").disabled = true;
		}
		return false;
	};

	window.Clear = function () {
		window.clearTblItem();
		return false;
	};

	window.clearTblItem = function () {
		var table = document.getElementById("tblDetail");
		if (!table) {
			return false;
		}
		while (table.rows.length > 1) {
			table.deleteRow(1);
		}
		return false;
	};

	window.Validate = function () {
		var issueTo = selectedOption(field("selIssueTo")).value;
		var issueParts;
		if (trim(issueTo).toUpperCase() !== "SELECT" && trim(issueTo) !== "") {
			issueParts = issueTo.split(":");
			if (field("hIssueToType")) {
				field("hIssueToType").value = issueParts[0] || "";
			}
			if (field("hIssueToCode")) {
				field("hIssueToCode").value = issueParts.length > 1 ? issueParts[1] : "";
			}
		} else {
			if (field("hIssueToType")) {
				field("hIssueToType").value = "";
			}
			if (field("hIssueToCode")) {
				field("hIssueToCode").value = "";
			}
		}
		if (field("hStatus")) {
			field("hStatus").value = checkedRadioValue("rStatus") === "Y" ? "Y" : checkedRadioValue("rStatus") === "N" ? "N" : "A";
		}
		if (field("hUser")) {
			field("hUser").value = field("selUser") && field("selUser").selectedIndex > 0 ? selectedOption(field("selUser")).value : "";
		}
		if (field("hFromDate")) {
			field("hFromDate").value = getDate("ctlRcptFromDate");
		}
		if (field("hToDate")) {
			field("hToDate").value = getDate("ctlRcptToDate");
		}
		submitTo("IssueMGMT.asp?ACTN=L&ISSTYPE=" + encodeURIComponent(selectedOption(field("cmbIssType")).value));
		return false;
	};

	window.ResetData = function () {
		["hFromDate", "hToDate", "hStatus", "hUser", "hIssueToCode", "hIssueToSubCode", "hIssueToType"].forEach(function (name) {
			if (field(name)) {
				field(name).value = "";
			}
		});
		submitTo("IssueMGMT.asp");
		return false;
	};

	window.DisplayExistingCriteria = function () {
		return false;
	};

	window.DoAction = function (action) {
		var issue = requireIssue();
		if (!issue) {
			return false;
		}
		if (action === "C") {
			submitTo("MatConsDetailsEntry.asp?RefDet=" + encodeURIComponent(issue.tempValues));
		} else if (action === "R") {
			submitTo("IssReturnEntry.asp?RefDet=" + encodeURIComponent(issue.tempValues));
		}
		return false;
	};

	window.ViewIssDet = function () {
		var issue = requireIssue();
		if (!issue) {
			return false;
		}
		openDialog("ViewIssueEntDetail.asp?IssNo=" + encodeURIComponent(issue.issueNo), "", "dialogHeight:500px;dialogWidth:600px;Status:No");
		return false;
	};

	window.EditIssue = function () {
		var issue = requireIssue();
		if (!issue) {
			return false;
		}
		submitTo("MATERIALISSUEENTRY.ASP?ISSNO=" + encodeURIComponent(issue.issueNo) + "&TYPE=" + encodeURIComponent(issue.issueTypeCode));
		return false;
	};

	window.DeleteIssue = function () {
		var issue = requireIssue();
		if (!issue) {
			return false;
		}
		submitTo("ViewIssueEntDetail.asp?IssNo=" + encodeURIComponent(issue.issueNo) + "&CallFor=D");
		return false;
	};

	window.CancelIssue = function () {
		if (!requireIssue()) {
			return false;
		}
		alert("Issue cancel action is not configured for this page.");
		return false;
	};
}(window, document));
