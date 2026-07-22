(function (window, document) {
	"use strict";

	var rowIndex = 0;

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function upper(value) {
		return trim(value).toUpperCase();
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

	function focusField(name) {
		var element = field(name);
		if (element && element.focus) {
			element.focus();
		}
	}

	function xmlRoot(value, fallback) {
		var candidate = value || fallback;
		if (!candidate) {
			return null;
		}
		if (candidate.documentElement) {
			return candidate.documentElement;
		}
		if (candidate.XMLDocument && candidate.XMLDocument.documentElement) {
			return candidate.XMLDocument.documentElement;
		}
		if (candidate._doc && candidate._doc.documentElement) {
			return candidate._doc.documentElement;
		}
		if (candidate.nodeType === 1) {
			return candidate;
		}
		return null;
	}

	function childElements(node) {
		return Array.prototype.filter.call(node && node.childNodes || [], function (child) {
			return child.nodeType === 1;
		});
	}

	function getAttr(node, name) {
		return node && node.getAttribute ? trim(node.getAttribute(name)) : "";
	}

	function openDialog(url, args, features, callback) {
		if (!window.ITMSModernCompat || !window.ITMSModernCompat.openModalDialog) {
			alert("Modern browser compatibility script is still loading. Please try again.");
			return;
		}
		window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
	}

	function submitTo(url) {
		var frm = form();
		if (!frm) {
			return;
		}
		frm.action = url;
		frm.submit();
	}

	function setDateControlValue(name, value) {
		var element = field(name);
		if (!element) {
			return;
		}
		if (typeof element.SetDate === "function") {
			element.SetDate(value);
		} else if (typeof element.setDate === "function") {
			element.setDate(value);
		} else if (window.ITMSModernCompat && window.ITMSModernCompat.toIsoDate) {
			element.value = window.ITMSModernCompat.toIsoDate(value);
		} else {
			element.value = value;
		}
	}

	function getDateControlValue(name) {
		var element = field(name);
		if (!element) {
			return "";
		}
		if (typeof element.getDate === "function") {
			return element.getDate();
		}
		if (typeof element.GetDate === "function") {
			return element.GetDate();
		}
		if (window.ITMSModernCompat && window.ITMSModernCompat.toDisplayDate) {
			return window.ITMSModernCompat.toDisplayDate(element.value);
		}
		return trim(element.value);
	}

	function setDisabled(name, disabled) {
		var element = field(name);
		if (element) {
			element.disabled = !!disabled;
		}
	}

	function setRadioValue(name, value) {
		var element = field(name);
		var i;
		if (!element) {
			return;
		}
		if (typeof element.length === "number" && !element.tagName) {
			for (i = 0; i < element.length; i += 1) {
				element[i].checked = trim(element[i].value) === trim(value);
			}
			return;
		}
		element.checked = trim(element.value) === trim(value);
	}

	function tableElement(id) {
		return document.getElementById(id) || window[id];
	}

	function processSupplierSelection(result) {
		var root = xmlRoot(result, window.OutData);
		if (!root || !root.hasChildNodes()) {
			return;
		}
		childElements(root).forEach(function (node) {
			if (upper(node.nodeName) === "SUPPLIER") {
				setValue("txtRefName", getAttr(node, "SuppName"));
				setValue("txtSupplier", getAttr(node, "SuppName"));
				setValue("hSupplierCode", getAttr(node, "SuppCode"));
			}
		});
	}

	function continueSupplierDialog(result) {
		var action = upper(getAttr(result, "Action"));
		var passQuery = getAttr(result, "PassQuery");
		if (action && action !== "DONE" && action !== "CLOSE") {
			openDialog("SupplierSelect.asp?" + passQuery, window.OutData, "status:no", continueSupplierDialog);
			return;
		}
		if (action !== "CLOSE") {
			processSupplierSelection(result);
		}
	}

	window.btnServiceBill_Click = function () {
		submitTo("GATEPASSSERVICEENTRY.ASP?SerType=I");
	};

	window.popSuppAgent = function () {
		var unitField = field("cmbUnit");
		var unit = trim(unitField && unitField.value) + ":O";
		openDialog("SupplierSelect.asp?OrgId=" + encodeURIComponent(unit) + "&hSelectMode=S&Flag=2&OrderTo=S", window.OutData, "status:no", continueSupplierDialog);
	};

	window.clearXML = function () {
		var root = window.OutData && window.OutData.documentElement;
		while (root && root.firstChild) {
			root.removeChild(root.firstChild);
		}
	};

	window.ClearTable = function () {
		var table = tableElement("tblDetails");
		if (!table) {
			return;
		}
		while (table.rows.length > 1) {
			table.deleteRow(1);
		}
		rowIndex = 0;
	};

	window.InvoiceDetails_onClick = function (gatePassNo, orgId) {
		submitTo("InvComeFromGatePass.asp?iGPNo=" + encodeURIComponent(gatePassNo) + "&ForUnit=" + encodeURIComponent(orgId));
	};

	window.Check = function (gatePassNo) {
		var unit = trim(field("hUnit") && field("hUnit").value);
		openDialog(
			"GatePassServiceAccountPop.asp?iGPNo=" + encodeURIComponent(gatePassNo) + "&ForUnit=" + encodeURIComponent(unit),
			"",
			"dialogHeight:350px;dialogWidth:600px;help:no;status:no",
			function (result) {
				var root = xmlRoot(result);
				var rateType = "";
				if (root && root.hasChildNodes()) {
					childElements(root).forEach(function (node) {
						if (node.nodeName === "WithInv") {
							rateType = getAttr(node, "Rate");
						}
					});
					submitTo("InvComeFromGatePass.asp?iGPNo=" + encodeURIComponent(gatePassNo) + "&ForUnit=" + encodeURIComponent(unit) + "&RateType=" + encodeURIComponent(rateType));
				} else {
					submitTo("GatePassServiceAccountEntry.asp");
				}
			}
		);
	};

	window.AssignPage = function (pageNo) {
		setValue("hPage", pageNo);
		submitTo("GatePassServiceAccountEntry.asp");
	};

	window.Validate = function () {
		var unit = field("cmbUnit");
		if (unit && unit.value === "0") {
			alert("Select Unit");
			focusField("cmbUnit");
			return;
		}
		setValue("hFromDate", getDateControlValue("ctlRcptFromDate"));
		setValue("hToDate", getDateControlValue("ctlRcptToDate"));
		setValue("hItemTypeName", "");
		submitTo("GatePassServiceAccountEntry.asp");
	};

	window.fninit = function () {
		var unit = field("cmbUnit");
		var hiddenUnit = trim(field("hUnit") && field("hUnit").value);
		var i;
		if (unit) {
			for (i = 0; i < unit.options.length; i += 1) {
				if (trim(unit.options[i].value) === hiddenUnit) {
					unit.selectedIndex = i;
					break;
				}
			}
		}
		if (trim(field("hMonthDate") && field("hMonthDate").value) === "D") {
			setRadioValue("rMonthDate", "D");
			setDateControlValue("ctlRcptFromDate", field("hFromDate") && field("hFromDate").value);
			setDateControlValue("ctlRcptToDate", field("hToDate") && field("hToDate").value);
		}
		setValue("txtRefName", field("txtSupplier") && field("txtSupplier").value);
		setDisabled("ctlRcptFromDate", true);
		setDisabled("ctlRcptToDate", true);
	};

	window.DisplayExistingCriteria = function () {};

	window.gotoPage = function () {
		submitTo("GatePassServiceAccountEntry.asp");
	};

	window.ResetForm = function () {
		var unit = field("cmbUnit");
		if (unit && String(unit.selectedIndex) === "0") {
			return;
		}
		submitTo("GatePassServiceAccountEntry.asp");
	};

	window.ResetData = window.ResetForm;
})(window, document);
