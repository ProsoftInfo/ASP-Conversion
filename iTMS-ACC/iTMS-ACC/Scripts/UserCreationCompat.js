(function (window, document) {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.formname || document.forms.formname || document.forms[0] || null;
	}

	function field(name) {
		var frm = form();
		return frm && (frm.elements[name] || frm[name]) || document.getElementById(name) || null;
	}

	function setValue(name, value) {
		var item = field(name);
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function attr(node, name) {
		return node && node.getAttribute ? node.getAttribute(name) || "" : "";
	}

	function childElements(node, name) {
		var wanted = name && String(name).toLowerCase();
		var result = [];
		for (var i = 0; node && i < node.childNodes.length; i += 1) {
			if (node.childNodes[i].nodeType === 1 && (!wanted || String(node.childNodes[i].nodeName).toLowerCase() === wanted)) {
				result.push(node.childNodes[i]);
			}
		}
		return result;
	}

	function getReturnedRoot(value) {
		return value && (value.documentElement || value.XMLDocument && value.XMLDocument.documentElement || value._doc && value._doc.documentElement || value) || null;
	}

	function parseXml(text) {
		if (!trim(text)) {
			return null;
		}
		return new DOMParser().parseFromString(text, "text/xml").documentElement;
	}

	function requestXml(url) {
		var xhr = new XMLHttpRequest();
		xhr.open("GET", url, false);
		xhr.send(null);
		return xhr.responseXML && xhr.responseXML.documentElement || parseXml(xhr.responseText || "");
	}

	function popupConfig(accessMode) {
		if (trim(accessMode).toUpperCase() === "I") {
			return {
				program: "EmpSelPop.asp",
				height: 500,
				width: 500,
				internal: true
			};
		}
		return {
			program: "PartySelPop.asp",
			height: 500,
			width: 500,
			internal: false
		};
	}

	function openSelectionDialog(config, query, callback) {
		var url = "../../Common/" + config.program + (trim(query) ? "?" + query : "");
		var features = "dialogHeight:" + config.height + "px;dialogWidth:" + config.width + "px;Status:No";
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, "", features, callback);
		} else {
			alert("Popup support is not available.");
		}
	}

	function runSelectionDialog(config, query, callback) {
		openSelectionDialog(config, query, function (returnedValue) {
			var root = getReturnedRoot(returnedValue);
			var action = trim(attr(root, "Action")).toUpperCase();
			var passQuery;
			if (!root || action === "CLOSE") {
				return;
			}
			if (action && action !== "DONE") {
				passQuery = attr(root, "PassQuery");
				runSelectionDialog(config, passQuery, callback);
				return;
			}
			callback(root);
		});
	}

	function fillInternalUser(entry) {
		var empId = attr(entry, "RetField1");
		var empCode = attr(entry, "RetField2");
		var nameParts = attr(entry, "RetField3").split("-");
		var addressRoot;
		setValue("hCode", empId);
		setValue("txtEmployeeID", empCode);
		setValue("txtFName", nameParts[0] || "");
		setValue("txtMName", nameParts[1] || "");
		setValue("txtLName", nameParts[2] || "");
		addressRoot = requestXml("../../Common/XMLGetEmpAddress.asp?EMPID=" + encodeURIComponent(empId));
		childElements(addressRoot, "EmpAdd").some(function (node) {
			setValue("txtStreet", attr(node, "Add1"));
			setValue("txtAddr2", attr(node, "Add2"));
			setValue("txtCity", attr(node, "City"));
			setValue("txtState", attr(node, "State"));
			setValue("txtPostal", attr(node, "Postal"));
			setValue("txtWorkPhone", attr(node, "WorkPhone"));
			setValue("txtWorkEmail", attr(node, "WorkEmail"));
			setValue("txtCell", attr(node, "MobilePhone"));
			return true;
		});
	}

	function fillExternalUser(entry) {
		var partyCode = attr(entry, "RetField1");
		var partyDisplayCode = attr(entry, "RetField2");
		var partyName = attr(entry, "RetField0");
		var addressRoot;
		setValue("txtEmployeeID", partyDisplayCode);
		setValue("txtFName", partyName);
		setValue("hCode", partyCode);
		addressRoot = requestXml("../../Include/XMLGetPartyAddress.asp?PartyCode=" + encodeURIComponent(partyCode));
		childElements(addressRoot, "Party").some(function (node) {
			setValue("txtStreet", attr(node, "Address1"));
			setValue("txtAddr2", attr(node, "Address2"));
			setValue("txtCity", attr(node, "City"));
			setValue("txtState", attr(node, "State"));
			setValue("txtWorkPhone", attr(node, "PhoneNo"));
			setValue("txtCell", attr(node, "MobileNo"));
			return true;
		});
	}

	function selectEmployee() {
		var config = popupConfig(field("hUAM") && field("hUAM").value);
		runSelectionDialog(config, "", function (root) {
			childElements(root).forEach(function (entry) {
				if (config.internal) {
					fillInternalUser(entry);
				} else {
					fillExternalUser(entry);
				}
			});
		});
	}

	function setIndex(select, value) {
		var target = trim(value);
		if (!select) {
			return false;
		}
		for (var i = 0; i < select.options.length; i += 1) {
			if (trim(select.options[i].value) === target) {
				select.selectedIndex = i;
				return true;
			}
		}
		return false;
	}

	function setUserType() {
		var radios = field("radUserType");
		var type = trim(field("hUserType") && field("hUserType").value);
		if (!radios || !type) {
			return false;
		}
		for (var i = 0; i < radios.length; i += 1) {
			radios[i].checked = trim(radios[i].value) === type;
		}
		return true;
	}

	window.setIndex = setIndex;
	window.SetIndex = setIndex;
	window.SetUserType = setUserType;
	window.SelectEmployee = selectEmployee;
}(window, document));
