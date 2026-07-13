(function (window, document) {
	"use strict";

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function xmlObject(name) {
		ensureCompat();
		return window[name] || document[name] || document.getElementById(name) || null;
	}

	function xmlRoot(name) {
		var object = xmlObject(name);
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || object && object._doc && object._doc.documentElement || null;
	}

	function childElements(node) {
		var result = [];
		for (var i = 0; node && i < node.childNodes.length; i += 1) {
			if (node.childNodes[i].nodeType === 1) {
				result.push(node.childNodes[i]);
			}
		}
		return result;
	}

	function attrByIndex(node, index) {
		return node && node.attributes && node.attributes[index] ? node.attributes[index].value : "";
	}

	function table() {
		return document.getElementById("tblLoc");
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function hOrgCode() {
		var frm = form();
		return frm && frm.elements.hOrgCode ? frm.elements.hOrgCode.value : "";
	}

	function selectedOrg(select) {
		if (select && select.selectedIndex > 0 && select.options[select.selectedIndex]) {
			return select.options[select.selectedIndex].value;
		}
		return hOrgCode();
	}

	function appCodes() {
		var codes = window.ITMS_STORAGE_APPS || window.ITMS_STORAGE_APP || "PU";
		if (typeof codes === "string") {
			codes = codes.split(",");
		}
		return codes.map(function (code) {
			return String(code).replace(/^\s+|\s+$/g, "");
		});
	}

	function allowed(appCode) {
		return appCodes().indexOf(appCode) !== -1;
	}

	function ClearTable() {
		var tbl = table();
		if (!tbl) {
			return;
		}
		while (tbl.rows.length > 1) {
			tbl.deleteRow(1);
		}
	}

	function addRow(serial, text) {
		var tbl = table();
		var row;
		var cell;
		if (!tbl) {
			return;
		}
		row = tbl.insertRow(-1);
		cell = row.insertCell(-1);
		cell.textContent = serial || "";
		cell.className = "ExcelSerial";
		cell.align = "center";
		cell = row.insertCell(-1);
		cell.textContent = text || "";
		cell.className = "ExcelDisplayCell";
		cell.align = "left";
	}

	function DisplayStoDet(select) {
		var root = xmlRoot("storageData");
		var orgCode = selectedOrg(select);
		var found = false;
		var serial = 1;
		ClearTable();
		if (!orgCode) {
			return false;
		}
		childElements(root).some(function (header) {
			if (attrByIndex(header, 0) !== orgCode) {
				return false;
			}
			childElements(header).forEach(function (storeNode) {
				if (allowed(attrByIndex(storeNode, 3))) {
					addRow(serial, attrByIndex(storeNode, 2));
					serial += 1;
					found = true;
				}
			});
			if (!found) {
				addRow("", "No Storage Defined");
			}
			return true;
		});
		return false;
	}

	function ClearAll() {
		ClearTable();
		if (form()) {
			form().reset();
		}
	}

	window.DisplayStoDet = DisplayStoDet;
	window.ClearTable = ClearTable;
	window.ClearAll = ClearAll;
}(window, document));
