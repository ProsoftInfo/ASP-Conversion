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

	function ClearTable() {
		var tbl = table();
		if (!tbl) {
			return;
		}
		while (tbl.rows.length > 1) {
			tbl.deleteRow(1);
		}
	}

	function DisplayStoDet(select) {
		var root = xmlRoot("storageData");
		var selected = select && select.selectedIndex > 0 ? select.options[select.selectedIndex].value : "";
		var serial = 1;
		var found = false;
		ClearTable();
		if (!selected) {
			return false;
		}
		childElements(root).some(function (header) {
			if (attrByIndex(header, 0) !== selected) {
				return false;
			}
			childElements(header).forEach(function (storeNode) {
				if (attrByIndex(storeNode, 3) === "PU") {
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
		if (document.forms.formname) {
			document.forms.formname.reset();
		}
	}

	window.DisplayStoDet = DisplayStoDet;
	window.ClearTable = ClearTable;
	window.ClearAll = ClearAll;
}(window, document));
