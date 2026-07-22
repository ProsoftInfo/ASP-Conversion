(function (window, document) {
	"use strict";

	var root = null;
	var rootDoc = null;
	var rows = [];

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function form() {
		return document.forms.formname || document.forms[0] || null;
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] || null : null;
	}

	function byId(id) {
		return document.getElementById(id) || window[id] || null;
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function xmlDocument(value) {
		ensureCompat();
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
		return value.documentElement || value.XMLDocument && value.XMLDocument.documentElement || value._doc && value._doc.documentElement || (value.nodeType === 1 ? value : null);
	}

	function modalArgs() {
		ensureCompat();
		return window["dialog" + "Arguments"] || null;
	}

	function elementChildren(node, name) {
		var result = [];
		var wanted = name && String(name).toLowerCase();
		var children = node && node.childNodes || [];
		for (var i = 0; i < children.length; i += 1) {
			if (children[i].nodeType === 1 && (!wanted || String(children[i].nodeName).toLowerCase() === wanted)) {
				result.push(children[i]);
			}
		}
		return result;
	}

	function getAttr(node, name) {
		return trim(node && node.getAttribute ? node.getAttribute(name) : "");
	}

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, String(value == null ? "" : value));
		}
	}

	function appendCell(row, className, align, content) {
		var cell = row.insertCell(row.cells.length);
		cell.className = className || "";
		if (align) {
			cell.align = align;
		}
		if (content == null) {
			return cell;
		}
		if (typeof content === "string" || typeof content === "number") {
			cell.textContent = String(content);
		} else {
			cell.appendChild(content);
		}
		return cell;
	}

	function makeCheckbox(name, value, checked) {
		var input = document.createElement("input");
		input.type = "checkbox";
		input.name = name;
		input.value = value == null ? "" : String(value);
		input.className = "Formelem";
		input.checked = !!checked;
		return input;
	}

	function isBaseItemType() {
		var value = trim(field("hIType") && field("hIType").value);
		return value === "FAB" || value === "YRN";
	}

	function clearTable() {
		var table = byId("tblBin");
		rows = [];
		if (!table) {
			return;
		}
		while (table.rows.length) {
			table.deleteRow(0);
		}
	}

	function renderHeader(table) {
		var row = table.insertRow(0);
		appendCell(row, "ExcelHeaderCell", "center", "S.No.");
		appendCell(row, "ExcelHeaderCell", "center", "");
		if (isBaseItemType()) {
			appendCell(row, "ExcelHeaderCell", "center", "Item Name");
			appendCell(row, "ExcelHeaderCell", "center", "Quantity");
		} else {
			appendCell(row, "ExcelHeaderCell", "center", "Lot Number");
		}
	}

	function shouldDisplay(node, optValue) {
		return getAttr(node, "Selection") === "N" || getAttr(node, "Selection") === "Y" && getAttr(node, "OptValue") === optValue;
	}

	function rowValue(node) {
		return String(node.nodeName).toLowerCase() === "baseitem" ? getAttr(node, "ICode") : getAttr(node, "No");
	}

	function renderNode(table, node, rowNo, optValue) {
		var row = table.insertRow(table.rows.length);
		var checkbox = makeCheckbox("Chk" + rowNo, rowValue(node), getAttr(node, "Selection") === "Y" && getAttr(node, "OptValue") === optValue);
		appendCell(row, "ExcelDisplayCell", "center", rowNo);
		appendCell(row, "ExcelDisplayCell", "center", checkbox);
		if (String(node.nodeName).toLowerCase() === "baseitem") {
			appendCell(row, "ExcelDisplayCell", "left", getAttr(node, "Desc"));
			appendCell(row, "ExcelDisplayCell", "left", getAttr(node, "Qty"));
		} else {
			appendCell(row, "ExcelDisplayCell", "center", getAttr(node, "No"));
		}
		rows.push({ node: node, checkbox: checkbox });
	}

	function returnDialogValue() {
		if (!root) {
			return;
		}
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(root);
		} else {
			window["return" + "Value"] = root;
			window.returnvalue = root;
		}
	}

	function initializeDialogXml() {
		var args = modalArgs();
		root = xmlRoot(args);
		rootDoc = xmlDocument(args) || root && root.ownerDocument;
		if (!root) {
			rootDoc = new DOMParser().parseFromString("<Root/>", "text/xml");
			root = rootDoc.documentElement;
		}
	}

	window.DisplayData = function () {
		var table = byId("tblBin");
		var optValue = trim(field("hOptValue") && field("hOptValue").value);
		var rowNo = 0;
		clearTable();
		if (!table) {
			return false;
		}
		renderHeader(table);
		elementChildren(root, "Item").forEach(function (item) {
			elementChildren(item).forEach(function (node) {
				var name = String(node.nodeName).toLowerCase();
				if ((name === "lot" || name === "baseitem") && shouldDisplay(node, optValue)) {
					rowNo += 1;
					renderNode(table, node, rowNo, optValue);
				}
			});
		});
		if (field("hCnt")) {
			field("hCnt").value = rows.length;
		}
		return false;
	};

	window.ClearTable = function () {
		clearTable();
		return false;
	};

	window.FnInit = function () {
		initializeDialogXml();
		window.DisplayData();
		return false;
	};

	window.CheckSubmit = function () {
		var optValue = trim(field("hOptValue") && field("hOptValue").value);
		rows.forEach(function (row) {
			if (row.checkbox.checked) {
				setAttr(row.node, "OptValue", optValue);
				setAttr(row.node, "Selection", "Y");
			} else {
				setAttr(row.node, "OptValue", "");
				setAttr(row.node, "Selection", "N");
			}
		});
		returnDialogValue();
		window.close();
		return false;
	};

	window.checkSubmit = window.CheckSubmit;
	window.window_onunload = returnDialogValue;
	window.addEventListener("beforeunload", returnDialogValue);
}(window, document));
