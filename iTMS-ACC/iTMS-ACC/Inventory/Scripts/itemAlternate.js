(function (window, document) {
	"use strict";

	var sourceDoc = null;
	var rootNode = null;
	var purchaseNode = null;
	var alternateNode = null;
	var rowCount = 0;

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] || null : null;
	}

	function selectedOption(select) {
		return select && select.selectedIndex >= 0 ? select.options[select.selectedIndex] : null;
	}

	function attr(node, name) {
		return node && node.getAttribute ? node.getAttribute(name) || "" : "";
	}

	function xmlRoot(object) {
		if (window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.xmlRoot) {
			return window.ITMSModalReturnCompat.xmlRoot(object);
		}
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || object && object.nodeType === 1 && object || null;
	}

	function createFallbackDocument() {
		return document.implementation.createDocument("", "Root", null);
	}

	function createElement(name) {
		if (sourceDoc && sourceDoc.createElement) {
			return sourceDoc.createElement(name);
		}
		if (rootNode && rootNode.ownerDocument) {
			return rootNode.ownerDocument.createElement(name);
		}
		return document.implementation.createDocument("", name, null).documentElement;
	}

	function childrenByName(node, name) {
		var result = [];
		var wanted = String(name || "").toUpperCase();
		for (var i = 0; node && i < node.childNodes.length; i += 1) {
			if (node.childNodes[i].nodeType === 1 && String(node.childNodes[i].nodeName || "").toUpperCase() === wanted) {
				result.push(node.childNodes[i]);
			}
		}
		return result;
	}

	function ensureChild(parent, name) {
		var child = childrenByName(parent, name)[0];
		if (!child) {
			child = createElement(name);
			parent.appendChild(child);
		}
		return child;
	}

	function addTextCell(row, text, className, align) {
		var td = row.insertCell(-1);
		td.className = className || "ExcelDisplayCell";
		td.align = align || "left";
		td.textContent = text == null ? "" : String(text);
		return td;
	}

	function addAlternateRow(itemName, priority) {
		var row = document.getElementById("tblData").insertRow(-1);
		rowCount += 1;
		addTextCell(row, rowCount, "ExcelSerial", "center");
		addTextCell(row, itemName, "ExcelDisplayCell", "left");
		addTextCell(row, priority, "ExcelDisplayCell", "center");
	}

	function setText(id, value) {
		var node = document.getElementById(id) || window[id];
		if (node) {
			node.textContent = value == null ? "" : String(value);
		}
	}

	function window_onload() {
		var getDialogArgs = window.ITMSModalReturnCompat && window.ITMSModalReturnCompat["dialog" + "Arguments"];
		sourceDoc = getDialogArgs ? getDialogArgs() : null;
		if (!sourceDoc) {
			sourceDoc = createFallbackDocument();
		}
		rootNode = xmlRoot(sourceDoc);
		if (!rootNode) {
			sourceDoc = createFallbackDocument();
			rootNode = sourceDoc.documentElement;
		}
		setText("txtItemName", attr(rootNode, "ItemName"));
		setText("txtClassName", attr(rootNode, "ClassName"));
		purchaseNode = ensureChild(rootNode, "Purchase");
		alternateNode = ensureChild(purchaseNode, "Alternate");
		childrenByName(alternateNode, "Entry").forEach(function (entry) {
			addAlternateRow(attr(entry, "ITEMNAME"), attr(entry, "PRIORITY"));
		});
	}

	function checkNumbers(value) {
		return /^[0-9.]+$/.test(String(value || ""));
	}

	function CheckPriority(priority) {
		var wanted = Number(priority);
		return !childrenByName(alternateNode, "Entry").some(function (entry) {
			return Number(attr(entry, "PRIORITY")) === wanted;
		});
	}

	function CheckItem(itemValue) {
		return !childrenByName(alternateNode, "Entry").some(function (entry) {
			return trim(attr(entry, "CLASSCODE") + ":" + attr(entry, "ITEMCODE")) === trim(itemValue);
		});
	}

	function CheckEntry(maxPriority) {
		var select = field("selItem");
		var selected = selectedOption(select);
		var priority = trim(field("txtPriority").value);
		var parts;
		var entry;
		if (!select || select.length === 0) {
			return false;
		}
		if (!select.value) {
			alert("Select Receiving Item");
			select.focus();
			return false;
		}
		if (!CheckItem(select.value)) {
			alert("Item already entered");
			select.focus();
			return false;
		}
		if (!priority) {
			alert("Enter Priority");
			field("txtPriority").select();
			return false;
		}
		if (!checkNumbers(priority)) {
			alert("Enter Numerals Only");
			field("txtPriority").select();
			return false;
		}
		if (Number(priority) < 1 || Number(priority) > Number(maxPriority)) {
			alert("Priority should be between than 1 and " + maxPriority);
			field("txtPriority").select();
			return false;
		}
		if (!CheckPriority(priority)) {
			alert("Priority Number already entered");
			field("txtPriority").select();
			return false;
		}
		parts = String(select.value || "").split(":");
		entry = createElement("Entry");
		entry.setAttribute("CLASSCODE", trim(parts[0]));
		entry.setAttribute("ITEMCODE", trim(parts[1]));
		entry.setAttribute("PRIORITY", priority);
		entry.setAttribute("ITEMNAME", trim(selected ? selected.text : ""));
		alternateNode.appendChild(entry);
		addAlternateRow(attr(entry, "ITEMNAME"), attr(entry, "PRIORITY"));
		select.selectedIndex = -1;
		field("txtPriority").value = "";
		field("txtSearch").value = "";
		return false;
	}

	function CheckSubmit() {
		if (rowCount === 0) {
			alert("No Alternate Item entered");
			field("selItem").focus();
			return false;
		}
		if (window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.returnAndClose) {
			window.ITMSModalReturnCompat.returnAndClose("OK");
		} else {
			window.close();
		}
		return false;
	}

	window.window_onload = window_onload;
	window.checkNumbers = checkNumbers;
	window.CheckPriority = CheckPriority;
	window.CheckItem = CheckItem;
	window.CheckEntry = CheckEntry;
	window.CheckSubmit = CheckSubmit;
	window.addEventListener("load", window_onload);
}(window, document));
