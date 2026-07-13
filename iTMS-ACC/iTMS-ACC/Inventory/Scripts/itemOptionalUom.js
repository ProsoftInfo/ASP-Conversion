(function (window, document) {
	"use strict";

	var sourceDoc = null;
	var rootNode = null;
	var optionalNode = null;
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

	function addOptionalRow(name, factor, operatorText) {
		var row = document.getElementById("tblData").insertRow(-1);
		rowCount += 1;
		addTextCell(row, rowCount, "ExcelSerial", "center");
		addTextCell(row, name, "ExcelDisplayCell", "left");
		addTextCell(row, factor, "ExcelDisplayCell", "left");
		addTextCell(row, operatorText, "ExcelDisplayCell", "center");
	}

	function setText(id, value) {
		var node = document.getElementById(id) || window[id];
		if (node) {
			node.textContent = value == null ? "" : String(value);
		}
	}

	function selectedControlNode() {
		var callFrom = trim(field("hCallFrom") && field("hCallFrom").value).toUpperCase();
		return ensureChild(rootNode, callFrom === "PUR" ? "Purchase" : "Sales");
	}

	function window_onload() {
		var getDialogArgs = window.ITMSModalReturnCompat && window.ITMSModalReturnCompat["dialog" + "Arguments"];
		var controlNode;
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
		controlNode = selectedControlNode();
		optionalNode = ensureChild(controlNode, "OptionalUOM");
		childrenByName(optionalNode, "OpUoMEntry").forEach(function (entry) {
			addOptionalRow(attr(entry, "UNAME"), attr(entry, "BRATE"), attr(entry, "OPERATORTEXT"));
		});
	}

	function checkNumbers(value) {
		return /^[0-9.]+$/.test(String(value || ""));
	}

	function CheckUoM(itemCode) {
		return !childrenByName(optionalNode, "OpUoMEntry").some(function (entry) {
			return trim(attr(entry, "UCODE")) === trim(itemCode);
		});
	}

	function CheckEntry() {
		var select = field("selItem");
		var selected = selectedOption(select);
		var selectedOperator = selectedOption(field("selOpe"));
		var factor = trim(field("txtFactor").value);
		var entry;
		if (!select || select.length === 0) {
			return false;
		}
		if (!select.value) {
			alert("Select Alternate UoM");
			select.focus();
			return false;
		}
		if (!CheckUoM(select.value)) {
			alert("Alternate UoM already entered");
			select.focus();
			return false;
		}
		if (!factor) {
			alert("Enter Factor");
			field("txtFactor").select();
			return false;
		}
		if (!checkNumbers(factor)) {
			alert("Enter Numerals Only");
			field("txtFactor").select();
			return false;
		}
		if (field("selOpe").value === "select") {
			alert("Select Operator");
			field("selOpe").focus();
			return false;
		}
		entry = createElement("OpUoMEntry");
		entry.setAttribute("UCODE", trim(select.value));
		entry.setAttribute("BRATE", factor);
		entry.setAttribute("OPERATOR", trim(field("selOpe").value));
		entry.setAttribute("UNAME", trim(selected ? selected.text : ""));
		entry.setAttribute("OPERATORTEXT", trim(selectedOperator ? selectedOperator.text : ""));
		optionalNode.appendChild(entry);
		addOptionalRow(attr(entry, "UNAME"), attr(entry, "BRATE"), attr(entry, "OPERATORTEXT"));
		select.selectedIndex = -1;
		field("txtFactor").value = "";
		field("txtSearch").value = "";
		field("selOpe").selectedIndex = 0;
		return false;
	}

	function CheckSubmit() {
		if (rowCount === 0) {
			alert("No Optional UoM entered");
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
	window.CheckUoM = CheckUoM;
	window.CheckEntry = CheckEntry;
	window.CheckSubmit = CheckSubmit;
	window.addEventListener("load", window_onload);
}(window, document));
