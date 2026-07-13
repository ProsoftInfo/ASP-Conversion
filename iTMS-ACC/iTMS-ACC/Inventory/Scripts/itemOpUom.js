(function (window, document) {
	"use strict";

	var sourceDoc = null;
	var rootNode = null;
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

	function uomDetailsNode() {
		return childrenByName(rootNode, "UOMDETAILS")[0] || null;
	}

	function ensureUomDetailsNode() {
		var node = uomDetailsNode();
		if (!node) {
			node = createElement("UOMDETAILS");
			rootNode.appendChild(node);
		}
		updateHeaderAttributes(node);
		return node;
	}

	function setIndex(select, value) {
		if (!select) {
			return;
		}
		for (var i = 0; i < select.length; i += 1) {
			if (trim(select.options[i].value) === trim(value)) {
				select.selectedIndex = i;
				return;
			}
		}
	}

	function setValue(name, value) {
		var item = field(name);
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function table() {
		return document.getElementById("tblData");
	}

	function addTextCell(row, text, className, align) {
		var td = row.insertCell(-1);
		td.className = className || "ExcelDisplayCell";
		td.align = align || "left";
		td.textContent = text == null ? "" : String(text);
		return td;
	}

	function addUomRow(name, factor, operatorText, forText) {
		var row = table().insertRow(-1);
		rowCount += 1;
		addTextCell(row, rowCount, "ExcelSerial", "center");
		addTextCell(row, name, "ExcelDisplayCell", "left");
		addTextCell(row, factor, "ExcelDisplayCell", "left");
		addTextCell(row, operatorText, "ExcelDisplayCell", "center");
		addTextCell(row, forText, "ExcelDisplayCell", "center");
	}

	function updateHeaderAttributes(node) {
		node.setAttribute("PUR", field("selUoMPurchase").value);
		node.setAttribute("MAN", field("selUoMManu").value);
		node.setAttribute("SAL", field("selUoMSales").value);
		node.setAttribute("PURFAC", trim(field("txtStToPur").value));
		node.setAttribute("PUROPE", field("selStToPur").value);
		node.setAttribute("SALFAC", trim(field("txtStToSales").value));
		node.setAttribute("SALOPE", field("selStToSales").value);
		node.setAttribute("MANFAC", trim(field("txtStToManu").value));
		node.setAttribute("MANOPE", field("selStToManu").value);
	}

	function window_onload() {
		var getDialogArgs = window.ITMSModalReturnCompat && window.ITMSModalReturnCompat["dialog" + "Arguments"];
		var details;
		sourceDoc = getDialogArgs ? getDialogArgs() : null;
		if (!sourceDoc) {
			sourceDoc = createFallbackDocument();
		}
		rootNode = xmlRoot(sourceDoc);
		if (!rootNode) {
			sourceDoc = createFallbackDocument();
			rootNode = sourceDoc.documentElement;
		}
		details = uomDetailsNode();
		if (!details) {
			return;
		}
		setIndex(field("selUoMPurchase"), attr(details, "PUR"));
		setIndex(field("selUoMManu"), attr(details, "MAN"));
		setIndex(field("selUoMSales"), attr(details, "SAL"));
		setValue("txtStToPur", attr(details, "PURFAC"));
		setIndex(field("selStToPur"), attr(details, "PUROPE"));
		setValue("txtStToSales", attr(details, "SALFAC"));
		setIndex(field("selStToSales"), attr(details, "SALOPE"));
		setValue("txtStToManu", attr(details, "MANFAC"));
		setIndex(field("selStToManu"), attr(details, "MANOPE"));
		childrenByName(details, "OPUOMENTRY").forEach(function (node) {
			addUomRow(attr(node, "UNAME"), attr(node, "BRATE"), attr(node, "OPERATORTEXT"), attr(node, "FOR"));
		});
	}

	function checkNumbers(value) {
		return /^[0-9.]+$/.test(String(value || ""));
	}

	function CheckUoM(itemCode) {
		var details = uomDetailsNode();
		var selectedFor = selectedOption(field("selFor"));
		var selectedForText = selectedFor ? trim(selectedFor.text) : "";
		return !childrenByName(details, "OPUOMENTRY").some(function (node) {
			return trim(attr(node, "FOR")) === selectedForText && trim(attr(node, "UCODE")) === trim(itemCode);
		});
	}

	function CheckEntry() {
		var item = field("selItem");
		var selectedItem = selectedOption(item);
		var selectedOperator = selectedOption(field("selOpe"));
		var selectedFor = selectedOption(field("selFor"));
		var details;
		var node;
		if (!item || item.length === 0) {
			return false;
		}
		if (typeof window.CheckUoMEntered === "function" && !window.CheckUoMEntered()) {
			return false;
		}
		if (field("selFor").selectedIndex === 0) {
			alert("Select Alternate UoM For");
			field("selFor").focus();
			return false;
		}
		if (field("selFor").selectedIndex === 1 && field("selUoMPurchase").selectedIndex === 0) {
			alert("Select Purchase UoM");
			field("selUoMPurchase").focus();
			return false;
		}
		if (field("selFor").selectedIndex === 2 && field("selUoMSales").selectedIndex === 0) {
			alert("Select Sales UoM");
			field("selUoMSales").focus();
			return false;
		}
		if (!item.value) {
			alert("Select Alternate UoM");
			item.focus();
			return false;
		}
		if (field("selFor").selectedIndex === 1 && field("selUoMPurchase").value === item.value) {
			alert("Purchase UoM and alternate UoM can't be same");
			item.focus();
			return false;
		}
		if (field("selFor").selectedIndex === 2 && field("selUoMSales").value === item.value) {
			alert("Sales UoM and alternate UoM can't be same");
			item.focus();
			return false;
		}
		if (!CheckUoM(item.value)) {
			alert("Alternate UoM already entered");
			item.focus();
			return false;
		}
		if (!trim(field("txtFactor").value)) {
			alert("Enter Factor");
			field("txtFactor").select();
			return false;
		}
		if (!checkNumbers(field("txtFactor").value)) {
			alert("Enter Numerals Only");
			field("txtFactor").select();
			return false;
		}
		if (field("selOpe").value === "select") {
			alert("Select Operator");
			field("selOpe").focus();
			return false;
		}
		details = ensureUomDetailsNode();
		node = createElement("OPUOMENTRY");
		node.setAttribute("UCODE", item.value);
		node.setAttribute("BRATE", trim(field("txtFactor").value));
		node.setAttribute("OPERATOR", field("selOpe").value);
		node.setAttribute("UNAME", trim(selectedItem ? selectedItem.text : ""));
		node.setAttribute("OPERATORTEXT", trim(selectedOperator ? selectedOperator.text : ""));
		node.setAttribute("FOR", trim(selectedFor ? selectedFor.text : ""));
		details.appendChild(node);
		addUomRow(attr(node, "UNAME"), attr(node, "BRATE"), attr(node, "OPERATORTEXT"), attr(node, "FOR"));
		item.selectedIndex = -1;
		setValue("txtFactor", "");
		setValue("txtSearch", "");
		field("selOpe").selectedIndex = 0;
		return false;
	}

	function CheckSubmit() {
		var details;
		if (typeof window.CheckUoMEntered === "function" && !window.CheckUoMEntered()) {
			return false;
		}
		details = ensureUomDetailsNode();
		if (window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.returnAndClose) {
			window.ITMSModalReturnCompat.returnAndClose(rootNode || details);
		} else {
			window.close();
		}
		return false;
	}

	window.window_onload = window_onload;
	window.setIndex = setIndex;
	window.checkNumbers = checkNumbers;
	window.CheckUoM = CheckUoM;
	window.CheckEntry = CheckEntry;
	window.CheckSubmit = CheckSubmit;
	window.addEventListener("load", window_onload);
}(window, document));
