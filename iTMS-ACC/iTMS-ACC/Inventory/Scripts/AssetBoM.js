(function (window, document) {
	"use strict";

	var objTemp = null;
	var rootO = null;
	var rowNo = 0;

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function numberValue(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] : null;
	}

	function byId(id) {
		return document.getElementById(id);
	}

	function setText(id, value) {
		var item = byId(id);
		if (item) {
			item.textContent = value == null || value === "" ? "\u00a0" : String(value);
		}
	}

	function textOf(id) {
		var item = byId(id);
		return trim(item ? item.textContent || "" : "");
	}

	function selectedText(select) {
		if (!select || select.selectedIndex < 0 || !select.options[select.selectedIndex]) {
			return "";
		}
		return select.options[select.selectedIndex].text;
	}

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function xmlObject(name) {
		ensureCompat();
		return window[name] || document[name] || document.getElementById(name) || null;
	}

	function xmlRoot(nameOrObject) {
		var object = typeof nameOrObject === "string" ? xmlObject(nameOrObject) : nameOrObject;
		if (window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.xmlRoot) {
			return window.ITMSModalReturnCompat.xmlRoot(object);
		}
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || object && object._doc && object._doc.documentElement || object && object.nodeType === 1 && object || null;
	}

	function dialogDocument() {
		var modal = window.ITMSModalReturnCompat;
		var getArgs = modal && modal["dialog" + "Arguments"];
		var getRoot = modal && modal["dialog" + "ArgumentsRoot"];
		var args = getArgs ? getArgs() : null;
		var root = getRoot ? getRoot() : xmlRoot(args);
		if (args && args.XMLDocument) {
			return args.XMLDocument;
		}
		if (args && args._doc) {
			return args._doc;
		}
		if (args && args.nodeType === 9) {
			return args;
		}
		if (root && root.ownerDocument) {
			return root.ownerDocument;
		}
		return document.implementation.createDocument("", "Root", null);
	}

	function childElements(node, name) {
		var result = [];
		var wanted = name ? String(name).toUpperCase() : "";
		var child;
		for (var i = 0; node && i < node.childNodes.length; i += 1) {
			child = node.childNodes[i];
			if (child.nodeType === 1 && (!wanted || String(child.nodeName || "").toUpperCase() === wanted)) {
				result.push(child);
			}
		}
		return result;
	}

	function attr(node, name) {
		return node && node.getAttribute ? node.getAttribute(name) || "" : "";
	}

	function attrByIndex(node, index) {
		return node && node.attributes && node.attributes[index] ? node.attributes[index].value : "";
	}

	function assets() {
		return childElements(rootO, "ASSET");
	}

	function assetNode(classCode, create) {
		var nodes = assets();
		var node;
		for (var i = 0; i < nodes.length; i += 1) {
			if (trim(attr(nodes[i], "CLACODE")) === trim(classCode)) {
				return nodes[i];
			}
		}
		if (create) {
			node = objTemp.createElement("ASSET");
			node.setAttribute("CLACODE", classCode);
			rootO.appendChild(node);
			return node;
		}
		return null;
	}

	function itemDetails(asset) {
		return childElements(asset, "ITMDET");
	}

	function itemDetailNode(asset, itemCode) {
		var nodes = itemDetails(asset);
		for (var i = 0; i < nodes.length; i += 1) {
			if (trim(attr(nodes[i], "ITMCODE")) === trim(itemCode)) {
				return nodes[i];
			}
		}
		return null;
	}

	function ClearTable() {
		var table = byId("tblData");
		if (table) {
			while (table.rows.length > 1) {
				table.deleteRow(1);
			}
		}
		rowNo = 0;
	}

	function appendCell(row, text, className, align, width) {
		var cell = row.insertCell(-1);
		cell.textContent = text == null ? "" : String(text);
		cell.className = className || "ExcelDisplayCell";
		if (align) {
			cell.align = align;
		}
		if (width) {
			cell.width = width;
		}
		return cell;
	}

	function appendRow(node) {
		var table = byId("tblData");
		var row;
		if (!table) {
			return;
		}
		rowNo += 1;
		row = table.insertRow(-1);
		appendCell(row, rowNo, "ExcelSerial", "center");
		appendCell(row, trim(attr(node, "DESC")), "ExcelDisplayCell", "left");
		appendCell(row, trim(attr(node, "QTY")), "ExcelDisplayCell", "", "10");
		appendCell(row, trim(attr(node, "SUOM")), "ExcelDisplayCell");
		appendCell(row, trim(attr(node, "ITYPENAME")), "ExcelDisplayCell");
	}

	function renderRows() {
		ClearTable();
		assets().forEach(function (asset) {
			itemDetails(asset).forEach(appendRow);
		});
	}

	function fnInit() {
		objTemp = dialogDocument();
		rootO = objTemp.documentElement;
		if (window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.install) {
			window.ITMSModalReturnCompat.install(function () {
				return rootO;
			});
		}
		renderRows();
	}

	function loadXmlIntoIsland(name, url) {
		var xhr = new XMLHttpRequest();
		var object;
		var text;
		xhr.open("GET", url, false);
		xhr.send(null);
		text = xhr.responseText || "";
		if (!text && xhr.responseXML) {
			text = new XMLSerializer().serializeToString(xhr.responseXML);
		}
		if (text) {
			object = xmlObject(name);
			if (object && object.loadXML) {
				object.loadXML(text);
			} else if (object && object.LoadXML) {
				object.LoadXML(text);
			}
		}
		return text ? xmlRoot(name) : null;
	}

	function addOption(select, value, text) {
		var option = document.createElement("option");
		option.value = value;
		option.text = text;
		select.add(option);
	}

	function DisplayUoM(obj) {
		var root;
		var classCode;
		if (!field("selItem") || field("selItem").length === 0) {
			return false;
		}
		setText("idUoM", "");
		classCode = field("selClass").value;
		root = xmlRoot("ItemData");
		childElements(root, "ITEM").forEach(function (node) {
			if (trim(attr(node, "CLACODE")) === trim(classCode) && trim(attr(node, "ITMCODE")) === trim(obj.value)) {
				setText("idUoM", attr(node, "SUOM"));
			}
		});
		return false;
	}

	function GetItem(obj) {
		var root;
		field("selItem").options.length = 0;
		setText("idUoM", "");
		if (!obj || obj.value === "select") {
			return false;
		}
		root = loadXmlIntoIsland("ItemData", "XMLAssetItem.asp?sOrgID=" + encodeURIComponent(field("hOrgID").value) + "&sClass=" + encodeURIComponent(obj.value));
		if (!root || !childElements(root).length) {
			alert("No Item defined for the Classification Selected");
			obj.focus();
			return false;
		}
		childElements(root).forEach(function (node) {
			addOption(field("selItem"), attrByIndex(node, 1), attrByIndex(node, 2));
		});
		DisplayUoM(field("selItem"));
		return false;
	}

	function radioItems() {
		var item = field("radType");
		if (!item) {
			return [];
		}
		if (typeof item.length === "number" && !item.tagName) {
			return Array.prototype.slice.call(item);
		}
		return [item];
	}

	function selectedType() {
		var radios = radioItems();
		for (var i = 0; i < radios.length; i += 1) {
			if (radios[i].checked) {
				return radios[i].value;
			}
		}
		return "";
	}

	function clearType() {
		radioItems().forEach(function (item) {
			item.checked = false;
		});
	}

	function typeName(value) {
		return value === "F" ? "Final Component" : "Assembly";
	}

	function validateEntry() {
		if (field("selClass").selectedIndex === 0) {
			alert("Select Classification");
			field("selClass").focus();
			return false;
		}
		if (!field("selItem").value) {
			alert("Select Item");
			field("selItem").focus();
			return false;
		}
		if (!trim(field("txtQty").value)) {
			alert("Enter Quantity");
			field("txtQty").select();
			return false;
		}
		if (isNaN(parseFloat(field("txtQty").value))) {
			alert("Enter Quantity in numerals");
			field("txtQty").select();
			return false;
		}
		if (numberValue(field("txtQty").value) <= 0) {
			alert("Quantity cannot be ZERO");
			field("txtQty").select();
			return false;
		}
		if (!selectedType()) {
			alert("Select Type");
			radioItems()[0].focus();
			return false;
		}
		return true;
	}

	function CheckEntry() {
		var classCode;
		var itemCode;
		var asset;
		var itemNode;
		var type;
		if (!validateEntry()) {
			return false;
		}
		classCode = field("selClass").value;
		itemCode = field("selItem").value;
		type = selectedType();
		asset = assetNode(classCode, true);
		itemNode = itemDetailNode(asset, itemCode);
		if (!itemNode) {
			itemNode = objTemp.createElement("ITMDET");
			asset.appendChild(itemNode);
		}
		itemNode.setAttribute("ITMCODE", itemCode);
		itemNode.setAttribute("DESC", selectedText(field("selItem")) + " / " + selectedText(field("selClass")));
		itemNode.setAttribute("QTY", trim(field("txtQty").value));
		itemNode.setAttribute("SUOM", textOf("idUoM"));
		itemNode.setAttribute("ITYPE", type);
		itemNode.setAttribute("ITYPENAME", typeName(type));
		renderRows();
		field("selClass").selectedIndex = 0;
		field("selItem").options.length = 0;
		field("txtQty").value = "";
		setText("idUoM", "");
		clearType();
		return false;
	}

	function CheckSubmit() {
		if (rowNo === 0) {
			alert("Enter BoM details");
			return false;
		}
		if (window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.returnAndClose) {
			window.ITMSModalReturnCompat.returnAndClose(rootO);
		} else {
			window.close();
		}
		return false;
	}

	window.fnInit = fnInit;
	window.DisplayUoM = DisplayUoM;
	window.GetItem = GetItem;
	window.CheckEntry = CheckEntry;
	window.CheckSubmit = CheckSubmit;
	window.ClearTable = ClearTable;
}(window, document));
