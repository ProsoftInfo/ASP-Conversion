(function (window, document) {
	"use strict";

	var sourceRoot = null;

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		var item;
		var wanted;
		var i;
		if (!frm || !frm.elements) {
			return null;
		}
		item = frm.elements[name];
		if (item) {
			return item;
		}
		wanted = String(name).toLowerCase();
		for (i = 0; i < frm.elements.length; i += 1) {
			if (String(frm.elements[i].name || "").toLowerCase() === wanted) {
				return frm.elements[i];
			}
		}
		return null;
	}

	function valueOf(name) {
		var item = field(name);
		return item ? item.value : "";
	}

	function setValue(name, value) {
		var item = field(name);
		if (item) {
			item.value = value == null ? "" : String(value);
		}
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

	function xmlDocument(name) {
		var object = xmlObject(name);
		var root = xmlRoot(object);
		return object && object.XMLDocument || object && object._doc || root && root.ownerDocument || null;
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

	function clearChildren(node) {
		while (node && node.firstChild) {
			node.removeChild(node.firstChild);
		}
	}

	function importNode(targetRoot, sourceNode) {
		var doc = targetRoot && targetRoot.ownerDocument;
		if (doc && doc.importNode) {
			return doc.importNode(sourceNode, true);
		}
		return sourceNode.cloneNode(true);
	}

	function dataRoot() {
		return xmlRoot("Data");
	}

	function createDataNode(name) {
		var doc = xmlDocument("Data") || document.implementation.createDocument("", "Root", null);
		return doc.createElement(name);
	}

	function createSourceRoot() {
		var doc = document.implementation.createDocument("", "Root", null);
		doc.documentElement.appendChild(doc.createElement("BOM"));
		return doc.documentElement;
	}

	function getDialogRoot() {
		var modal = window.ITMSModalReturnCompat;
		var getRoot = modal && modal["dialog" + "ArgumentsRoot"];
		return getRoot ? getRoot() : null;
	}

	function bomNode() {
		var node;
		if (!sourceRoot) {
			sourceRoot = getDialogRoot() || createSourceRoot();
		}
		node = childElements(sourceRoot, "BOM")[0];
		if (!node) {
			node = sourceRoot.ownerDocument.createElement("BOM");
			sourceRoot.appendChild(node);
		}
		return node;
	}

	function span(id) {
		return document.getElementById(id);
	}

	function setSpan(id, value) {
		var item = span(id);
		if (item) {
			item.textContent = value == null || value === "" ? "\u00a0" : String(value);
		}
	}

	function spanValue(id) {
		var item = span(id);
		return trim(item ? item.textContent || item.innerHTML : "");
	}

	function radioItems(name) {
		var items = form() && form().elements ? form().elements[name] : null;
		if (!items) {
			return [];
		}
		if (typeof items.length === "number" && !items.tagName) {
			return Array.prototype.slice.call(items);
		}
		return [items];
	}

	function radioValue(name) {
		var items = radioItems(name);
		for (var i = 0; i < items.length; i += 1) {
			if (items[i].checked) {
				return items[i].value;
			}
		}
		return "";
	}

	function setRadioValue(name, value) {
		radioItems(name).forEach(function (item) {
			item.checked = item.value === value;
		});
	}

	function currentTypeName() {
		return radioValue("radType") === "A" ? "Assembly" : "Final Component";
	}

	function currentConsumable() {
		var item = field("ChkConsumable") || field("chkConsumable");
		return item && item.checked ? "Y" : "N";
	}

	function setConsumable(value) {
		var item = field("ChkConsumable") || field("chkConsumable");
		if (item) {
			item.checked = trim(value) === "Y";
		}
	}

	function ClearTable() {
		var table = document.getElementById("tblData");
		if (!table) {
			return;
		}
		while (table.rows.length > 1) {
			table.deleteRow(1);
		}
	}

	function Reset() {
		setValue("hItemCode", "");
		setValue("hClassCode", "");
		setSpan("spSelItem", "");
		setValue("txtQty", "");
		setSpan("idUoM", "");
		setValue("hConsumable", "");
		setRadioValue("radType", "F");
		setConsumable("N");
	}

	function addTextCell(row, text, className, align) {
		var cell = row.insertCell(-1);
		cell.className = className || "ExcelDisplayCell";
		cell.align = align || "left";
		cell.textContent = text == null ? "" : String(text);
		return cell;
	}

	function addCheckboxCell(row, index, node) {
		var cell = addTextCell(row, "", "ExcelDisplayCell", "center");
		var checkbox = document.createElement("input");
		checkbox.type = "checkbox";
		checkbox.name = "chkItemZ" + index;
		checkbox.value = trim(attr(node, "ItemCode")) + "Z" + attr(node, "ClassCode");
		checkbox.className = "FormElem";
		cell.appendChild(checkbox);
	}

	function addItemLinkCell(row, node) {
		var cell = addTextCell(row, "", "ExcelDisplayCell", "left");
		var link = document.createElement("a");
		link.href = "#";
		link.className = "ExcelDisplayLink";
		link.textContent = attr(node, "ItemName");
		link.onclick = function () {
			EditItem(attr(node, "ItemCode"), attr(node, "ClassCode"));
			return false;
		};
		cell.appendChild(link);
	}

	function DisplayTable() {
		var root = dataRoot();
		var count = 0;
		var table = document.getElementById("tblData");
		if (!table) {
			return;
		}
		childElements(root, "Item").forEach(function (node) {
			var row = table.insertRow(-1);
			count += 1;
			addTextCell(row, count, "ExcelSerial", "center");
			addCheckboxCell(row, count, node);
			addItemLinkCell(row, node);
			addTextCell(row, trim(attr(node, "Qty") + " " + attr(node, "UoM")), "ExcelDisplayCell", "right");
			addTextCell(row, attr(node, "TypeName"), "ExcelDisplayCell", "center");
			addTextCell(row, trim(attr(node, "Consumable")) === "Y" ? "Yes" : "No", "ExcelDisplayCell", "center");
		});
		setValue("hRowCtr", count);
	}

	function findDataItem(itemCode) {
		var result = null;
		childElements(dataRoot(), "Item").some(function (node) {
			if (attr(node, "ItemCode") === itemCode) {
				result = node;
				return true;
			}
			return false;
		});
		return result;
	}

	function EditItem(itemCode) {
		var node = findDataItem(itemCode);
		if (!node) {
			return false;
		}
		setValue("hItemCode", attr(node, "ItemCode"));
		setValue("hClassCode", attr(node, "ClassCode"));
		setSpan("spSelItem", attr(node, "ItemName"));
		setValue("txtQty", attr(node, "Qty"));
		setSpan("idUoM", attr(node, "UoM"));
		setRadioValue("radType", attr(node, "Type") === "A" ? "A" : "F");
		setConsumable(attr(node, "Consumable"));
		return false;
	}

	function DelItem() {
		var root = dataRoot();
		var count = Number(valueOf("hRowCtr")) || 0;
		var selected = {};
		var item;
		var parts;
		for (var i = 1; i <= count; i += 1) {
			item = field("chkItemZ" + i);
			if (item && item.checked) {
				parts = String(item.value || "").split("Z");
				selected[parts[0]] = true;
			}
		}
		childElements(root, "Item").forEach(function (node) {
			if (selected[attr(node, "ItemCode")]) {
				root.removeChild(node);
			}
		});
		ClearTable();
		Reset();
		DisplayTable();
		return false;
	}

	function setItemAttributes(node) {
		var type = radioValue("radType") === "A" ? "A" : "F";
		node.setAttribute("ItemCode", valueOf("hItemCode"));
		node.setAttribute("ClassCode", valueOf("hClassCode"));
		node.setAttribute("ItemName", spanValue("spSelItem"));
		node.setAttribute("Qty", valueOf("txtQty"));
		node.setAttribute("UoM", spanValue("idUoM"));
		node.setAttribute("Type", type);
		node.setAttribute("TypeName", currentTypeName());
		node.setAttribute("Consumable", currentConsumable());
	}

	function CheckEntry() {
		var root = dataRoot();
		var node;
		if (!valueOf("hItemCode")) {
			alert("Select Item");
			return false;
		}
		if (!valueOf("txtQty")) {
			alert("Enter Quantity");
			field("txtQty").focus();
			return false;
		}
		node = findDataItem(valueOf("hItemCode"));
		if (!node) {
			node = createDataNode("Item");
			root.appendChild(node);
		}
		setItemAttributes(node);
		ClearTable();
		Reset();
		DisplayTable();
		return false;
	}

	function CheckSubmit() {
		var bom = bomNode();
		clearChildren(bom);
		childElements(dataRoot(), "Item").forEach(function (node) {
			bom.appendChild(importNode(bom, node));
		});
		if (window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.returnAndClose) {
			window.ITMSModalReturnCompat.returnAndClose(sourceRoot);
		} else {
			window.close();
		}
		return false;
	}

	function applySelectedItem() {
		var itemRoot = xmlRoot("ItemData");
		var duplicate;
		childElements(itemRoot).forEach(function (node) {
			setValue("hItemCode", attr(node, "ItemCode"));
			setValue("hClassCode", attr(node, "ClassCode"));
			setSpan("spSelItem", attr(node, "ItemName"));
			setSpan("idUoM", attr(node, "StoresUoM"));
		});
		duplicate = findDataItem(valueOf("hItemCode"));
		if (duplicate) {
			Reset();
			alert("This Item is Already Selected");
		}
	}

	function selectorSize() {
		var value = typeof window.GetWindowSizeForPopup === "function" ? window.GetWindowSizeForPopup("1") : "ItemSelectRelPartyCommon.asp:500:850";
		var parts = String(value || "ItemSelectRelPartyCommon.asp:500:850").split(":");
		return {
			program: parts[0] || "ItemSelectRelPartyCommon.asp",
			height: parts[1] || "500",
			width: parts[2] || "850"
		};
	}

	function continueSelection(returnedValue) {
		var root = xmlRoot(returnedValue);
		var size;
		var query;
		var action = String(attr(root, "Action")).toUpperCase();
		if (!root || action === "CLOSE") {
			return;
		}
		if (action && action !== "DONE") {
			query = attr(root, "PassQuery");
			if (query) {
				size = selectorSize();
				openDialog("../../Common/" + size.program + "?" + query, xmlObject("ItemData"), "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", continueSelection);
			}
			return;
		}
		applySelectedItem();
	}

	function openDialog(url, args, features, callback) {
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			return window.ITMSModernCompat.openModalDialog(url, args, features, callback);
		}
		window.open(url, "_blank");
		return null;
	}

	function SelectItem() {
		var size = selectorSize();
		var flag = window.nFlag == null ? "" : String(window.nFlag);
		var url = "../../Common/" + size.program +
			"?orgID=" + encodeURIComponent(valueOf("hOrgID")) +
			"&hSelectMode=R&sIType=" + encodeURIComponent(valueOf("hItemType")) +
			"&Stock=N&Flag=" + encodeURIComponent(flag);
		openDialog(url, xmlObject("ItemData"), "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", continueSelection);
		return false;
	}

	function Init() {
		var bom;
		var root = dataRoot();
		sourceRoot = getDialogRoot() || createSourceRoot();
		bom = bomNode();
		clearChildren(root);
		childElements(bom, "Item").forEach(function (node) {
			root.appendChild(importNode(root, node));
		});
		ClearTable();
		Reset();
		DisplayTable();
	}

	function windowOnUnload() {
		if (window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.returnValue) {
			window.ITMSModalReturnCompat.returnValue(sourceRoot || getDialogRoot() || dataRoot());
		}
		return sourceRoot || getDialogRoot() || dataRoot();
	}

	if (window.ITMSModalReturnCompat) {
		window.ITMSModalReturnCompat.install(function () {
			return sourceRoot || getDialogRoot() || dataRoot();
		});
	}

	window.Init = Init;
	window.EditItem = EditItem;
	window.DelItem = DelItem;
	window.CheckEntry = CheckEntry;
	window.CheckSubmit = CheckSubmit;
	window.ClearTable = ClearTable;
	window.Reset = Reset;
	window.DisplayTable = DisplayTable;
	window.SelectItem = SelectItem;
	window.window_onunload = windowOnUnload;
}(window, document));
