(function (window, document) {
	"use strict";

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

	function valueOf(name) {
		var item = field(name);
		return item ? item.value || "" : "";
	}

	function setValue(name, value) {
		var item = field(name);
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function buildQuery(params) {
		var pairs = [];
		Object.keys(params).forEach(function (name) {
			if (params[name] !== undefined && params[name] !== null && params[name] !== "") {
				pairs.push(encodeURIComponent(name) + "=" + encodeURIComponent(params[name]));
			}
		});
		return pairs.join("&");
	}

	function xmlIsland(name) {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
		return window[name] || document[name] || document.getElementById(name) || null;
	}

	function xmlRoot(name) {
		var island = xmlIsland(name);
		return island && island.documentElement || island && island.XMLDocument && island.XMLDocument.documentElement || null;
	}

	function loadXmlIntoIsland(name, text) {
		var island = xmlIsland(name);
		if (island && island.loadXML) {
			island.loadXML(text || "<Root/>");
		}
		return xmlRoot(name);
	}

	function childElements(node, name) {
		var list = [];
		var wanted = name && String(name).toLowerCase();
		for (var i = 0; node && i < node.childNodes.length; i += 1) {
			if (node.childNodes[i].nodeType === 1 && (!wanted || String(node.childNodes[i].nodeName).toLowerCase() === wanted)) {
				list.push(node.childNodes[i]);
			}
		}
		return list;
	}

	function getAttr(node, name) {
		return node && node.getAttribute ? node.getAttribute(name) || "" : "";
	}

	function table() {
		return document.getElementById("tblData");
	}

	function clearRows() {
		var grid = table();
		if (!grid) {
			return;
		}
		while (grid.rows.length > 1) {
			grid.deleteRow(1);
		}
	}

	function addTextCell(row, text, className, align) {
		var td = row.insertCell(-1);
		td.className = className || "ExcelDisplayCell";
		td.align = align || "Left";
		td.textContent = text == null ? "" : String(text);
		return td;
	}

	function addRadioCell(row, itemCode, classCode, serialNo) {
		var td = row.insertCell(-1);
		var radio = document.createElement("input");
		td.className = "ExcelDisplayCell";
		td.align = "center";
		radio.type = "radio";
		radio.name = "radItem";
		radio.id = "radItem" + serialNo;
		radio.value = itemCode + "|" + classCode;
		radio.setAttribute("data-item-code", itemCode);
		radio.setAttribute("data-class-code", classCode);
		td.appendChild(radio);
		return radio;
	}

	function selectedItemType() {
		return trim(valueOf("selIType")) || trim(valueOf("hItemType"));
	}

	function renderItems() {
		var root = xmlRoot("ItemData");
		var items = childElements(root, "Item");
		var grid = table();
		var emptyRow;
		var emptyCell;
		if (!grid) {
			return;
		}
		clearRows();
		if (!items.length) {
			emptyRow = grid.insertRow(-1);
			emptyCell = addTextCell(emptyRow, "Data Not Found", "ExcelDisplayCell", "Center");
			emptyCell.colSpan = 6;
			return;
		}
		for (var i = 0; i < items.length; i += 1) {
			(function (item, index) {
				var row = grid.insertRow(-1);
				var serialNo = getAttr(item, "SNO") || String(index + 1);
				var itemCode = getAttr(item, "ItemCode");
				var classCode = getAttr(item, "ClassCode");
				var uom = getAttr(item, "UOM") || "-";
				var radio;
				addTextCell(row, serialNo, "ExcelSerial", "Center");
				addTextCell(row, getAttr(item, "ItemName"), "ExcelDisplayCell", "Left");
				addTextCell(row, getAttr(item, "ComItemCode") || itemCode, "ExcelDisplayCell", "Left");
				addTextCell(row, uom, "ExcelDisplayCell", "Left");
				addTextCell(row, uom, "ExcelDisplayCell", "Left");
				radio = addRadioCell(row, itemCode, classCode, serialNo);
				row.onclick = function (event) {
					var target = event && event.target;
					if (!target || String(target.tagName || "").toLowerCase() !== "input") {
						radio.checked = true;
					}
				};
			}(items[i], i));
		}
	}

	function CheckItem() {
		var itemType = selectedItemType();
		var xhr = new XMLHttpRequest();
		var query = buildQuery({
			sIType: itemType,
			IType: itemType,
			Stock: "Y",
			hSelectMode: "S",
			Flag: "1",
			hDispButt: "Y",
			hDispItem: "Y",
			CallFrom: "INV",
			PageSize: "500",
			Page: "1"
		});
		xhr.open("GET", "../../Common/XMLGetItemSelectRel.asp?" + query, false);
		xhr.send(null);
		if (xhr.status && (xhr.status < 200 || xhr.status >= 300)) {
			alert("Unable to load item list.");
			return false;
		}
		loadXmlIntoIsland("ItemData", xhr.responseText || "<Root/>");
		renderItems();
		return false;
	}

	function selectedRadio() {
		var radios = field("radItem");
		if (!radios) {
			return null;
		}
		if (radios.length === undefined) {
			return radios.checked ? radios : null;
		}
		for (var i = 0; i < radios.length; i += 1) {
			if (radios[i].checked) {
				return radios[i];
			}
		}
		return null;
	}

	function selectedItem() {
		var radio = selectedRadio();
		var parts;
		if (!radio) {
			return null;
		}
		parts = String(radio.value || "").split("|");
		return {
			itemCode: trim(radio.getAttribute("data-item-code") || parts[0]),
			classCode: trim(radio.getAttribute("data-class-code") || parts[1])
		};
	}

	function submitTo(action) {
		var frm = form();
		frm.action = action;
		frm.submit();
	}

	function CheckSubmit(action) {
		var selected;
		var mode = trim(action || "R").toUpperCase();
		if (mode === "C") {
			submitTo("ITMCREATIONDEFINITIONENTRY.ASP");
			return false;
		}
		if (!trim(valueOf("hTempItemCode"))) {
			alert("Temporary Item is required");
			return false;
		}
		selected = selectedItem();
		if (!selected) {
			alert("Select an Item");
			return false;
		}
		setValue("hItemCode", selected.itemCode);
		setValue("hClassCode", selected.classCode);
		submitTo("TempItemRelateUpdation.asp");
		return false;
	}

	function ClearAll() {
		setValue("hItemCode", "");
		setValue("hClassCode", "");
		loadXmlIntoIsland("ItemData", "<Root/>");
		clearRows();
		return false;
	}

	window.CheckItem = CheckItem;
	window.CheckSubmit = CheckSubmit;
	window.ClearAll = ClearAll;
}(window, document));
