(function (window, document) {
	"use strict";

	var objTemp = null;
	var root = null;
	var rootDoc = null;
	var itemCode = "";
	var classCode = "";
	var orgId = "";
	var recType = "";
	var storeCode = "";
	var binCode = "";
	var tareValue = "";
	var rowCount = 1;
	var currentStoreEntryNo = "";

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] : null;
	}

	function byId(id) {
		return document.getElementById(id) || window[id] || null;
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function formatNumber(value, decimals) {
		var places = decimals == null ? 3 : decimals;
		return Number(toNumber(value).toFixed(places));
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
		return value.documentElement || value.XMLDocument && value.XMLDocument.documentElement || value._doc && value._doc.documentElement || value;
	}

	function dataRoot(id) {
		ensureCompat();
		return xmlRoot(byId(id));
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

	function attrAt(node, index) {
		var attr = node && node.attributes && node.attributes.item(index);
		return trim(attr ? attr.nodeValue || attr.value || "" : "");
	}

	function setText(id, value) {
		var element = byId(id);
		if (element) {
			element.textContent = String(value == null ? "" : value);
			element.innerText = String(value == null ? "" : value);
		}
	}

	function makeInput(type, name, value, className, readOnly) {
		var input = document.createElement("input");
		input.type = type || "text";
		input.name = name;
		input.value = value == null ? "" : String(value);
		input.className = className || "Formelem";
		if (readOnly) {
			input.readOnly = true;
		}
		return input;
	}

	function makeText(name, value, size, align) {
		var input = makeInput("text", name, value, "FormelemRead", true);
		input.size = size || 12;
		input.maxLength = 30;
		input.style.textAlign = align || "right";
		return input;
	}

	function appendHidden(name, value) {
		var input = makeInput("hidden", name, value, "", false);
		form().appendChild(input);
		return input;
	}

	function appendCell(row, className, align, content) {
		var cell = row.insertCell(row.cells.length);
		cell.className = className;
		if (align) {
			cell.align = align;
		}
		if (typeof content === "string" || typeof content === "number") {
			cell.innerHTML = String(content);
		} else if (content) {
			cell.appendChild(content);
		}
		return cell;
	}

	function table() {
		return byId("tblLot");
	}

	function findItem() {
		var items = elementChildren(root, "ITEM");
		for (var i = 0; i < items.length; i += 1) {
			if (getAttr(items[i], "ITEM") === trim(itemCode) && getAttr(items[i], "CLASS") === trim(classCode)) {
				return items[i];
			}
		}
		return null;
	}

	function findStorage(itemNode, store, bin) {
		var stores = elementChildren(itemNode, "STORAGE");
		for (var i = 0; i < stores.length; i += 1) {
			if (getAttr(stores[i], "STORE") === trim(store) && getAttr(stores[i], "BIN") === trim(bin)) {
				return stores[i];
			}
		}
		return null;
	}

	function lookupDisplay(islandId, code) {
		var lookupRoot = dataRoot(islandId);
		var nodes = elementChildren(lookupRoot);
		for (var i = 0; i < nodes.length; i += 1) {
			if (attrAt(nodes[i], 0) === trim(code)) {
				return attrAt(nodes[i], 1);
			}
		}
		return "";
	}

	function lookupCode(islandId, display) {
		var lookupRoot = dataRoot(islandId);
		var nodes = elementChildren(lookupRoot);
		for (var i = 0; i < nodes.length; i += 1) {
			if (attrAt(nodes[i], 1) === trim(display)) {
				return attrAt(nodes[i], 0);
			}
		}
		return "0";
	}

	function matchingLotRows(storage) {
		return elementChildren(storage, "LOT");
	}

	function selectedRowsQuantity() {
		var qty = 0;
		for (var i = 1; i < rowCount; i += 1) {
			if (field("ChkItem" + i) && field("ChkItem" + i).checked) {
				qty += toNumber(field("txtGross" + i) && field("txtGross" + i).value);
			}
		}
		return qty;
	}

	function appendLotHeader(lotNo) {
		var row = table().insertRow(table().rows.length);
		var cell = appendCell(row, "ExcelDisplayCell", "left", lotNo || "&nbsp;");
		cell.colSpan = 10;
	}

	function appendLotRow(lot) {
		var row = table().insertRow(table().rows.length);
		var checkbox;
		var flag = getAttr(lot, "FLAG");
		var lotNo = getAttr(lot, "LOT");
		var serialNo = getAttr(lot, "SERIALNO");
		var grossQty = getAttr(lot, "GROSSQTY");
		var qty = getAttr(lot, "QTY");
		var tare = toNumber(grossQty) - toNumber(qty);
		var selling = lookupDisplay("SellingData", getAttr(lot, "SELLINGNUMBER"));
		var packing = lookupDisplay("PackingData", getAttr(lot, "PACKINGCODE"));
		var formDisplay = lookupDisplay("SellingFormData", getAttr(lot, "SELLINGFORM"));

		appendCell(row, "ExcelDisplayCell", "center", rowCount);
		checkbox = makeInput("checkbox", "ChkItem" + rowCount, "", "Formelem", false);
		if (flag !== "N") {
			checkbox.checked = true;
		}
		if (flag === "N") {
			checkbox.disabled = true;
		}
		appendCell(row, "ExcelFieldCell", "center", checkbox).width = "8";

		appendHidden("txtSerial" + rowCount, serialNo);
		appendHidden("txtLot" + rowCount, lotNo);
		appendCell(row, "ExcelDisplayCell", null, makeText("txtG" + rowCount, grossQty, 12, "right")).width = "10";
		appendCell(row, "ExcelDisplayCell", null, makeText("txtGross" + rowCount, qty, 12, "right")).width = "10";
		appendCell(row, "ExcelDisplayCell", null, makeText("txtTare" + rowCount, tare, 12, "right")).width = "10";
		appendCell(row, "ExcelDisplayCell", "left", makeText("selSell" + rowCount, selling, 12, "left"));
		appendCell(row, "ExcelDisplayCell", null, makeText("txtSellWeight" + rowCount, getAttr(lot, "WEIGHTPERSELLINGFORM"), 12, "right")).width = "10";
		appendHidden("txtTempGross" + rowCount, grossQty);
		appendCell(row, "ExcelDisplayCell", "left", makeText("selPack" + rowCount, packing, 12, "left"));
		appendCell(row, "ExcelDisplayCell", "left", makeText("selForm" + rowCount, formDisplay, 12, "left"));
		appendHidden("hAttList" + rowCount, getAttr(lot, "ATTRIBUTE"));
		appendHidden("hSQ" + rowCount, getAttr(lot, "SQ"));
		appendHidden("hRate" + rowCount, getAttr(lot, "RATE"));
		appendCell(row, "ExcelDisplayCell", null, makeText("txtPackNo" + rowCount, getAttr(lot, "PACKINGNUMBER"), 17, "left")).width = "10";
		rowCount += 1;
	}

	function clearItemStorages(itemNode) {
		var stores = elementChildren(itemNode, "STORAGE");
		var lots;
		for (var i = 0; i < stores.length; i += 1) {
			lots = elementChildren(stores[i], "LOT");
			for (var l = 0; l < lots.length; l += 1) {
				stores[i].removeChild(lots[l]);
			}
			setAttr(stores[i], "STOREQTY", "0");
		}
	}

	function pruneEmptyStorages(itemNode) {
		var stores = elementChildren(itemNode, "STORAGE");
		for (var i = 0; i < stores.length; i += 1) {
			if (toNumber(getAttr(stores[i], "STOREQTY")) <= 0) {
				itemNode.removeChild(stores[i]);
			}
		}
	}

	function setReturnValue() {
		if (!root) {
			root = xmlRoot(objTemp);
		}
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

	function closeWithReturn() {
		setReturnValue();
		window.close();
	}

	window.LoadDetails = function (arg) {
		var parts = String(arg || "").split(":");
		var itemNode;
		var storage;
		var stores;
		var lots;
		var itemQty = 0;
		var enteredQty = 0;
		var previousLot = null;

		recType = parts[0] || "";
		itemCode = parts[1] || "";
		classCode = parts[2] || "";
		orgId = parts[3] || "";
		storeCode = parts[4] || "";
		binCode = parts[5] || "";
		tareValue = parts[7] || "";

		objTemp = modalArgs();
		root = xmlRoot(objTemp);
		rootDoc = xmlDocument(objTemp);
		itemNode = findItem();
		if (!itemNode) {
			return;
		}

		itemQty = toNumber(getAttr(itemNode, "ITEMQTY"));
		setText("idItemName", getAttr(itemNode, "ITEMNAME"));
		stores = elementChildren(itemNode, "STORAGE");
		for (var s = 0; s < stores.length; s += 1) {
			enteredQty += toNumber(getAttr(stores[s], "STOREQTY"));
		}
		setText("idQty", formatNumber(itemQty - enteredQty, 3));

		storage = findStorage(itemNode, storeCode, binCode);
		currentStoreEntryNo = getAttr(storage, "STOENTRYNO");
		lots = matchingLotRows(storage);
		for (var i = 0; i < lots.length; i += 1) {
			if (previousLot !== getAttr(lots[i], "LOT")) {
				previousLot = getAttr(lots[i], "LOT");
				appendLotHeader(previousLot);
			}
			appendLotRow(lots[i]);
		}
	};

	window.AddChecked = function () {
		var itemNode = findItem();
		var targetStorage;
		var lot;
		var doc = rootDoc || root && root.ownerDocument;
		if (!itemNode || !doc) {
			return;
		}
		clearItemStorages(itemNode);
		targetStorage = findStorage(itemNode, storeCode, binCode);
		if (!targetStorage) {
			return;
		}
		for (var i = 1; i < rowCount; i += 1) {
			if (!field("ChkItem" + i) || !field("ChkItem" + i).checked) {
				continue;
			}
			lot = doc.createElement("LOT");
			setAttr(lot, "LOTENTRYNO", currentStoreEntryNo);
			setAttr(lot, "ITEM", itemCode);
			setAttr(lot, "CLASS", classCode);
			setAttr(lot, "STORE", storeCode);
			setAttr(lot, "BIN", binCode);
			setAttr(lot, "LOT", trim(field("txtLot" + i) && field("txtLot" + i).value));
			setAttr(lot, "QTY", trim(field("txtGross" + i) && field("txtGross" + i).value));
			setAttr(lot, "RATE", field("hRate" + i) && field("hRate" + i).value);
			setAttr(lot, "GROSSQTY", trim(field("txtTempGross" + i) && field("txtTempGross" + i).value));
			setAttr(lot, "PACKINGNUMBER", trim(field("txtPackNo" + i) && field("txtPackNo" + i).value));
			setAttr(lot, "PACKINGCODE", window.GetPackingCode(field("selPack" + i) && field("selPack" + i).value));
			setAttr(lot, "SELLINGNUMBER", window.GetSellingForm(field("selSell" + i) && field("selSell" + i).value));
			setAttr(lot, "WEIGHTPERSELLINGFORM", trim(field("txtSellWeight" + i) && field("txtSellWeight" + i).value));
			setAttr(lot, "SELLINGFORM", window.GetPackingForm(field("selForm" + i) && field("selForm" + i).value));
			setAttr(lot, "STAGE", "0");
			setAttr(lot, "ATTRIBUTE", field("hAttList" + i) && field("hAttList" + i).value);
			setAttr(lot, "SERIALNO", trim(field("txtSerial" + i) && field("txtSerial" + i).value));
			setAttr(lot, "SQ", field("hSQ" + i) && field("hSQ" + i).value);
			setAttr(lot, "FLAG", "Y");
			targetStorage.appendChild(lot);
		}
		setAttr(targetStorage, "STOREQTY", window.CalculateQty());
		pruneEmptyStorages(itemNode);
		closeWithReturn();
	};

	window.CalculateQty = function () {
		return selectedRowsQuantity();
	};

	window.CalculateLotQty = function (lotNo) {
		var qty = 0;
		for (var i = 1; i < rowCount; i += 1) {
			if (field("ChkItem" + i) && field("ChkItem" + i).checked && field("txtLot" + i) && field("txtLot" + i).value === lotNo) {
				qty += toNumber(field("txtGross" + i) && field("txtGross" + i).value) - toNumber(field("txtTare" + i) && field("txtTare" + i).value);
			}
		}
		return qty;
	};

	window.SelectAll = function () {
		var checked = field("ChkAll") && field("ChkAll").checked;
		for (var i = 1; i < rowCount; i += 1) {
			if (field("ChkItem" + i) && !field("ChkItem" + i).disabled) {
				field("ChkItem" + i).checked = checked;
			}
		}
	};

	window.GetSellingForm = function (sellingForm) {
		return lookupCode("SellingData", sellingForm);
	};

	window.GetPackingCode = function (packingCode) {
		return lookupCode("PackingData", packingCode);
	};

	window.GetPackingForm = function (packingForm) {
		return lookupCode("SellingFormData", packingForm);
	};

	window.CloseForMissingPackingNumber = function () {
		alert("Number Series for Packing Number has not been defined.");
		objTemp = modalArgs();
		root = xmlRoot(objTemp);
		rootDoc = xmlDocument(objTemp);
		closeWithReturn();
	};

	window.CloseForMissingPackingDefinition = function () {
		alert("Packing Type or Selling Form has not been defined.");
		objTemp = modalArgs();
		root = xmlRoot(objTemp);
		rootDoc = xmlDocument(objTemp);
		closeWithReturn();
	};

	window.AddDetails = function () {
		alert("Manual lot/serial add script is not available in the uploaded source.");
		return false;
	};
	window.AddSerial = window.AddDetails;
	window.AddLot = window.AddDetails;
	window.DisableTxt = function () { return true; };
	window.DisableTxtQty = function () { return true; };
	window.DisableselSellType = function () { return true; };
	window.DisableTxtWt = function () { return true; };
	window.DisableselPack = function () { return true; };
	window.DisableselForm = function () { return true; };
	window.window_onunload = setReturnValue;
	window.addEventListener("beforeunload", setReturnValue);
}(window, document));
