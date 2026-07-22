(function (window, document) {
	"use strict";

	var objTemp = null;
	var root = null;
	var rootDoc = null;
	var recType = "";
	var itemCode = "";
	var classCode = "";
	var orgId = "";
	var storeCode = "";
	var binCode = "";
	var tareValue = "";
	var rowCount = 1;
	var lotCounter = 0;

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
			cell.textContent = String(content);
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
			if (getAttr(items[i], "ICODE") === trim(itemCode) && getAttr(items[i], "CCODE") === trim(classCode) && getAttr(items[i], "OCODE") === trim(orgId)) {
				return items[i];
			}
		}
		return null;
	}

	function findStorage(itemNode) {
		var stores = elementChildren(itemNode, "STORAGE");
		for (var i = 0; i < stores.length; i += 1) {
			if (getAttr(stores[i], "STORE") === trim(storeCode) && getAttr(stores[i], "BIN") === trim(binCode)) {
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

	function storedSerialNodes() {
		var result = [];
		var containers = elementChildren(root, "STOREDLOTDETAILS");
		var i;
		var serials;
		for (i = 0; i < containers.length; i += 1) {
			serials = elementChildren(containers[i], "SERIAL");
			for (var s = 0; s < serials.length; s += 1) {
				result.push(serials[s]);
			}
		}
		return result;
	}

	function selectedLotSerialDetails(storage) {
		var result = [];
		var lots = elementChildren(storage, "LotSerial");
		var details;
		for (var i = 0; i < lots.length; i += 1) {
			details = elementChildren(lots[i], "LotSerialDetails");
			for (var d = 0; d < details.length; d += 1) {
				result.push(details[d]);
			}
		}
		return result;
	}

	function existingSelectionFlag(serialNode, storage) {
		var details = selectedLotSerialDetails(storage);
		for (var i = 0; i < details.length; i += 1) {
			if (getAttr(details[i], "LOTSERIAL") === getAttr(serialNode, "LOTSERIAL") &&
					getAttr(details[i], "QTYREC") === getAttr(serialNode, "QTYREC") &&
					getAttr(details[i], "PACKNUMBER") === getAttr(serialNode, "PACKNUMBER")) {
				return "E";
			}
		}
		return getAttr(serialNode, "FLAG");
	}

	function appendLotHeader(lotNo) {
		var row = table().insertRow(table().rows.length);
		var cell = appendCell(row, "ExcelDisplayCell", "left", lotNo || "");
		cell.colSpan = 9;
	}

	function appendLotRow(serialNode, flag) {
		var row = table().insertRow(table().rows.length);
		var checkbox;
		var lotNo = getAttr(serialNode, "LOT");
		var serialNo = getAttr(serialNode, "LOTSERIAL");
		var qtyRec = getAttr(serialNode, "QTYREC");
		var tareRec = getAttr(serialNode, "TAREREC");
		var selling = lookupDisplay("SellingData", getAttr(serialNode, "SELLINGTYPE"));
		var packing = lookupDisplay("PackingData", getAttr(serialNode, "PACKINGTYPE"));
		var formDisplay = lookupDisplay("SellingFormData", getAttr(serialNode, "SELLINGFORM"));
		var tempGross = getAttr(serialNode, "QTYGRO");

		appendCell(row, "ExcelDisplayCell", "center", rowCount);
		checkbox = makeInput("checkbox", "ChkItem" + rowCount, "", "Formelem", false);
		if (flag === "E") {
			checkbox.checked = true;
		}
		if (flag === "N") {
			checkbox.disabled = true;
		}
		appendCell(row, "ExcelFieldCell", "center", checkbox).width = "8";

		appendHidden("txtSerial" + rowCount, serialNo);
		appendHidden("txtLot" + rowCount, lotNo);
		appendCell(row, "ExcelDisplayCell", null, makeText("txtGross" + rowCount, qtyRec, 12, "right")).width = "10";
		appendCell(row, "ExcelDisplayCell", null, makeText("txtTare" + rowCount, tareRec, 12, "right")).width = "10";
		appendCell(row, "ExcelDisplayCell", "left", makeText("selSell" + rowCount, selling, 12, "left"));
		appendCell(row, "ExcelDisplayCell", null, makeText("txtSellWeight" + rowCount, getAttr(serialNode, "WEIGHTSTYPE"), 12, "right")).width = "10";
		appendHidden("txtTempGross" + rowCount, tempGross);
		appendCell(row, "ExcelDisplayCell", "left", makeText("selPack" + rowCount, packing, 12, "left"));
		appendCell(row, "ExcelDisplayCell", "left", makeText("selForm" + rowCount, formDisplay, 12, "left"));
		appendCell(row, "ExcelDisplayCell", null, makeText("txtPackNo" + rowCount, getAttr(serialNode, "PACKNUMBER"), 17, "left")).width = "10";
		rowCount += 1;
	}

	function removeChildren(node) {
		while (node && node.firstChild) {
			node.removeChild(node.firstChild);
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

	function matchingStoredSerial(serialNo, qtyRec, tareRec, sellingType, weightType, packingType, sellingForm, packNumber) {
		var nodes = storedSerialNodes();
		for (var i = 0; i < nodes.length; i += 1) {
			if (getAttr(nodes[i], "LOTSERIAL") === serialNo &&
					getAttr(nodes[i], "QTYREC") === qtyRec &&
					getAttr(nodes[i], "TAREREC") === tareRec &&
					getAttr(nodes[i], "SELLINGTYPE") === sellingType &&
					getAttr(nodes[i], "WEIGHTSTYPE") === weightType &&
					getAttr(nodes[i], "PACKINGTYPE") === packingType &&
					getAttr(nodes[i], "SELLINGFORM") === sellingForm &&
					getAttr(nodes[i], "PACKNUMBER") === packNumber) {
				return nodes[i];
			}
		}
		return null;
	}

	function findLotNode(storage, lotNo) {
		var lots = elementChildren(storage, "LotSerial");
		for (var i = 0; i < lots.length; i += 1) {
			if (getAttr(lots[i], "LOT") === trim(lotNo)) {
				return lots[i];
			}
		}
		return null;
	}

	function createLotNode(doc, lotNo) {
		var lot = doc.createElement("LotSerial");
		lotCounter += 1;
		setAttr(lot, "QTYIN", "");
		setAttr(lot, "TARE", "");
		setAttr(lot, "LOT", lotNo);
		setAttr(lot, "SERIALFROM", "");
		setAttr(lot, "SERIALTO", "");
		setAttr(lot, "TAREWEIGHT", "");
		setAttr(lot, "QTY", window.CalculateLotQty(lotNo));
		setAttr(lot, "COUNTER", lotCounter);
		return lot;
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

	window.LoadDetails = function (arg) {
		var parts = String(arg || "").split(":");
		var itemNode;
		var storage;
		var stores;
		var serials;
		var itemQty = 0;
		var enteredQty = 0;
		var previousLot = null;
		var flag;

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

		itemQty = toNumber(getAttr(itemNode, "QTY"));
		setText("idItemName", getAttr(itemNode, "ITEMNAME"));
		stores = elementChildren(itemNode, "STORAGE");
		for (var i = 0; i < stores.length; i += 1) {
			enteredQty += toNumber(getAttr(stores[i], "QTY"));
		}
		setText("idQty", itemQty - enteredQty);

		storage = findStorage(itemNode);
		serials = storedSerialNodes();
		for (var s = 0; s < serials.length; s += 1) {
			if (previousLot !== getAttr(serials[s], "LOT")) {
				previousLot = getAttr(serials[s], "LOT");
				appendLotHeader(previousLot);
			}
			flag = storage ? existingSelectionFlag(serials[s], storage) : getAttr(serials[s], "FLAG");
			appendLotRow(serials[s], flag);
		}
	};

	window.AddChecked = function () {
		var itemNode = findItem();
		var storage = itemNode ? findStorage(itemNode) : null;
		var doc = rootDoc || root && root.ownerDocument;
		var lotNo;
		var lotNode;
		var detailNode;
		var sellingType;
		var packingType;
		var sellingForm;
		var weightType;
		var sourceSerial;

		if (!storage || !doc) {
			return;
		}

		removeChildren(storage);
		setAttr(storage, "QTY", window.CalculateQty());

		for (var i = 1; i < rowCount; i += 1) {
			if (!field("ChkItem" + i) || !field("ChkItem" + i).checked) {
				continue;
			}
			lotNo = trim(field("txtLot" + i) && field("txtLot" + i).value);
			lotNode = findLotNode(storage, lotNo);
			if (!lotNode) {
				lotNode = createLotNode(doc, lotNo);
				storage.appendChild(lotNode);
				if (recType === "S") {
					break;
				}
			}
		}

		for (var row = 1; row < rowCount; row += 1) {
			if (!field("ChkItem" + row)) {
				continue;
			}

			lotNo = trim(field("txtLot" + row) && field("txtLot" + row).value);
			sellingType = window.GetSellingForm(field("selSell" + row) && field("selSell" + row).value);
			packingType = window.GetPackingCode(field("selPack" + row) && field("selPack" + row).value);
			sellingForm = window.GetPackingForm(field("selForm" + row) && field("selForm" + row).value);
			weightType = trim(field("txtSellWeight" + row) && field("txtSellWeight" + row).value);

			if (field("ChkItem" + row).checked) {
				lotNode = findLotNode(storage, lotNo);
				if (lotNode) {
					detailNode = doc.createElement("LotSerialDetails");
					setAttr(detailNode, "LOTSERIAL", trim(field("txtSerial" + row) && field("txtSerial" + row).value));
					setAttr(detailNode, "QTYREC", trim(field("txtGross" + row) && field("txtGross" + row).value));
					setAttr(detailNode, "TAREREC", trim(field("txtTare" + row) && field("txtTare" + row).value));
					setAttr(detailNode, "SELLINGTYPE", sellingType);
					setAttr(detailNode, "WEIGHTSTYPE", weightType);
					setAttr(detailNode, "PACKINGTYPE", packingType);
					setAttr(detailNode, "LOT", lotNo);
					setAttr(detailNode, "SELLINGFORM", sellingForm);
					setAttr(detailNode, "PACKNUMBER", trim(field("txtPackNo" + row) && field("txtPackNo" + row).value));
					setAttr(detailNode, "QTYGRO", trim(field("txtTempGross" + row) && field("txtTempGross" + row).value));
					lotNode.appendChild(detailNode);
				}

				sourceSerial = matchingStoredSerial(
					trim(field("txtSerial" + row) && field("txtSerial" + row).value),
					trim(field("txtGross" + row) && field("txtGross" + row).value),
					trim(field("txtTare" + row) && field("txtTare" + row).value),
					sellingType,
					weightType,
					packingType,
					sellingForm,
					trim(field("txtPackNo" + row) && field("txtPackNo" + row).value)
				);
				if (sourceSerial) {
					setAttr(sourceSerial, "FLAG", "N");
				}
			} else if (!field("ChkItem" + row).disabled) {
				sourceSerial = matchingStoredSerial(
					trim(field("txtSerial" + row) && field("txtSerial" + row).value),
					trim(field("txtGross" + row) && field("txtGross" + row).value),
					trim(field("txtTare" + row) && field("txtTare" + row).value),
					sellingType,
					weightType,
					packingType,
					sellingForm,
					trim(field("txtPackNo" + row) && field("txtPackNo" + row).value)
				);
				if (sourceSerial) {
					setAttr(sourceSerial, "FLAG", "Y");
				}
			}
		}

		closeWithReturn();
	};

	window.CalculateQty = function () {
		return selectedRowsQuantity();
	};

	window.CalculateLotQty = function (lotNo) {
		var qty = 0;
		var tare = 0;
		for (var i = 1; i < rowCount; i += 1) {
			if (field("ChkItem" + i) && field("ChkItem" + i).checked && field("txtLot" + i) && field("txtLot" + i).value === lotNo) {
				qty += toNumber(field("txtGross" + i) && field("txtGross" + i).value);
				tare += toNumber(field("txtTare" + i) && field("txtTare" + i).value);
			}
		}
		return qty - tare;
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
	window.DisableTxt = function (obj) {
		if (field("txtTare")) {
			field("txtTare").value = "";
			field("txtTare").disabled = obj && obj.value === "I";
		}
		return true;
	};
	window.DisableTxtQty = function () { return true; };
	window.DisableselSellType = function () { return true; };
	window.DisableTxtWt = function () { return true; };
	window.DisableselPack = function () { return true; };
	window.DisableselForm = function () { return true; };
	window.window_onunload = setReturnValue;
	window.addEventListener("beforeunload", setReturnValue);
}(window, document));
