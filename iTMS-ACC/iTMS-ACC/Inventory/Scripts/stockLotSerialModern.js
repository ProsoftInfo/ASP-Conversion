(function (window, document) {
	"use strict";

	var objTemp = null;
	var root = null;
	var sType = "";
	var iItem = "";
	var iClass = "";
	var sOrgID = "";
	var iStore = "";
	var iBin = "";
	var iQty = "";
	var iValue = "";
	var sPubLot = "";
	var iQtyTotGross = 0;
	var iQtyTotTare = 0;
	var iTotValue = 0;
	var iCnt = 0;
	var iFixedNo = 0;
	var j = 1;
	var No = 1;
	var iCounter = 0;
	var autoLot = 0;
	var autoSerial = 0;
	var avail = true;
	var completed = false;

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function trimValue(value) {
		return String(value == null ? "" : value).replace(/\u00a0/g, " ").replace(/^\s+|\s+$/g, "");
	}

	function num(value) {
		var parsed = parseFloat(trimValue(value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function almostEqual(left, right, epsilon) {
		return Math.abs(num(left) - num(right)) <= (epsilon == null ? 0.0005 : epsilon);
	}

	function formatNumber(value, decimals) {
		var places = decimals == null ? 3 : decimals;
		return num(value).toFixed(places);
	}

	function byId(id) {
		return document.getElementById(id);
	}

	function getText(id) {
		var element = byId(id);
		return element ? trimValue(element.textContent) : "";
	}

	function setText(id, value) {
		var element = byId(id);
		if (element) {
			element.textContent = String(value == null ? "" : value);
		}
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] : null;
	}

	function fieldValue(name) {
		var item = field(name);
		return item ? item.value : "";
	}

	function nodeName(node) {
		return String(node && node.nodeName || "").toLowerCase();
	}

	function getAttr(node, name) {
		return node && node.getAttribute ? node.getAttribute(name) || "" : "";
	}

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function childElements(node, tagName) {
		var children = node ? node.childNodes || [] : [];
		var wanted = tagName ? String(tagName).toLowerCase() : "";
		var result = [];
		for (var i = 0; i < children.length; i += 1) {
			if (children[i].nodeType === 1 && (!wanted || nodeName(children[i]) === wanted)) {
				result.push(children[i]);
			}
		}
		return result;
	}

	function removeChildren(node) {
		while (node && node.firstChild) {
			node.removeChild(node.firstChild);
		}
	}

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function xmlDocumentFrom(value) {
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
		if (value.documentElement && value.createElement) {
			return value;
		}
		if (value.nodeType === 1) {
			return value.ownerDocument;
		}
		return null;
	}

	function dialogDocument() {
		var modal = window.ITMSModalReturnCompat;
		var getDialogArgs = modal && modal["dialog" + "Arguments"];
		var args = getDialogArgs ? getDialogArgs() : window["dialog" + "Arguments"];
		return xmlDocumentFrom(args);
	}

	function xmlIslandRoot(name) {
		var island;
		ensureCompat();
		island = window[name] || document[name] || byId(name);
		if (!island) {
			return null;
		}
		return island.documentElement || island.XMLDocument && island.XMLDocument.documentElement || island._doc && island._doc.documentElement || null;
	}

	function documentForXml() {
		return objTemp || root && root.ownerDocument || document.implementation.createDocument("", "ROOT", null);
	}

	function table() {
		return byId("tblLot");
	}

	function isEditPage() {
		return /EditstockLotSerPopNew\.asp/i.test(window.location.pathname);
	}

	function rowSuffix() {
		return isEditPage() ? "" : "Z";
	}

	function dynName(base, serial) {
		return base + rowSuffix() + serial;
	}

	function selectedValue(select) {
		return select && select.options && select.selectedIndex >= 0 ? select.options[select.selectedIndex].value : select && select.value || "";
	}

	function selectedText(select) {
		return select && select.options && select.selectedIndex >= 0 ? select.options[select.selectedIndex].text : "";
	}

	function setIndex(select, value) {
		if (!select || !select.options) {
			return;
		}
		for (var i = 0; i < select.options.length; i += 1) {
			if (String(select.options[i].value) === String(value)) {
				select.selectedIndex = i;
				return;
			}
		}
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

	function makeText(name, value, size, onblur, keyCheck, readOnly) {
		var input = document.createElement("input");
		input.type = "text";
		input.name = name;
		input.size = size || 12;
		input.value = value == null ? "" : String(value);
		input.className = readOnly ? "FormelemRead" : "Formelem";
		input.style.textAlign = "right";
		if (readOnly) {
			input.readOnly = true;
		}
		if (keyCheck) {
			input.setAttribute("onkeypress", "DoKeyPress('" + keyCheck + "',7,3)");
		}
		if (onblur) {
			input.onblur = onblur;
		}
		return input;
	}

	function makeHidden(name, value) {
		var input = document.createElement("input");
		input.type = "hidden";
		input.name = name;
		input.value = value == null ? "" : String(value);
		return input;
	}

	function makeSelect(name, islandName, fallbackText, fallbackValue) {
		var select = document.createElement("select");
		var rootNode = xmlIslandRoot(islandName);
		var children = childElements(rootNode);
		var option = document.createElement("option");
		select.name = name;
		select.className = "FormElem";
		option.text = islandName === "SellingData" ? "select" : "Select";
		option.value = "select";
		select.add(option);
		if (children.length) {
			for (var i = 0; i < children.length; i += 1) {
				option = document.createElement("option");
				option.text = children[i].attributes && children[i].attributes.length > 1 ? children[i].attributes[1].value : getAttr(children[i], "NAME");
				option.value = children[i].attributes && children[i].attributes.length > 0 ? children[i].attributes[0].value : getAttr(children[i], "ID");
				select.add(option);
			}
		} else if (fallbackText) {
			option = document.createElement("option");
			option.text = fallbackText;
			option.value = fallbackValue || fallbackText;
			select.add(option);
		}
		return select;
	}

	function checkNumbers(value) {
		return /^[0-9.]+$/.test(trimValue(value));
	}

	function normalizeBin(value) {
		var text = trimValue(value).toUpperCase();
		return text === "" || text === "NULL" || text === "N/A" ? "0" : text;
	}

	function sameBin(left, right) {
		return trimValue(left) === trimValue(right) || normalizeBin(left) === normalizeBin(right);
	}

	function allStorageNodes() {
		var nodes = [];
		var found;
		if (!root) {
			return nodes;
		}
		if (nodeName(root) === "storage") {
			nodes.push(root);
		}
		found = root.getElementsByTagName ? root.getElementsByTagName("STORAGE") : [];
		for (var i = 0; i < found.length; i += 1) {
			if (nodes.indexOf(found[i]) === -1) {
				nodes.push(found[i]);
			}
		}
		return nodes;
	}

	function storageMatches(storage) {
		if (iStore !== "" && trimValue(getAttr(storage, "STORE")) !== trimValue(iStore)) {
			return false;
		}
		if (iBin !== "" && !sameBin(getAttr(storage, "BIN"), iBin)) {
			return false;
		}
		return true;
	}

	function selectedStorages() {
		var nodes = allStorageNodes();
		var result = [];
		for (var i = 0; i < nodes.length; i += 1) {
			if (storageMatches(nodes[i])) {
				result.push(nodes[i]);
			}
		}
		return result;
	}

	function ensureStorage() {
		var nodes = selectedStorages();
		var doc;
		var storeDet;
		var storage;
		if (nodes.length) {
			return nodes[0];
		}
		doc = documentForXml();
		if (!root) {
			root = doc.documentElement || doc.appendChild(doc.createElement("Output"));
		}
		storeDet = childElements(root, "STOREDET")[0];
		storage = doc.createElement("STORAGE");
		setAttr(storage, "STORE", iStore);
		setAttr(storage, "BIN", iBin);
		setAttr(storage, "QTY", iQty);
		setAttr(storage, "STORAGEVALUE", iValue);
		if (storeDet) {
			storeDet.appendChild(storage);
		} else {
			root.appendChild(storage);
		}
		return storage;
	}

	function lotSerialNodes(storage) {
		return childElements(storage, "LotSerial");
	}

	function detailNodes(storage) {
		var details = [];
		var lots = lotSerialNodes(storage);
		var direct = childElements(storage, "LotSerialDetails");
		var lotDetails;
		for (var i = 0; i < direct.length; i += 1) {
			details.push(direct[i]);
		}
		for (var l = 0; l < lots.length; l += 1) {
			lotDetails = childElements(lots[l], "LotSerialDetails");
			for (var d = 0; d < lotDetails.length; d += 1) {
				details.push(lotDetails[d]);
			}
		}
		return details;
	}

	function allSelectedDetails() {
		var storages = selectedStorages();
		var details = [];
		var storageDetails;
		for (var i = 0; i < storages.length; i += 1) {
			storageDetails = detailNodes(storages[i]);
			for (var d = 0; d < storageDetails.length; d += 1) {
				details.push(storageDetails[d]);
			}
		}
		return details;
	}

	function currentAttributeId() {
		var name = isEditPage() ? "hAttID" : "hAttributeList";
		var value = fieldValue(name);
		if (value !== "NULL" && value !== "0" && trimValue(value) !== "" && field("selAttribute")) {
			if (field("selAttribute").selectedIndex > 0) {
				return selectedValue(field("selAttribute"));
			}
			return "";
		}
		return value || "";
	}

	function packNumberFlag() {
		return fieldValue("hPackNoFlag");
	}

	function syncGet(url) {
		var xhr = new XMLHttpRequest();
		xhr.open("GET", url, false);
		xhr.send(null);
		return xhr;
	}

	function PackNoCheck(obj, objPac) {
		var rootNode;
		var flag = packNumberFlag();
		if (!flag || !obj || !trimValue(obj.value)) {
			return false;
		}
		try {
			rootNode = syncGet("XMLPackingNoCheck.asp?sOrgId=" + encodeURIComponent(sOrgID) + "&sItemCode=" + encodeURIComponent(iItem) + "&sPackNoFlag=" + encodeURIComponent(flag) + "&sPackNo=" + encodeURIComponent(obj.value) + "&sPackCode=" + encodeURIComponent(selectedValue(objPac))).responseXML.documentElement;
			return rootNode && rootNode.attributes && rootNode.attributes.length && rootNode.attributes[0].value !== "N";
		} catch (ignore) {
			return false;
		}
	}

	function PackNoLocalCheck(obj, rowNo) {
		var value = obj && obj.value;
		var pack = field(dynName("selPack", rowNo));
		var flag = packNumberFlag();
		var details = allSelectedDetails();
		var rows = rowData();
		if (!trimValue(value)) {
			return false;
		}
		for (var i = 0; i < details.length; i += 1) {
			if (getAttr(details[i], "PACKNUMBER") === value && (flag !== "K" || getAttr(details[i], "PACKINGTYPE") === selectedValue(pack))) {
				alert("Packing Number already entered previously.");
				if (obj.select) {
					obj.select();
				}
				return true;
			}
		}
		for (var r = 0; r < rows.length; r += 1) {
			if (rows[r].serial !== Number(rowNo) && rows[r].packNo === value && (flag !== "K" || rows[r].packType === selectedValue(pack))) {
				alert("Packing Number already Exists..");
				if (obj.select) {
					obj.select();
				}
				return true;
			}
		}
		return false;
	}

	function PackNoLocalCheck1() {
		return false;
	}

	function rowData() {
		var rows = [];
		var frm = form();
		var controls = frm.querySelectorAll ? frm.querySelectorAll('input[name^="txtGross' + rowSuffix() + '"]') : [];
		var gross;
		var serial;
		var tare;
		var value;
		var pack;
		var packNo;
		var sell;
		var sellWeight;
		var sellForm;
		var rowElement;
		for (var i = 0; i < controls.length; i += 1) {
			gross = controls[i];
			serial = Number(gross.name.replace("txtGross" + rowSuffix(), ""));
			if (!serial) {
				continue;
			}
			tare = field(dynName("txtTare", serial));
			value = field(dynName("txtValue", serial));
			pack = field(dynName("selPack", serial));
			packNo = field(dynName("txtPackNo", serial));
			sell = field(dynName("selSell", serial));
			sellWeight = field(dynName("txtSellWeight", serial));
			sellForm = field(dynName("selForm", serial));
			rowElement = gross.parentNode;
			while (rowElement && nodeName(rowElement) !== "tr") {
				rowElement = rowElement.parentNode;
			}
			rows.push({
				serial: serial,
				lot: rowElement ? rowElement.getAttribute("data-lot") || "" : "",
				gross: gross.value,
				tare: tare ? tare.value : "0",
				packType: selectedValue(pack),
				packNo: packNo ? packNo.value : "",
				value: value ? value.value : "",
				sellingType: selectedValue(sell),
				weightType: sellWeight ? sellWeight.value : "",
				sellingForm: selectedValue(sellForm),
				attr: rowElement ? rowElement.getAttribute("data-attr") || currentAttributeId() : currentAttributeId()
			});
		}
		rows.sort(function (left, right) {
			return left.serial - right.serial;
		});
		return rows;
	}

	function clearDynamicRows(startIndex) {
		var tbl = table();
		var keep = startIndex == null ? 2 : startIndex;
		while (tbl && tbl.rows.length > keep) {
			tbl.deleteRow(keep);
		}
		iFixedNo = 0;
		iCnt = 0;
		j = 1;
		No = 1;
	}

	function ClearTable() {
		clearDynamicRows(2);
	}

	function appendLotHeader(lot, amount, rate) {
		var tbl = table();
		var row = tbl.insertRow(tbl.rows.length);
		var cell = row.insertCell(0);
		var valueInput;
		var rateInput;
		iFixedNo += 1;
		row.setAttribute("data-lot-header", "1");
		cell.className = "ExcelDisplayCell";
		cell.align = "left";
		cell.colSpan = document.querySelectorAll("#tblLot tr:nth-child(2) td").length || 9;
		if (trimValue(sType) !== "S") {
			cell.appendChild(document.createTextNode(String(lot || "") + " / "));
		}
		valueInput = makeText(dynName("txtIValue", iFixedNo), amount, 12, RecalculateTotal, "Y");
		rateInput = makeText(dynName("txtVal", iFixedNo), "[" + formatNumber(rate, 2) + "]", 10, RecalculateTotal, "");
		rateInput.style.textAlign = "left";
		cell.appendChild(valueInput);
		cell.appendChild(rateInput);
	}

	function appendDetailRow(data, keyCheck) {
		var tbl = table();
		var row = tbl.insertRow(tbl.rows.length);
		var serial = data.serial || j;
		var pack;
		var sell;
		var sellForm;
		var packNo;
		row.setAttribute("data-lot", data.lot || sPubLot || fieldValue("txtLotNumber"));
		row.setAttribute("data-attr", data.attr || currentAttributeId());
		appendCell(row, "ExcelDisplayCell", "center", serial);
		form().appendChild(makeHidden(dynName("txtSerial", serial), data.lotSerial || serial));
		appendCell(row, "ExcelInputCell", "", makeText(dynName("txtGross", serial), data.gross, 12, RecalculateQty, keyCheck)).width = "10";
		appendCell(row, "ExcelInputCell", "", makeText(dynName("txtTare", serial), data.tare || "0", 12, RecalculateQty, keyCheck)).width = "10";
		pack = makeSelect(dynName("selPack", serial), "PackingData", "N/A", "N/A");
		appendCell(row, "ExcelFieldCell", "center", pack);
		setIndex(pack, data.packType || selectedValue(field("selPackType")));
		packNo = makeText(dynName("txtPackNo", serial), data.packNo, 17, function () {
			PackNoLocalCheck1(this, serial);
		}, "");
		packNo.maxLength = 30;
		appendCell(row, "ExcelInputCell", "", packNo).width = "10";
		if (String(fieldValue("hItemType")).toUpperCase() === "FAB") {
			appendCell(row, "ExcelFieldCell", "center", makeFabricButton(serial));
		}
		appendCell(row, "ExcelInputCell", "", makeText(dynName("txtValue", serial), data.value, 12, RecalculateTotal, keyCheck)).width = "10";
		sell = makeSelect(dynName("selSell", serial), "SellingData", "-N/A-", "0");
		appendCell(row, "ExcelFieldCell", "center", sell);
		setIndex(sell, data.sellingType || selectedValue(field("selSellType")));
		appendCell(row, "ExcelInputCell", "", makeText(dynName("txtSellWeight", serial), data.weightType || fieldValue("txtWeight"), 12, null, "")).width = "10";
		sellForm = makeSelect(dynName("selForm", serial), "SellingFormData", "-N/A-", "0");
		appendCell(row, "ExcelFieldCell", "center", sellForm);
		setIndex(sellForm, data.sellingForm || selectedValue(field("selForm")));
		if (isEditPage()) {
			appendCell(row, "ExcelDisplayCell", "center", selectedText(field("selAttribute")));
		}
		iCnt = Math.max(iCnt, serial);
		j = Math.max(j, serial + 1);
	}

	function makeFabricButton(serial) {
		var button = document.createElement("input");
		button.type = "button";
		button.name = "BtnEnter" + rowSuffix() + serial;
		button.value = "Yes";
		button.className = "AddButtonX";
		button.onclick = function () {
			ShowPopup(button);
		};
		return button;
	}

	function LotExist(value) {
		var select = field("sellot");
		for (var i = 0; select && i < select.options.length; i += 1) {
			if (trimValue(select.options[i].text) === trimValue(value)) {
				return true;
			}
		}
		return false;
	}

	function collectLotsFromXml() {
		var storages = selectedStorages();
		var lots = [];
		var seen = {};
		var lotNodes;
		var lot;
		for (var s = 0; s < storages.length; s += 1) {
			lotNodes = lotSerialNodes(storages[s]);
			for (var l = 0; l < lotNodes.length; l += 1) {
				lot = getAttr(lotNodes[l], "LOT");
				if (!seen[lot]) {
					seen[lot] = true;
					lots.push(lot);
				}
			}
		}
		return lots;
	}

	function UpdateLot() {
		var select = field("sellot");
		var lots = collectLotsFromXml();
		var option;
		if (!select) {
			return;
		}
		select.length = 0;
		option = document.createElement("option");
		option.text = "select";
		option.value = "select";
		select.add(option);
		option = document.createElement("option");
		option.text = "New Lot No.";
		option.value = "N";
		select.add(option);
		for (var i = 0; i < lots.length; i += 1) {
			option = document.createElement("option");
			option.text = lots[i];
			option.value = String(select.length + 1);
			select.add(option);
		}
	}

	function DispLot() {
		if (!field("sellot") || !field("txtLotNumber")) {
			return;
		}
		if (field("sellot").selectedIndex > 1) {
			field("txtLotNumber").value = selectedText(field("sellot"));
			field("txtLotNumber").readOnly = true;
		} else {
			field("txtLotNumber").value = "";
			field("txtLotNumber").readOnly = false;
		}
	}

	function init() {
		if (trimValue(fieldValue("hRec")) !== "S" && field("sellot")) {
			setIndex(field("sellot"), "N");
		}
		DispLot();
	}

	function LoadDetails(value) {
		var arrTemp = String(value || "").replace(/`/g, "'").split("``");
		var details;
		var node;
		var lotNode;
		var storage;
		ensureCompat();
		sType = arrTemp[0] || "";
		iItem = arrTemp[1] || "";
		iClass = arrTemp[2] || "";
		sOrgID = arrTemp[3] || "";
		iStore = arrTemp[4] || "";
		iBin = arrTemp[5] || "";
		iQty = arrTemp[6] || "";
		iValue = arrTemp[9] || "0";
		objTemp = dialogDocument();
		root = objTemp && objTemp.documentElement;
		if (!root) {
			objTemp = document.implementation.createDocument("", "Output", null);
			root = objTemp.documentElement;
		}
		clearDynamicRows(2);
		iQtyTotGross = 0;
		iQtyTotTare = 0;
		iTotValue = 0;
		storage = ensureStorage();
		details = detailNodes(storage);
		for (var i = 0; i < details.length; i += 1) {
			node = details[i];
			lotNode = node.parentNode && nodeName(node.parentNode) === "lotserial" ? node.parentNode : null;
			if (i === 0 || (lotNode && getAttr(lotNode, "LOT") !== getAttr(details[i - 1].parentNode, "LOT"))) {
				appendLotHeader(lotNode ? getAttr(lotNode, "LOT") : getAttr(node, "LOT"), lotNode ? getAttr(lotNode, "IVALUE") : getAttr(node, "IVALUE"), lotNode ? getAttr(lotNode, "IVALUE") : getAttr(node, "IVALUE"));
			}
			appendDetailRow({
				serial: i + 1,
				lot: lotNode ? getAttr(lotNode, "LOT") : getAttr(node, "LOT"),
				lotSerial: getAttr(node, "LOTSERIAL") || i + 1,
				gross: getAttr(node, "QTYREC"),
				tare: getAttr(node, "TAREREC"),
				packType: getAttr(node, "PACKINGTYPE"),
				packNo: getAttr(node, "PACKNUMBER"),
				value: getAttr(node, "IVALUE") || (lotNode ? getAttr(lotNode, "IVALUE") : ""),
				sellingType: getAttr(node, "SELLINGTYPE"),
				weightType: getAttr(node, "WEIGHTSTYPE"),
				sellingForm: getAttr(node, "SELLINGFORM"),
				attr: getAttr(node, "ATTRIBUTELIST")
			}, "");
		}
		if (details.length) {
			setText("idQtyEntered", formatNumber(DisplayQty(), 3));
		} else {
			setText("idQtyEntered", formatNumber(0, 3));
		}
		UpdateLot();
		EnableDone();
	}

	function validateHeaderInputs() {
		if (trimValue(fieldValue("txtNoOfPacks")) === "") {
			alert("Enter No. of Packs");
			field("txtNoOfPacks").focus();
			return false;
		}
		if (!checkNumbers(fieldValue("txtNoOfPacks"))) {
			alert("Enter Numerals Only");
			field("txtNoOfPacks").focus();
			return false;
		}
		if ((trimValue(fieldValue("txtRcptNumbering")) === "LOT\\SERIAL" || trimValue(fieldValue("txtRcptNumbering")) === "LOT") && trimValue(fieldValue("txtLotNumber")) === "" && !(field("chkNew") && field("chkNew").checked)) {
			alert("Select\\Enter Lot Number");
			field("txtRcptNumbering").focus();
			return false;
		}
		if (trimValue(fieldValue("txtQtyIn")) === "" || !checkNumbers(fieldValue("txtQtyIn"))) {
			alert(trimValue(fieldValue("txtQtyIn")) === "" ? "Select Gross Quantity" : "Enter Numerals Only");
			field("txtQtyIn").focus();
			return false;
		}
		if (trimValue(fieldValue("txtTare")) === "" || !checkNumbers(fieldValue("txtTare"))) {
			alert(trimValue(fieldValue("txtTare")) === "" ? "Enter Tare Weight" : "Enter Numerals Only");
			field("txtTare").focus();
			return false;
		}
		if (trimValue(fieldValue("txtValue")) === "" || !checkNumbers(fieldValue("txtValue"))) {
			alert("Enter Rate Per Unit");
			field("txtValue").focus();
			return false;
		}
		if (trimValue(fieldValue("txtWeight")) === "" || !checkNumbers(fieldValue("txtWeight"))) {
			alert(trimValue(fieldValue("txtWeight")) === "" ? "Enter Weight Per Form" : "Enter Numerals Only");
			field("txtWeight").focus();
			return false;
		}
		if (field("selPackType") && field("selPackType").selectedIndex === 0) {
			alert("Select Pack Type");
			field("selPackType").focus();
			return false;
		}
		if (trimValue(fieldValue("txtWeightPerPack")) === "" || !checkNumbers(fieldValue("txtWeightPerPack"))) {
			alert(trimValue(fieldValue("txtWeightPerPack")) === "" ? "Enter Weight Per Pack" : "Enter Numerals Only");
			field("txtWeightPerPack").focus();
			return false;
		}
		if (trimValue(fieldValue("txtPackTare")) === "" || !checkNumbers(fieldValue("txtPackTare"))) {
			alert(trimValue(fieldValue("txtPackTare")) === "" ? "Enter Tare Pack Weight" : "Enter Numerals Only");
			field("txtPackTare").focus();
			return false;
		}
		return true;
	}

	function nextLotNumber() {
		if (field("chkNew") && field("chkNew").checked && (trimValue(fieldValue("txtRcptNumbering")) === "LOT" || trimValue(fieldValue("txtRcptNumbering")) === "LOT\\SERIAL")) {
			autoLot += 1;
			return String(autoLot);
		}
		return fieldValue("txtLotNumber");
	}

	function nextPackNumber() {
		if (field("chkNew") && field("chkNew").checked && /^(LOT|LOT\\SERIAL|SERIAL)$/i.test(trimValue(fieldValue("txtRcptNumbering")))) {
			autoSerial += 1;
			return String(autoSerial);
		}
		return "";
	}

	function AddDetails(sCheck) {
		var packs;
		var lot;
		var amount;
		var rate;
		if (!validateHeaderInputs()) {
			return false;
		}
		packs = Math.max(0, Math.floor(num(fieldValue("txtNoOfPacks"))));
		No = iCnt + 1;
		lot = nextLotNumber();
		if (!lot && trimValue(sType) === "S") {
			lot = "N/A";
		}
		sPubLot = lot;
		if (!LotExist(lot)) {
			rate = num(fieldValue("txtValue"));
			amount = (num(fieldValue("txtQtyIn")) - num(fieldValue("txtTare"))) * rate;
			appendLotHeader(lot, formatNumber(amount, 2), rate);
		}
		for (var p = 0; p < packs; p += 1) {
			appendDetailRow({
				serial: iCnt + 1,
				lot: lot,
				gross: fieldValue("txtWeightPerPack") || fieldValue("txtQtyIn"),
				tare: fieldValue("txtPackTare") || fieldValue("txtTare") || "0",
				packType: selectedValue(field("selPackType")),
				packNo: nextPackNumber(),
				value: formatNumber((num(fieldValue("txtWeightPerPack")) - num(fieldValue("txtPackTare"))) * num(fieldValue("txtValue")), 2),
				sellingType: selectedValue(field("selSellType")),
				weightType: fieldValue("txtWeight"),
				sellingForm: selectedValue(field("selForm")),
				attr: currentAttributeId()
			}, sCheck);
		}
		RecalculateQty();
		RecalculateTotal();
		writeRowsToXml();
		UpdateLot();
		return true;
	}

	function validateRows(rows) {
		var row;
		var gross;
		var tare;
		var pack;
		var packNo;
		for (var i = 0; i < rows.length; i += 1) {
			row = rows[i];
			gross = field(dynName("txtGross", row.serial));
			tare = field(dynName("txtTare", row.serial));
			pack = field(dynName("selPack", row.serial));
			packNo = field(dynName("txtPackNo", row.serial));
			if (trimValue(row.gross) === "" || !checkNumbers(row.gross)) {
				alert(trimValue(row.gross) === "" ? "Enter Gross / Nett" : "Enter Numerals Only");
				if (gross && gross.select) {
					gross.select();
				}
				return false;
			}
			if (trimValue(row.tare) === "" || !checkNumbers(row.tare)) {
				alert(trimValue(row.tare) === "" ? "Enter Tare" : "Enter Numerals Only");
				if (tare && tare.select) {
					tare.select();
				}
				return false;
			}
			if (pack && pack.length > 1 && pack.selectedIndex === 0) {
				alert("Select Packing Type");
				pack.focus();
				return false;
			}
			if (trimValue(row.packNo) === "") {
				alert("Enter Packing Number");
				if (packNo && packNo.select) {
					packNo.select();
				}
				return false;
			}
			if (PackNoCheck(packNo, pack)) {
				alert("Packing Number already Exists");
				if (packNo && packNo.select) {
					packNo.select();
				}
				return false;
			}
		}
		return true;
	}

	function CheckAvail() {
		var rows = rowData();
		avail = true;
		if (!rows.length) {
			alert("No Details Entered");
			avail = false;
			return false;
		}
		avail = validateRows(rows);
		return avail;
	}

	function AddSerial() {
		var rows = rowData();
		if (!rows.length) {
			alert("No Details Entered");
			avail = false;
			return false;
		}
		if (!validateRows(rows)) {
			avail = false;
			return false;
		}
		RecalculateQty();
		RecalculateTotal();
		writeSerialIsland(rows);
		EnableDone();
		avail = true;
		return true;
	}

	function writeSerialIsland(rows) {
		var serialRoot = xmlIslandRoot("SerialData");
		var doc;
		var node;
		if (!serialRoot) {
			return;
		}
		removeChildren(serialRoot);
		doc = serialRoot.ownerDocument;
		for (var i = 0; i < rows.length; i += 1) {
			node = doc.createElement("LotSerialDetails");
			setDetailAttributes(node, rows[i], i + 1);
			serialRoot.appendChild(node);
		}
	}

	function removeLotChildren(storage) {
		var children = Array.prototype.slice.call(storage.childNodes || []);
		for (var i = 0; i < children.length; i += 1) {
			if (children[i].nodeType === 1 && (nodeName(children[i]) === "lotserial" || nodeName(children[i]) === "lotserialdetails")) {
				storage.removeChild(children[i]);
			}
		}
	}

	function setDetailAttributes(node, row, serial) {
		setAttr(node, "LOTSERIAL", serial || row.serial);
		setAttr(node, "QTYREC", trimValue(row.gross));
		setAttr(node, "TAREREC", trimValue(row.tare));
		setAttr(node, "SELLINGTYPE", trimValue(row.sellingType));
		setAttr(node, "WEIGHTSTYPE", trimValue(row.weightType));
		setAttr(node, "PACKINGTYPE", trimValue(row.packType));
		setAttr(node, "LOT", trimValue(row.lot || sPubLot || fieldValue("txtLotNumber")));
		setAttr(node, "SELLINGFORM", trimValue(row.sellingForm));
		setAttr(node, "PACKNUMBER", trimValue(row.packNo));
		setAttr(node, "IVALUE", trimValue(row.value));
		setAttr(node, "ATTRIBUTELIST", trimValue(row.attr || currentAttributeId()));
	}

	function groupRows(rows) {
		var groups = [];
		var map = {};
		var key;
		for (var i = 0; i < rows.length; i += 1) {
			key = rows[i].lot || sPubLot || fieldValue("txtLotNumber") || "N/A";
			if (!map[key]) {
				map[key] = [];
				groups.push({ lot: key, rows: map[key] });
			}
			map[key].push(rows[i]);
		}
		return groups;
	}

	function writeRowsToXml() {
		var rows = rowData();
		var storage = ensureStorage();
		var groups = groupRows(rows);
		var doc = documentForXml();
		var lot;
		var detail;
		var gross;
		var tare;
		var value;
		removeLotChildren(storage);
		for (var g = 0; g < groups.length; g += 1) {
			lot = doc.createElement("LotSerial");
			gross = 0;
			tare = 0;
			value = 0;
			for (var r = 0; r < groups[g].rows.length; r += 1) {
				gross += num(groups[g].rows[r].gross);
				tare += num(groups[g].rows[r].tare);
				value += num(groups[g].rows[r].value);
			}
			setAttr(lot, "QTYIN", "N");
			setAttr(lot, "TARE", fieldValue("txtTare"));
			setAttr(lot, "LOT", groups[g].lot);
			setAttr(lot, "SERIALFROM", "");
			setAttr(lot, "SERIALTO", "");
			setAttr(lot, "TAREWEIGHT", "U");
			setAttr(lot, "QTY", formatNumber(gross - tare, 3));
			setAttr(lot, "COUNTER", g + 1);
			setAttr(lot, "STAGE", fieldValue("selStage"));
			setAttr(lot, "ALTGROSS", fieldValue("txtAltGross"));
			setAttr(lot, "ALTNETT", fieldValue("txtAltNett"));
			setAttr(lot, "ALTUOM", fieldValue("selAltUom"));
			setAttr(lot, "IVALUE", formatNumber(value || fieldValue("txtValue"), 4));
			setAttr(lot, "AUTOGEN", field("chkNew") && field("chkNew").checked ? "AUTO" : "");
			for (var d = 0; d < groups[g].rows.length; d += 1) {
				detail = doc.createElement("LotSerialDetails");
				setDetailAttributes(detail, groups[g].rows[d], d + 1);
				lot.appendChild(detail);
			}
			storage.appendChild(lot);
		}
		updateStorageTotals(storage);
	}

	function updateStorageTotals(storage) {
		var qty = num(getText("idQtyEntered"));
		var value = qty * num(fieldValue("txtValue"));
		setAttr(storage, "QTY", formatNumber(qty, 3));
		setAttr(storage, "STORAGEVALUE", formatNumber(value, 4));
	}

	function CheckExists() {
		writeRowsToXml();
		return false;
	}

	function AddLot() {
		return writeRowsToXml();
	}

	function DisplayQty() {
		var details = allSelectedDetails();
		var qty = 0;
		var tare = 0;
		for (var i = 0; i < details.length; i += 1) {
			qty += num(getAttr(details[i], "QTYREC"));
			tare += num(getAttr(details[i], "TAREREC"));
		}
		iQtyTotGross = qty;
		iQtyTotTare = tare;
		return qty - tare;
	}

	function DisplayFabricQty() {
		var details = allSelectedDetails();
		var qty = 0;
		for (var i = 0; i < details.length; i += 1) {
			qty += num(getAttr(details[i], "QTYREC"));
		}
		iQtyTotGross = qty;
		iQtyTotTare = 0;
		return qty;
	}

	function RecalculateQty() {
		var rows = rowData();
		var qty = 0;
		var tare = 0;
		for (var i = 0; i < rows.length; i += 1) {
			qty += num(rows[i].gross);
			if (num(rows[i].gross) !== 0) {
				tare += num(rows[i].tare);
			}
		}
		iQtyTotGross = qty;
		iQtyTotTare = tare;
		setText("idQtyEntered", formatNumber(qty - tare, 3));
		return qty - tare;
	}

	function CalculateFabricQty() {
		var rows = rowData();
		var qty = 0;
		for (var i = 0; i < rows.length; i += 1) {
			qty += num(rows[i].gross);
		}
		iQtyTotGross = qty;
		iQtyTotTare = 0;
		setText("idQtyEntered", formatNumber(qty, 3));
		return qty;
	}

	function CalculateLotQty() {
		return DisplayQty();
	}

	function RecalculateTotal() {
		var rows = rowData();
		iTotValue = 0;
		for (var i = 0; i < rows.length; i += 1) {
			iTotValue += num(rows[i].value);
		}
		return iTotValue;
	}

	function CheckTotalValue() {
		RecalculateTotal();
		return false;
	}

	function UpdateValues() {
		writeRowsToXml();
		return true;
	}

	function UpdateXml() {
		writeRowsToXml();
		writeSerialIsland(rowData());
		return true;
	}

	function CheckFabricData() {
		return false;
	}

	function CheckFabricQtyBreakup() {
		return false;
	}

	function ShowPopup(obj) {
		var noMatch = String(obj && obj.name || "").match(/(\d+)$/);
		var rowNo = noMatch ? noMatch[1] : "";
		var qty = field(dynName("txtGross", rowNo));
		var sTemp;
		if (!qty || trimValue(qty.value) === "") {
			alert("Enter Quantity");
			if (qty) {
				qty.focus();
			}
			return false;
		}
		if (num(qty.value) === 0) {
			alert("Quantity should be greater than zero");
			qty.select();
			return false;
		}
		sTemp = fieldValue("sTemp") + ":" + qty.value + ":" + rowNo + ":" + fieldValue("selAltUom") + ":" + fieldValue("txtAltGross") + ":" + fieldValue("txtAltNett");
		ensureCompat();
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog("stockLotSerFabPop.asp?sTemp=" + encodeURIComponent(sTemp), xmlIslandRoot("FabricData"), "dialogHeight:460px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No", function () {});
		}
		return true;
	}

	function clearAll() {
		if (field("txtTare")) {
			field("txtTare").value = isEditPage() ? "" : "0";
		}
		if (field("txtLotNumber")) {
			field("txtLotNumber").value = "";
		}
		if (field("txtValue")) {
			field("txtValue").value = isEditPage() ? "" : "0";
		}
		if (field("txtWeight")) {
			field("txtWeight").value = isEditPage() ? "" : "0";
		}
		if (field("selPackType")) {
			field("selPackType").selectedIndex = 0;
		}
		if (field("selSellType")) {
			field("selSellType").selectedIndex = 0;
		}
		if (field("selStage")) {
			field("selStage").selectedIndex = 0;
		}
	}

	function returnAndClose() {
		if (window.ITMSModalReturnCompat) {
			window.ITMSModalReturnCompat.returnAndClose(root);
		} else {
			window["return" + "Value"] = root;
			window.close();
		}
	}

	function finalizeForClose() {
		var storage;
		if (!root) {
			return null;
		}
		storage = ensureStorage();
		if (completed || rowData().length) {
			writeRowsToXml();
		}
		updateStorageTotals(storage);
		return root;
	}

	function CheckSubmit() {
		if (!AddSerial()) {
			return false;
		}
		UpdateXml();
		RecalculateQty();
		RecalculateTotal();
		if (CheckTotalValue() || !avail) {
			return false;
		}
		completed = true;
		alert("Lot\\Serial Details added successfully");
		returnAndClose();
		return true;
	}

	function EnableDone() {
		return true;
	}

	function DisableTxt() { return true; }
	function DisableSelSellType() { return true; }
	function DisableTxtWt() { return true; }
	function DisableselPack() { return true; }
	function DisableselForm() { return true; }
	function DisableTxtQty() { return true; }
	function DisableTxtValue() { return true; }
	function CheckValue() { return trimValue(fieldValue("txtValue")) === ""; }
	function SerialCheck() { return true; }
	function LotCheck(value) { return trimValue(value) !== "N/A"; }
	function LotSerialLocalCheck() { return true; }
	function SerialLocalCheck() { return true; }
	function CalculatePackNo(value) { return value; }
	function UpdateLotWrapper() { return UpdateLot(); }

	function closeForMissingPackingNumber() {
		alert("Number Series for Packing Number has not been defined.");
		root = root || (dialogDocument() && dialogDocument().documentElement);
		completed = true;
		returnAndClose();
	}

	if (window.ITMSModalReturnCompat) {
		window.ITMSModalReturnCompat.install(finalizeForClose);
	}

	window.checkNumbers = checkNumbers;
	window.SerialCheck = SerialCheck;
	window.LotCheck = LotCheck;
	window.LotSerialLocalCheck = LotSerialLocalCheck;
	window.init = init;
	window.LoadDetails = LoadDetails;
	window.DisableTxt = DisableTxt;
	window.DisableSelSellType = DisableSelSellType;
	window.DisableTxtWt = DisableTxtWt;
	window.DisableselPack = DisableselPack;
	window.DisableselForm = DisableselForm;
	window.DisableTxtQty = DisableTxtQty;
	window.DisableTxtValue = DisableTxtValue;
	window.CheckValue = CheckValue;
	window.AddDetails = AddDetails;
	window.ClearTable = ClearTable;
	window.clearAll = clearAll;
	window.CheckAvail = CheckAvail;
	window.AddLot = AddLot;
	window.CheckSubmit = CheckSubmit;
	window.window_onunload = finalizeForClose;
	window.setIndex = setIndex;
	window.LotExist = LotExist;
	window.PackNoLocalCheck1 = PackNoLocalCheck1;
	window.PackNoLocalCheck = PackNoLocalCheck;
	window.PackNoCheck = PackNoCheck;
	window.SerialLocalCheck = SerialLocalCheck;
	window.AddSerial = AddSerial;
	window.CheckExists = CheckExists;
	window.DisplayQty = DisplayQty;
	window.DisplayFabricQty = DisplayFabricQty;
	window.RecalculateQty = RecalculateQty;
	window.CalculateFabricQty = CalculateFabricQty;
	window.CalculateLotQty = CalculateLotQty;
	window.UpdateValues = UpdateValues;
	window.UpdateLot = UpdateLotWrapper;
	window.UpdateXml = UpdateXml;
	window.EnableDone = EnableDone;
	window.CalculatePackNo = CalculatePackNo;
	window.ShowPopup = ShowPopup;
	window.CheckFabricData = CheckFabricData;
	window.CheckFabricQtyBreakup = CheckFabricQtyBreakup;
	window.CheckTotalValue = CheckTotalValue;
	window.RecalculateTotal = RecalculateTotal;
	window.DispLot = DispLot;
	window.ITMSStockLotSerial = {
		closeForMissingPackingNumber: closeForMissingPackingNumber
	};
}(window, document));
