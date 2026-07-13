(function (window, document) {
	"use strict";

	var objTemp = null;
	var root = null;
	var sType = "";
	var iClass = "";
	var sOrgID = "";
	var iStore = "";
	var iBin = "";
	var iQty = "";
	var iValue = "";
	var sAttr = "";
	var iAttr = "";
	var iQtyTotGross = 0;
	var iQtyTotTare = 0;
	var iTotValue = 0;
	var No = 1;
	var iFixedNo = 0;
	var j = 1;
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
			element.textContent = String(value);
		}
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] : null;
	}

	function getAttr(node, name) {
		return node && node.getAttribute ? node.getAttribute(name) || "" : "";
	}

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function nodeName(node) {
		return String(node && node.nodeName || "").toLowerCase();
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

	function normalizeBin(value) {
		var text = trimValue(value).toUpperCase();
		return text === "" || text === "NULL" || text === "N/A" ? "0" : text;
	}

	function sameBin(left, right) {
		return trimValue(left) === trimValue(right) || normalizeBin(left) === normalizeBin(right);
	}

	function storageMatches(storage, strict) {
		if (iStore !== "" && trimValue(getAttr(storage, "STORE")) !== trimValue(iStore)) {
			return false;
		}
		if (iBin !== "" && !sameBin(getAttr(storage, "BIN"), iBin)) {
			return false;
		}
		if (strict) {
			if (iQty !== "" && !almostEqual(getAttr(storage, "QTY"), iQty)) {
				return false;
			}
			if (iValue !== "" && !almostEqual(getAttr(storage, "STORAGEVALUE"), iValue, 0.00005)) {
				return false;
			}
		}
		return true;
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

	function matchingStorages(strict) {
		var nodes = allStorageNodes();
		var result = [];
		for (var i = 0; i < nodes.length; i += 1) {
			if (storageMatches(nodes[i], strict)) {
				result.push(nodes[i]);
			}
		}
		return result;
	}

	function selectedStorages() {
		var nodes = matchingStorages(true);
		return nodes.length ? nodes : matchingStorages(false);
	}

	function ensureStorage() {
		var nodes = selectedStorages();
		var doc;
		var storeDet;
		var storage;
		if (nodes.length) {
			return nodes[0];
		}
		if (!root) {
			doc = documentForXml();
			root = doc.documentElement || doc.appendChild(doc.createElement("Output"));
		}
		doc = documentForXml();
		storeDet = nodeName(root) === "storedet" ? root : childElements(root, "STOREDET")[0];
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
		var lots = lotSerialNodes(storage);
		var details = [];
		var direct = childElements(storage, "LotSerialDetails");
		var lotDetails;
		for (var i = 0; i < direct.length; i += 1) {
			details.push(direct[i]);
		}
		for (var jLot = 0; jLot < lots.length; jLot += 1) {
			lotDetails = childElements(lots[jLot], "LotSerialDetails");
			for (var jDetail = 0; jDetail < lotDetails.length; jDetail += 1) {
				details.push(lotDetails[jDetail]);
			}
		}
		return details;
	}

	function matchingDetailNodes(strict) {
		var storages = strict ? matchingStorages(true) : selectedStorages();
		var details = [];
		var storageDetails;
		for (var i = 0; i < storages.length; i += 1) {
			storageDetails = detailNodes(storages[i]);
			for (var jDetail = 0; jDetail < storageDetails.length; jDetail += 1) {
				details.push(storageDetails[jDetail]);
			}
		}
		if (strict && !details.length) {
			return matchingDetailNodes(false);
		}
		return details;
	}

	function copyAttributes(source, target) {
		if (!source || !source.attributes) {
			return;
		}
		for (var i = 0; i < source.attributes.length; i += 1) {
			target.setAttribute(source.attributes[i].name, source.attributes[i].value);
		}
	}

	function table() {
		return byId("tblLot");
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

	function makeInput(name, size, value, onblur, keyCheck) {
		var input = document.createElement("input");
		input.type = "text";
		input.name = name;
		input.size = size;
		input.value = value == null ? "" : String(value);
		input.className = "Formelem";
		input.style.textAlign = name.indexOf("txtBarCode") === 0 ? "left" : "right";
		if (name.indexOf("txtBarCode") === 0) {
			input.maxLength = 30;
		}
		if (keyCheck) {
			input.setAttribute("onkeypress", "DoKeyPress('" + keyCheck + "',7,3)");
		}
		if (onblur) {
			input.onblur = onblur;
		}
		return input;
	}

	function appendDetailRow(serialNo, quantity, value, barcode, attrId, attrList, attrText, keyCheck) {
		var row = table().insertRow(table().rows.length);
		var checkbox = document.createElement("input");
		checkbox.type = "checkbox";
		checkbox.name = "chkDelete" + serialNo;
		checkbox.value = serialNo;
		checkbox.className = "Formelem";
		checkbox.style.textAlign = "right";
		row.setAttribute("data-attr-id", attrId || "0");
		row.setAttribute("data-attr-list", attrList || "0");
		row.setAttribute("data-attr-text", attrText || "");
		appendCell(row, "ExcelDisplayCell", "center", serialNo);
		appendCell(row, "ExcelInputCell", "center", checkbox).width = "8";
		appendCell(row, "ExcelInputCell", "", makeInput("txtIQty" + serialNo, 12, quantity, RecalculateQty, keyCheck)).width = "10";
		appendCell(row, "ExcelInputCell", "", makeInput("txtIValue" + serialNo, 12, value, RecalculateTotal, keyCheck)).width = "10";
		appendCell(row, "ExcelInputCell", "", makeInput("txtBarCode" + serialNo, 17, barcode, null, "")).width = "10";
		if (field("hSno")) {
			field("hSno").value = serialNo;
		}
		j = Math.max(j, Number(serialNo) + 1);
	}

	function appendAttributeRow(text) {
		var row;
		var cell;
		if (!trimValue(text)) {
			return;
		}
		iFixedNo += 1;
		row = table().insertRow(table().rows.length);
		row.setAttribute("data-attribute-header", "1");
		cell = row.insertCell(0);
		cell.innerHTML = text;
		cell.className = "ExcelDisplayCell";
		cell.align = "left";
		cell.colSpan = 5;
	}

	function ClearTable() {
		var tbl = table();
		while (tbl && tbl.rows.length > 2) {
			tbl.deleteRow(2);
		}
		iFixedNo = 0;
		j = 1;
	}

	function checkNumbers(value) {
		return /^[0-9.]+$/.test(trimValue(value));
	}

	function rowData() {
		var inputs = form().querySelectorAll ? form().querySelectorAll('input[name^="txtIQty"]') : [];
		var rows = [];
		var qtyInput;
		var serial;
		var row;
		var valueInput;
		var barcodeInput;
		for (var i = 0; i < inputs.length; i += 1) {
			qtyInput = inputs[i];
			serial = Number(String(qtyInput.name).replace("txtIQty", ""));
			if (!serial) {
				continue;
			}
			row = qtyInput.parentNode;
			while (row && nodeName(row) !== "tr") {
				row = row.parentNode;
			}
			valueInput = field("txtIValue" + serial);
			barcodeInput = field("txtBarCode" + serial);
			rows.push({
				serial: serial,
				qty: qtyInput.value,
				value: valueInput ? valueInput.value : "",
				barcode: barcodeInput ? barcodeInput.value : "",
				attrId: row ? row.getAttribute("data-attr-id") || "0" : "0",
				attrList: row ? row.getAttribute("data-attr-list") || "0" : "0",
				attrText: row ? row.getAttribute("data-attr-text") || "" : ""
			});
		}
		rows.sort(function (left, right) {
			return left.serial - right.serial;
		});
		return rows;
	}

	function collectAttributes(requireSelection) {
		var elements = form().elements;
		var text = "";
		var values = "";
		var control;
		var option;
		for (var i = 0; i < elements.length; i += 1) {
			control = elements[i];
			if (!control || control.type !== "select-one" || String(control.name || "").indexOf("selAttr") !== 0) {
				continue;
			}
			if (requireSelection && control.selectedIndex === 0) {
				alert("Select Attribute");
				control.focus();
				return null;
			}
			if (control.selectedIndex >= 0) {
				option = control.options[control.selectedIndex];
				text += " - " + trimValue(option.text);
				values += "," + trimValue(option.value);
			}
		}
		return {
			text: text ? text.substring(3) : "",
			values: values ? values.substring(1) : "0"
		};
	}

	function rebuildSerialIsland(rows) {
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
			setAttr(node, "LOTSERIAL", i + 1);
			setAttr(node, "SERIALQTY", trimValue(rows[i].qty));
			setAttr(node, "IVALUE", trimValue(rows[i].value));
			setAttr(node, "BARCODE", trimValue(rows[i].barcode));
			serialRoot.appendChild(node);
		}
	}

	function removeLotSerialChildren(storage) {
		var children = Array.prototype.slice.call(storage.childNodes || []);
		for (var i = 0; i < children.length; i += 1) {
			if (children[i].nodeType === 1 && (nodeName(children[i]) === "lotserial" || nodeName(children[i]) === "lotserialdetails")) {
				storage.removeChild(children[i]);
			}
		}
	}

	function groupRows(rows) {
		var groups = [];
		var byKey = {};
		var key;
		for (var i = 0; i < rows.length; i += 1) {
			key = rows[i].attrList || rows[i].attrId || "0";
			if (!byKey[key]) {
				byKey[key] = {
					attrId: rows[i].attrId || "0",
					attrList: rows[i].attrList || "0",
					attrText: rows[i].attrText || "",
					rows: []
				};
				groups.push(byKey[key]);
			}
			byKey[key].rows.push(rows[i]);
		}
		return groups;
	}

	function writeRowsToXml(rows) {
		var storages = selectedStorages();
		var storage;
		var existingLots;
		var template;
		var groups;
		var doc;
		var lot;
		var detail;
		var groupQty;
		var groupValue;
		var row;
		if (!storages.length) {
			storages = [ensureStorage()];
		}
		doc = documentForXml();
		groups = groupRows(rows);
		for (var s = 0; s < storages.length; s += 1) {
			storage = storages[s];
			existingLots = lotSerialNodes(storage);
			template = existingLots[0] || null;
			removeLotSerialChildren(storage);
			for (var g = 0; g < groups.length; g += 1) {
				lot = doc.createElement("LotSerial");
				copyAttributes(template, lot);
				groupQty = 0;
				groupValue = 0;
				for (var r = 0; r < groups[g].rows.length; r += 1) {
					groupQty += num(groups[g].rows[r].qty);
					groupValue += num(groups[g].rows[r].value);
				}
				setAttr(lot, "LOT", getAttr(lot, "LOT") || "NULL");
				setAttr(lot, "ATTRIBUTE", groups[g].attrList || "0");
				setAttr(lot, "ATTLIST", groups[g].attrList || "0");
				setAttr(lot, "SERIALFROM", getAttr(lot, "SERIALFROM") || "0");
				setAttr(lot, "SERIALTO", getAttr(lot, "SERIALTO") || "0");
				setAttr(lot, "QTY", formatNumber(groupQty, 3));
				setAttr(lot, "COUNTER", g + 1);
				setAttr(lot, "IVALUE", formatNumber(groupValue, 4));
				for (var d = 0; d < groups[g].rows.length; d += 1) {
					row = groups[g].rows[d];
					detail = doc.createElement("LotSerialDetails");
					setAttr(detail, "LOTSERIAL", d + 1);
					setAttr(detail, "SERIALQTY", trimValue(row.qty));
					setAttr(detail, "LOT", "0");
					setAttr(detail, "PACKNUMBER", "0");
					setAttr(detail, "IVALUE", trimValue(row.value));
					setAttr(detail, "ATTRIBUTEID", row.attrId || "0");
					setAttr(detail, "ATTRIBUTELIST", row.attrList || "0");
					setAttr(detail, "BARCODE", trimValue(row.barcode));
					lot.appendChild(detail);
				}
				storage.appendChild(lot);
			}
		}
	}

	function LoadDetails(value) {
		var arrTemp = String(value || "").split("``");
		var details;
		var node;
		ensureCompat();
		sType = arrTemp[0] || "";
		iClass = arrTemp[1] || "";
		sOrgID = arrTemp[2] || "";
		iStore = arrTemp[3] || "";
		iBin = arrTemp[4] || "";
		iQty = arrTemp[5] || "";
		iValue = arrTemp[8] || "";
		sAttr = arrTemp[11] || "";
		objTemp = dialogDocument();
		root = objTemp && objTemp.documentElement;
		iQtyTotGross = 0;
		iQtyTotTare = 0;
		iTotValue = 0;
		ClearTable();
		if (!root) {
			setText("idQtyEntered", formatNumber(0, 3));
			return;
		}
		details = matchingDetailNodes(true);
		for (var i = 0; i < details.length; i += 1) {
			node = details[i];
			setAttr(node, "LOTSERIAL", i + 1);
			appendDetailRow(
				i + 1,
				getAttr(node, "SERIALQTY") || getAttr(node, "QTYREC"),
				getAttr(node, "IVALUE"),
				getAttr(node, "BARCODE"),
				getAttr(node, "ATTRIBUTEID"),
				getAttr(node, "ATTRIBUTELIST") || getAttr(node.parentNode, "ATTLIST") || getAttr(node.parentNode, "ATTRIBUTE"),
				"",
				""
			);
		}
		setText("idQtyEntered", formatNumber(DisplayQty(), 3));
		RecalculateTotal();
		EnableDone();
	}

	function DeleteItems() {
		var rows = rowData();
		var selected = [];
		var checkbox;
		for (var i = 0; i < rows.length; i += 1) {
			checkbox = field("chkDelete" + rows[i].serial);
			if (checkbox && checkbox.checked) {
				selected.push(rows[i].serial);
			}
		}
		if (selected.length > 1) {
			alert("Delete Entries One by One");
			return false;
		}
		if (!selected.length) {
			return false;
		}
		rows = rows.filter(function (row) {
			return row.serial !== selected[0];
		});
		for (var r = 0; r < rows.length; r += 1) {
			rows[r].serial = r + 1;
		}
		writeRowsToXml(rows);
		ClearTable();
		for (var a = 0; a < rows.length; a += 1) {
			appendDetailRow(a + 1, rows[a].qty, rows[a].value, rows[a].barcode, rows[a].attrId, rows[a].attrList, rows[a].attrText, "");
		}
		RecalculateQty();
		RecalculateTotal();
		EnableDone();
		return true;
	}

	function CheckValue() {
		if (trimValue(field("txtValue").value) === "") {
			alert("Enter Value");
			field("txtValue").focus();
			return true;
		}
		if (trimValue(field("txtQuantity").value) === "") {
			alert("Enter Quantity");
			field("txtQuantity").focus();
			return true;
		}
		if (!checkNumbers(field("txtQuantity").value)) {
			alert("Enter Numerals Only");
			field("txtQuantity").select();
			return true;
		}
		return false;
	}

	function CheckQtySum() {
		return num(getText("idQtyEntered")) + num(field("txtQuantity").value) > num(getText("idQty"));
	}

	function CheckTotalValue() {
		var rows = rowData();
		var total = num(field("txtValue").value);
		for (var i = 0; i < rows.length; i += 1) {
			total += num(rows[i].value);
		}
		iTotValue = total;
		if (total > num(iValue)) {
			alert("Total value should be equal to (" + iValue + "). You have Entered (" + total + ").");
			field("txtValue").select();
			return true;
		}
		return false;
	}

	function AddDetails(sCheck) {
		var attrs;
		var tbl = table();
		var serial;
		if (almostEqual(getText("idQtyEntered"), getText("idQty"))) {
			alert("Total Received Quantity has been entered");
			return false;
		}
		if (CheckValue()) {
			return false;
		}
		if (CheckQtySum()) {
			alert("Total Received Quantity should be Less than or equal to Quantity (" + getText("idQty") + ")");
			return false;
		}
		if (CheckTotalValue()) {
			return false;
		}
		attrs = collectAttributes(true);
		if (!attrs) {
			return false;
		}
		iAttr = attrs.values || "0";
		sAttr = attrs.text || "";
		No = Math.max(1, tbl.rows.length - 1 - iFixedNo);
		serial = No;
		appendAttributeRow(attrs.text);
		appendDetailRow(serial, trimValue(field("txtQuantity").value), trimValue(field("txtValue").value), "", attrs.values, attrs.values, attrs.text, sCheck);
		if (field("txtIQty" + serial)) {
			field("txtIQty" + serial).focus();
		}
		AddSerial();
		return true;
	}

	function clearAll() {
		field("txtQuantity").value = "";
		field("txtValue").value = "";
	}

	function validateRows(rows) {
		var row;
		var qtyInput;
		var valueInput;
		for (var i = 0; i < rows.length; i += 1) {
			row = rows[i];
			qtyInput = field("txtIQty" + row.serial);
			valueInput = field("txtIValue" + row.serial);
			if (trimValue(row.qty) === "") {
				alert("Enter Quantity");
				if (qtyInput) {
					qtyInput.select();
				}
				return false;
			}
			if (!checkNumbers(row.qty)) {
				alert("Enter Numerals Only");
				if (qtyInput) {
					qtyInput.select();
				}
				return false;
			}
			if (trimValue(row.value) === "") {
				alert("Enter Value");
				if (valueInput) {
					valueInput.focus();
				}
				return false;
			}
		}
		return true;
	}

	function AddSerial() {
		var rows = rowData();
		if (!rows.length) {
			alert("No Details Entered");
			return false;
		}
		if (!validateRows(rows)) {
			return false;
		}
		RecalculateQty();
		if (num(getText("idQtyEntered")) > num(getText("idQty"))) {
			alert("Total Received Quantity should be Less than or equal to Quantity (" + getText("idQty") + ")");
			return false;
		}
		RecalculateTotal();
		if (num(iTotValue) > num(iValue)) {
			alert("Total Value should be equal to (" + iValue + ")");
			return false;
		}
		rebuildSerialIsland(rows);
		EnableDone();
		return true;
	}

	function CheckExists() {
		writeRowsToXml(rowData());
		return false;
	}

	function DisplayQty() {
		var details = matchingDetailNodes(false);
		var total = 0;
		for (var i = 0; i < details.length; i += 1) {
			total += num(getAttr(details[i], "SERIALQTY") || getAttr(details[i], "QTYREC"));
		}
		iQtyTotGross = total;
		iQtyTotTare = 0;
		return total;
	}

	function RecalculateQty() {
		var rows = rowData();
		var total = 0;
		for (var i = 0; i < rows.length; i += 1) {
			total += num(rows[i].qty);
		}
		iQtyTotGross = total;
		iQtyTotTare = 0;
		setText("idQtyEntered", formatNumber(total, 3));
		EnableDone();
		return total;
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

	function UpdateXml() {
		var rows = rowData();
		writeRowsToXml(rows);
		rebuildSerialIsland(rows);
		return true;
	}

	function clearMatchingLotSerial() {
		var storages = selectedStorages();
		for (var i = 0; i < storages.length; i += 1) {
			removeLotSerialChildren(storages[i]);
		}
	}

	function finalizeForClose() {
		var calculated;
		var expected;
		if (!root) {
			return null;
		}
		if (!completed) {
			if (String(field("hItemType") && field("hItemType").value || "").toUpperCase() !== "FAB") {
				calculated = formatNumber(CalculateLotQty(), 3);
				expected = formatNumber(getText("idQty"), 3);
			} else {
				calculated = formatNumber(getText("idQtyEntered"), 3);
				expected = formatNumber(getText("idQty"), 3);
			}
			if (calculated !== expected) {
				clearMatchingLotSerial();
			}
		}
		return root;
	}

	function returnAndClose() {
		if (window.ITMSModalReturnCompat) {
			window.ITMSModalReturnCompat.returnAndClose(root);
		} else {
			window["return" + "Value"] = root;
			window.close();
		}
	}

	function CheckSubmit() {
		var rows = rowData();
		if (xmlIslandRoot("SerialData") && xmlIslandRoot("SerialData").hasChildNodes()) {
			CheckExists();
		}
		if (!validateRows(rows)) {
			return false;
		}
		UpdateXml();
		RecalculateQty();
		RecalculateTotal();
		if (!almostEqual(getText("idQtyEntered"), getText("idQty"))) {
			alert("Total Received Quantity should be equal to Quantity (" + getText("idQty") + ")");
			return false;
		}
		if (!almostEqual(iTotValue, iValue, 0.00005)) {
			alert("Total Received Value should be equal to (" + formatNumber(iValue, 4) + ")");
			return false;
		}
		iQtyTotGross = 0;
		iQtyTotTare = 0;
		completed = true;
		returnAndClose();
		return true;
	}

	function setIndex(objSch, value) {
		for (var i = 0; objSch && i < objSch.length; i += 1) {
			if (String(value) === String(objSch.options[i].value)) {
				objSch.selectedIndex = i;
				return;
			}
		}
	}

	function EnableDone() {
		return true;
	}

	if (window.ITMSModalReturnCompat) {
		window.ITMSModalReturnCompat.install(finalizeForClose);
	}

	window.checkNumbers = checkNumbers;
	window.LoadDetails = LoadDetails;
	window.DeleteItems = DeleteItems;
	window.ClearTable = ClearTable;
	window.AddDetails = AddDetails;
	window.clearAll = clearAll;
	window.CheckValue = CheckValue;
	window.CheckSubmit = CheckSubmit;
	window.window_onunload = finalizeForClose;
	window.setIndex = setIndex;
	window.AddSerial = AddSerial;
	window.CheckExists = CheckExists;
	window.CheckQtySum = CheckQtySum;
	window.DisplayQty = DisplayQty;
	window.RecalculateQty = RecalculateQty;
	window.CalculateLotQty = CalculateLotQty;
	window.UpdateXml = UpdateXml;
	window.EnableDone = EnableDone;
	window.CheckTotalValue = CheckTotalValue;
	window.RecalculateTotal = RecalculateTotal;
}(window, document));
