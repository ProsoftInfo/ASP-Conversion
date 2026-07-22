(function (window, document) {
	"use strict";

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
		var number = toNumber(value);
		return Number(number.toFixed(decimals == null ? 3 : decimals));
	}

	function checkNumbers(value) {
		return /^[0-9.]+$/.test(String(value || ""));
	}

	function pad2(value) {
		return value < 10 ? "0" + value : String(value);
	}

	function parseDate(value) {
		var text = trim(value);
		var match;
		var year;
		var date;
		if (!text) {
			return null;
		}
		match = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
		if (match) {
			year = Number(match[3]);
			if (year < 100) {
				year += 2000;
			}
			date = new Date(year, Number(match[2]) - 1, Number(match[1]));
			return date.getFullYear() === year && date.getMonth() === Number(match[2]) - 1 && date.getDate() === Number(match[1]) ? date : null;
		}
		match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
		if (match) {
			date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
			return date.getFullYear() === Number(match[1]) && date.getMonth() === Number(match[2]) - 1 && date.getDate() === Number(match[3]) ? date : null;
		}
		return null;
	}

	function toIsoDate(value) {
		var date = parseDate(value);
		return date ? date.getFullYear() + "-" + pad2(date.getMonth() + 1) + "-" + pad2(date.getDate()) : "";
	}

	function dateDiffDays(start, end) {
		var startDate = parseDate(start);
		var endDate = parseDate(end);
		if (!startDate || !endDate) {
			return 0;
		}
		return Math.floor((endDate.getTime() - startDate.getTime()) / 86400000);
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

	function dataIsland(id) {
		ensureCompat();
		return byId(id || "Data");
	}

	function dataDocument(id) {
		return xmlDocument(dataIsland(id || "Data"));
	}

	function dataRoot(id) {
		return xmlRoot(dataIsland(id || "Data"));
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

	function serializeXml(value) {
		var doc = value && (value.nodeType === 9 ? value : value.ownerDocument);
		var target = doc || value;
		return target ? new XMLSerializer().serializeToString(target) : "";
	}

	function responseXmlText(xhr) {
		if (xhr.responseXML && xhr.responseXML.documentElement) {
			return serializeXml(xhr.responseXML);
		}
		return xhr.responseText || "";
	}

	function loadXmlIntoIsland(id, text) {
		var island = dataIsland(id);
		if (island && typeof island.loadXML === "function") {
			return island.loadXML(text);
		}
		if (island && island._itmsXmlIsland && typeof island._itmsXmlIsland.loadXML === "function") {
			return island._itmsXmlIsland.loadXML(text);
		}
		return false;
	}

	function importForDocument(doc, node) {
		if (!doc || !node) {
			return null;
		}
		return doc.importNode ? doc.importNode(node, true) : node.cloneNode(true);
	}

	function syncGet(url) {
		var xhr = new XMLHttpRequest();
		xhr.open("GET", url, false);
		xhr.send(null);
		return xhr;
	}

	function syncPost(url, body) {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", url, false);
		xhr.send(body == null ? null : body);
		return xhr;
	}

	function openDialog(url, args, features, callback) {
		ensureCompat();
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			return window.ITMSModernCompat.openModalDialog(url, args, features, callback);
		}
		return window.open(url, "_blank");
	}

	function getDatePicker() {
		ensureCompat();
		return field("ctlDDate");
	}

	function getDatePickerValue() {
		var control = getDatePicker();
		if (!control) {
			return "";
		}
		if (typeof control.GetDate === "function") {
			return control.GetDate();
		}
		if (typeof control.getDate === "function") {
			return control.getDate();
		}
		return control.value || "";
	}

	function setDatePickerValue(value) {
		var control = getDatePicker();
		if (!control || !value) {
			return;
		}
		if (typeof control.SetDate === "function") {
			control.SetDate(value);
		} else if (typeof control.setDate === "function") {
			control.setDate(value);
		} else {
			control.value = value;
		}
	}

	function setDatePickerMin(value) {
		var control = getDatePicker();
		if (!control || !value) {
			return;
		}
		if (typeof control.SetMinDate === "function") {
			control.SetMinDate(value);
		} else if (typeof control.setMinDate === "function") {
			control.setMinDate(value);
		} else {
			control.min = toIsoDate(value);
		}
	}

	function findItem(root, itemCode, classCode) {
		var items = elementChildren(root, "ITEM");
		for (var i = 0; i < items.length; i += 1) {
			if ((getAttr(items[i], "ITEM") || getAttr(items[i], "ICODE")) === trim(itemCode) && (getAttr(items[i], "CLASS") || getAttr(items[i], "CCODE")) === trim(classCode)) {
				return items[i];
			}
		}
		return null;
	}

	function findStorage(itemNode, storeCode, binCode) {
		var stores = elementChildren(itemNode, "STORAGE");
		for (var i = 0; i < stores.length; i += 1) {
			if (getAttr(stores[i], "STORE") === trim(storeCode) && getAttr(stores[i], "BIN") === trim(binCode)) {
				return stores[i];
			}
		}
		return null;
	}

	function updateMatchingQuantityField(fieldPrefix, value) {
		var elements = form() && form().elements || [];
		for (var i = 0; i < elements.length; i += 1) {
			if (String(elements[i].type || "").toLowerCase() === "text" && String(elements[i].name || "").indexOf(fieldPrefix) === 0) {
				elements[i].value = formatNumber(value, 3);
			}
		}
	}

	function updateStoreQuantityDisplay(itemCode, classCode, orgId, storeCode, binCode, balQty, grnDate) {
		var root = dataRoot("Data");
		var item = findItem(root, itemCode, classCode);
		var storage = findStorage(item, storeCode, binCode);
		var qty = storage ? getAttr(storage, "STOREQTY") || getAttr(storage, "QTY") : "";
		var exactPrefix = "txtZ" + classCode + ":" + itemCode + ":" + orgId + ":" + balQty + ":" + grnDate;
		if (qty !== "") {
			updateMatchingQuantityField(exactPrefix, qty);
		}
	}

	function rowReceiptNumber(button) {
		var elements = form() && form().elements || [];
		var index = -1;
		for (var i = 0; i < elements.length; i += 1) {
			if (elements[i] === button) {
				index = i;
				break;
			}
		}
		for (i = index; i >= 0; i -= 1) {
			if (elements[i] && elements[i].name === "hRecNo") {
				return trim(elements[i].value);
			}
		}
		return trim(field("hRec") && field("hRec").value) || "N";
	}

	function buttonRows() {
		var elements = form() && form().elements || [];
		var rows = [];
		var parts;
		for (var i = 0; i < elements.length; i += 1) {
			if (String(elements[i].type || "").toLowerCase() === "button" && String(elements[i].name || "").indexOf("btn:") === 0) {
				parts = String(elements[i].name).split(":");
				rows.push({
					rowNo: rows.length + 1,
					button: elements[i],
					classCode: parts[1] || "",
					itemCode: parts[2] || "",
					orgId: parts[3] || "",
					balQty: parts[4] || "",
					grnDate: parts[5] || "",
					storeCode: parts[6] || "",
					binCode: parts[7] || ""
				});
			}
		}
		return rows;
	}

	function quantityFlagForItem(classCode, itemCode, orgId) {
		var elements = form() && form().elements || [];
		var search = "btn:" + classCode + ":" + itemCode + ":" + orgId;
		var flag = "";
		for (var i = 0; i < elements.length; i += 1) {
			if (String(elements[i].type || "").toLowerCase() === "button" && String(elements[i].name || "").indexOf(search) > -1) {
				if (elements[i - 4] && trim(elements[i - 4].value) !== "N") {
					flag = "0";
				}
			}
		}
		return flag || "N";
	}

	function appendResponseXmlToData(text) {
		var dataDoc = dataDocument("Data");
		var data = dataRoot("Data");
		var xmlDataRoot;
		if (!trim(text) || !dataDoc || !data) {
			return;
		}
		loadXmlIntoIsland("XmlData", text);
		xmlDataRoot = dataRoot("XmlData");
		if (xmlDataRoot) {
			data.appendChild(importForDocument(dataDoc, xmlDataRoot));
		}
	}

	function updateVisibleQuantitiesFromData() {
		var root = dataRoot("Data");
		var receivedOn = getAttr(root, "RECEIVEDON");
		var unit = getAttr(root, "UNIT");
		var items = elementChildren(root, "ITEM");
		var itemCode;
		var classCode;
		var orgId;
		var qtyAcc;
		var stores;
		var qty;
		for (var i = 0; i < items.length; i += 1) {
			itemCode = getAttr(items[i], "ITEM") || getAttr(items[i], "ICODE");
			classCode = getAttr(items[i], "CLASS") || getAttr(items[i], "CCODE");
			orgId = getAttr(items[i], "OCODE") || unit;
			qtyAcc = getAttr(items[i], "QTYACC") || getAttr(items[i], "ITEMQTY");
			stores = elementChildren(items[i], "STORAGE");
			for (var s = 0; s < stores.length; s += 1) {
				qty = getAttr(stores[s], "QTY") || getAttr(stores[s], "STOREQTY");
				updateMatchingQuantityField("txtZ" + classCode + ":" + itemCode + ":" + orgId + ":" + qtyAcc + ":" + receivedOn, qty);
			}
		}
	}

	function validateReceiptDate(todaysdate) {
		var receiptDate = getDatePickerValue();
		var grnDate = field("hGDate") && field("hGDate").value;
		var finFrom = field("hFinFrom") && field("hFinFrom").value;
		var finTo = field("hFinTo") && field("hFinTo").value;

		if (dateDiffDays(grnDate, receiptDate) < 0) {
			alert("Date should be greater than or equal to GRN Date");
			return false;
		}
		if (dateDiffDays(todaysdate, receiptDate) > 0) {
			alert("Date should be less than or equal to " + todaysdate);
			return false;
		}
		if (dateDiffDays(receiptDate, finTo) < 0 || dateDiffDays(receiptDate, finFrom) > 0) {
			alert("Date should be Selected Between " + finFrom + " and " + finTo);
			return false;
		}
		return true;
	}

	function validateStorageQuantities() {
		var root = dataRoot("Data");
		var items = elementChildren(root, "ITEM");
		var total = 0;
		var itemTotal;
		var itemQty;
		var itemValue;
		var stores;
		for (var i = 0; i < items.length; i += 1) {
			itemTotal = 0;
			itemQty = toNumber(getAttr(items[i], "ITEMQTY") || getAttr(items[i], "QTY"));
			itemValue = toNumber(getAttr(items[i], "ITEMVALUE"));
			stores = elementChildren(items[i], "STORAGE");
			for (var s = 0; s < stores.length; s += 1) {
				itemTotal += toNumber(getAttr(stores[s], "STOREQTY") || getAttr(stores[s], "QTY"));
			}
			if (Math.abs(itemQty - itemTotal) > 0.000001) {
				alert("Total Quantity received should be accounted fully for the Item " + getAttr(items[i], "ITEMNAME"));
				return false;
			}
			if (itemValue !== 0 && itemQty === 0) {
				alert("Total Quantity received should be accounted fully for the Item " + getAttr(items[i], "ITEMNAME"));
				return false;
			}
			total += itemTotal;
		}
		if (total <= 0) {
			alert("Cannot account ZERO Quantity");
			return false;
		}
		return true;
	}

	window.checkNumbers = checkNumbers;

	window.DisableTxt = function (obj) {
		var parts = String(obj && obj.name || "").split("n");
		var tare = field("txtTare" + (parts[1] || ""));
		if (!tare) {
			return;
		}
		tare.value = "";
		tare.disabled = obj && obj.value === "I";
	};

	window.GetLot = function (obj, recType, rowNo, storeName) {
		var parts = String(obj && obj.name || "").split(":");
		var classCode = parts[1] || "";
		var itemCode = parts[2] || "";
		var orgId = parts[3] || "";
		var balQty = parts[4] || "";
		var grnDate = parts[5] || "";
		var storeCode = parts[6] || "";
		var binCode = parts[7] || "";
		var type = recType || rowReceiptNumber(obj) || "N";
		var tempValues;

		if (!recType) {
			tempValues = itemCode + ":" + classCode + ":" + orgId + ":" + balQty + ":" + quantityFlagForItem(classCode, itemCode, orgId);
			openDialog("receiptStorePop.asp?sTemp=" + tempValues, dataDocument("Data"), "dialogHeight:350px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No");
			return;
		}

		tempValues = type + ":" + itemCode + ":" + classCode + ":" + orgId + ":" + storeCode + ":" + binCode + ":" + String(storeName || "").replace(/&amp;/g, "and") + "::True";
		openDialog("receiptLotSerPop.asp?sTemp=" + encodeURIComponent(tempValues), dataDocument("Data"), "dialogHeight:370px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No", function () {
			updateStoreQuantityDisplay(itemCode, classCode, orgId, storeCode, binCode, balQty, grnDate);
		});
	};

	window.DisplayItem = function (obj) {
		var orgName = field("hOrgName");
		openDialog("itmDetailsPop.asp?sTemp=" + encodeURIComponent(obj || ""), orgName ? orgName.value : "", "dialogHeight:400px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No");
	};

	window.SelectStore = function (entryNo, itemCode, orgCode) {
		openDialog("StorageSelect.asp?sUnit=" + encodeURIComponent(orgCode) + "&ItemCode=" + encodeURIComponent(itemCode), "", "dialogHeight:200px;dialogWidth:300px;", function (outValue) {
			var root = xmlRoot(outValue);
			var stores = elementChildren(root);
			var storeName;
			var storeBin;
			var splitStore;
			for (var i = 0; i < stores.length; i += 1) {
				storeName = getAttr(stores[i], "STORE");
				storeBin = getAttr(stores[i], "UNITSTORE");
				splitStore = storeBin.split("-");
				if (field("hStoreNo" + entryNo)) {
					field("hStoreNo" + entryNo).value = splitStore[0] || "";
				}
				if (field("hBinNo" + entryNo)) {
					field("hBinNo" + entryNo).value = splitStore[1] || "";
				}
				if (byId("spanStoreName" + entryNo)) {
					byId("spanStoreName" + entryNo).innerHTML = storeName;
				}
			}
		});
	};

	window.GetLotSingle = function (obj, rowNo) {
		var parts = String(obj && obj.name || "").split(":");
		var classCode = parts[1] || "";
		var itemCode = parts[2] || "";
		var orgId = parts[3] || "";
		var balQty = parts[4] || "";
		var grnDate = parts[5] || "";
		var storeCode = field("hStoreNo" + rowNo) && field("hStoreNo" + rowNo).value;
		var binCode = field("hBinNo" + rowNo) && field("hBinNo" + rowNo).value;
		var storeNameElement = byId("spanStoreName" + rowNo);
		var storeName;
		var recType;
		var tempValues;

		if (!trim(storeCode) || trim(storeCode) === "0") {
			alert("Please Selet the Store ");
			return;
		}
		storeName = storeNameElement ? (storeNameElement.innerHTML || storeNameElement.textContent || "") : "";
		storeName = storeName.replace(/&amp;/g, "and");
		recType = rowReceiptNumber(obj) || "N";
		tempValues = recType + ":" + itemCode + ":" + classCode + ":" + orgId + ":" + storeCode + ":" + binCode + ":" + storeName + "::True";

		openDialog("receiptLotSerPop.asp?sTemp=" + encodeURIComponent(tempValues), dataDocument("Data"), "dialogHeight:370px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No", function () {
			updateStoreQuantityDisplay(itemCode, classCode, orgId, storeCode, binCode, balQty, grnDate);
		});
	};

	window.GetLotDetails = function (recNo, itemCode, classCode, orgId, flag) {
		var hRec = field("hRec");
		var xhr = syncGet(
			"XMLGetLotDetails.asp?iRecNo=" + encodeURIComponent(recNo) +
			"&iItem=" + encodeURIComponent(itemCode) +
			"&iClass=" + encodeURIComponent(classCode) +
			"&sOrgID=" + encodeURIComponent(orgId) +
			"&sFlag=" + encodeURIComponent(flag) +
			"&sRecType=" + encodeURIComponent(hRec ? hRec.value : "")
		);
		appendResponseXmlToData(responseXmlText(xhr));
	};

	window.ViewLotDetails = function (obj) {
		openDialog("receiptLotSerDetPop.asp?Temp=" + encodeURIComponent(obj && obj.name || ""), "", "dialogHeight:370px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No");
	};

	window.Init = function () {
		var rows = buttonRows();
		var hFinFrom = field("hFinFrom");
		var formattedDate = field("hFormattedGDate");
		var xhr;
		var text;
		if (hFinFrom) {
			setDatePickerMin(hFinFrom.value);
		}
		if (formattedDate) {
			setDatePickerValue(formattedDate.value);
		}
		for (var i = 0; i < rows.length; i += 1) {
			if (rows[i].grnDate) {
				setDatePickerValue(rows[i].grnDate);
			}
			xhr = syncGet("XMLReceiptQuantity.asp?Item=" + encodeURIComponent(rows[i].itemCode) + "&Class=" + encodeURIComponent(rows[i].classCode) + "&OrgID=" + encodeURIComponent(rows[i].orgId) + "&ReceiptNo=" + encodeURIComponent(field("hiRecNo") && field("hiRecNo").value || ""));
			text = responseXmlText(xhr);
			if (trim(text)) {
				loadXmlIntoIsland("TempData", text);
			} else {
				alert("File is Not Found or It is returning Invalid Data - XMLReceiptQuantity.asp");
			}
		}
		updateVisibleQuantitiesFromData();
	};

	window.CheckSubmit = function (todaysdate) {
		var root = dataRoot("Data");
		var response;
		var saveButton = field("B1");

		if (!validateReceiptDate(todaysdate)) {
			return;
		}
		if (!validateStorageQuantities()) {
			return;
		}

		setAttr(root, "RECEIVEDON", getDatePickerValue());
		response = syncPost("receiptInsert.asp", serializeXml(dataDocument("Data")));
		if (trim(response.responseText) === "") {
			alert("Receipt has been done Successfully");
			window.location.href = "MATERIALRECEIPTS.ASP?RCPT=A";
		} else {
			alert(response.responseText);
			if (saveButton) {
				saveButton.disabled = false;
			}
		}
	};
}(window, document));
