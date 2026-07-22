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
		if (!target) {
			return "";
		}
		return new XMLSerializer().serializeToString(target);
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

	function selectValue(select, value) {
		if (!select) {
			return;
		}
		select.value = value;
		if (select.value !== value) {
			for (var i = 0; i < select.options.length; i += 1) {
				if (select.options[i].value === value) {
					select.selectedIndex = i;
					break;
				}
			}
		}
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
					qty: parts[4] || "",
					receiptNumbering: parts[5] || ""
				});
			}
		}
		return rows;
	}

	function findItem(root, itemCode, classCode, orgId) {
		var items = elementChildren(root, "ITEM");
		for (var i = 0; i < items.length; i += 1) {
			if (getAttr(items[i], "ICODE") === trim(itemCode) && getAttr(items[i], "CCODE") === trim(classCode) && getAttr(items[i], "OCODE") === trim(orgId)) {
				return items[i];
			}
		}
		return null;
	}

	function quantityFlagForItem(classCode, itemCode, orgId) {
		var rows = buttonRows();
		var flag = "";
		var select;
		for (var i = 0; i < rows.length; i += 1) {
			if (rows[i].classCode === trim(classCode) && rows[i].itemCode === trim(itemCode) && rows[i].orgId === trim(orgId)) {
				select = field("selQtyIn" + rows[i].rowNo);
				if (select && trim(select.value) !== "N") {
					flag = "0";
				}
			}
		}
		return flag || "N";
	}

	function updateJobWorkDetailAttributes() {
		var root = dataRoot("Data");
		var rows = buttonRows();
		var item;
		var rowNo;
		for (var i = 0; i < rows.length; i += 1) {
			rowNo = rows[i].rowNo;
			item = findItem(root, rows[i].itemCode, rows[i].classCode, rows[i].orgId);
			if (!item) {
				continue;
			}
			setAttr(item, "PRUNFROM", trim(field("txtPressRunFro" + rowNo) && field("txtPressRunFro" + rowNo).value));
			setAttr(item, "PRUNTO", trim(field("txtPressRunTo" + rowNo) && field("txtPressRunTo" + rowNo).value));
			setAttr(item, "QTYIN", trim(field("selQtyIn" + rowNo) && field("selQtyIn" + rowNo).value));
			setAttr(item, "PRESS", trim(field("txtPressMark" + rowNo) && field("txtPressMark" + rowNo).value));
			setAttr(item, "CROP", trim(field("txtCropYear" + rowNo) && field("txtCropYear" + rowNo).value));
		}
	}

	function validateStorageQuantities() {
		var root = dataRoot("Data");
		var items = elementChildren(root, "ITEM");
		var total = 0;
		var itemTotal;
		var stores;
		var itemQty;
		for (var i = 0; i < items.length; i += 1) {
			itemTotal = 0;
			itemQty = toNumber(getAttr(items[i], "QTY"));
			stores = elementChildren(items[i], "STORAGE");
			for (var s = 0; s < stores.length; s += 1) {
				itemTotal += toNumber(getAttr(stores[s], "QTY"));
			}
			if (Math.abs(itemQty - itemTotal) > 0.000001) {
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

	window.GetLot = function (obj) {
		var parts = String(obj && obj.name || "").split(":");
		var classCode = parts[1] || "";
		var itemCode = parts[2] || "";
		var orgId = parts[3] || "";
		var qty = parts[4] || "";
		var rec = parts[5] || "";
		var tempValues = itemCode + ":" + classCode + ":" + orgId + ":" + qty + ":" + quantityFlagForItem(classCode, itemCode, orgId) + ":" + rec;
		openDialog("JobWorkReceiptStorePop.asp?sTemp=" + tempValues, dataDocument("Data"), "dialogHeight:350px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No");
	};

	window.DisplayItem = function (obj) {
		var orgName = field("hOrgName");
		openDialog("itmDetailsPop.asp?sTemp=" + encodeURIComponent(obj || ""), orgName ? orgName.value : "", "dialogHeight:400px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No");
	};

	window.CheckSubmit = function (todaysdate) {
		var root = dataRoot("Data");
		var receiptDate = getDatePickerValue();
		var saveButton = field("B1");
		var saveResponse;
		var insertResponse;

		if (dateDiffDays(field("hGDate") && field("hGDate").value, receiptDate) < 0) {
			alert("Date should be greater than or equal to GRN Date");
			return;
		}
		if (dateDiffDays(todaysdate, receiptDate) > 0) {
			alert("Date should be less than or equal to Today's Date");
			return;
		}

		updateJobWorkDetailAttributes();
		if (!validateStorageQuantities()) {
			return;
		}

		setAttr(root, "RECEIVEDON", receiptDate);
		saveResponse = syncPost("XMLSave.asp?SessionFlag=False&Value=JobWorkReceipt&Folder=Transaction", serializeXml(dataDocument("Data")));
		if (saveResponse.status >= 400) {
			alert(saveResponse.responseText || "Unable to save receipt XML.");
			return;
		}

		if (saveButton) {
			saveButton.disabled = true;
		}
		insertResponse = syncPost("JobWorkReceiptInsert.asp", serializeXml(dataDocument("Data")));
		if (insertResponse.responseText === "Y<BR>") {
			alert("Receipt has been done Successfully");
			if (window.confirm("Do you want to create another Receipt")) {
				window.location.href = "MATERIALRECEIPTS.ASP";
			} else {
				window.location.href = "../welcome_Inventory.asp";
			}
		} else {
			alert(insertResponse.responseText);
			if (saveButton) {
				saveButton.disabled = false;
			}
		}
	};

	window.Init = function () {
		var rows = buttonRows();
		var hGDate = field("hGDate");
		var row;
		var xhr;
		var root;
		var text;
		var qtyIn;

		if (hGDate) {
			setDatePickerValue(hGDate.value);
		}

		for (var i = 0; i < rows.length; i += 1) {
			row = rows[i];
			xhr = syncGet("XMLJobWorkReceiptQuantity.asp?Item=" + encodeURIComponent(row.itemCode) + "&Class=" + encodeURIComponent(row.classCode) + "&OrgID=" + encodeURIComponent(row.orgId) + "&ReceiptNo=" + encodeURIComponent(field("hiRecNo") && field("hiRecNo").value || ""));
			text = responseXmlText(xhr);
			if (!trim(text)) {
				alert("File is Not Found or It is returning Invalid Data - XMLJobWorkReceiptQuantity.asp");
				continue;
			}
			loadXmlIntoIsland("TempData", text);
			root = dataRoot("TempData");
			if (!root) {
				continue;
			}
			qtyIn = getAttr(root, "QUANTITYIN");
			if (qtyIn === "G" || qtyIn === "N") {
				selectValue(field("selQtyIn" + row.rowNo), qtyIn);
				if (field("selQtyIn" + row.rowNo)) {
					field("selQtyIn" + row.rowNo).disabled = true;
				}
			}
			if (field("txtPressMark" + row.rowNo)) {
				field("txtPressMark" + row.rowNo).value = getAttr(root, "PRESSMARKNO");
			}
			if (field("txtPressRunFro" + row.rowNo)) {
				field("txtPressRunFro" + row.rowNo).value = getAttr(root, "PRNOFROM");
			}
			if (field("txtCropYear" + row.rowNo)) {
				field("txtCropYear" + row.rowNo).value = getAttr(root, "CROPYEAR");
			}
			if (field("txtPressRunTo" + row.rowNo)) {
				field("txtPressRunTo" + row.rowNo).value = getAttr(root, "PRNOTO");
			}
		}
	};
}(window, document));
