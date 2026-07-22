(function (window, document) {
	"use strict";

	var objTemp = null;
	var rootNode = null;
	var lineNo = "";
	var pickNo = "";
	var lotNo = "";
	var storeCode = "";
	var binCode = "";
	var scheduleMinDate = "";
	var scheduleMaxDate = "";

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

	function radio(index) {
		var radios = field("radDate");
		if (!radios) {
			return null;
		}
		return radios.length ? radios[index] : (index === 0 ? radios : null);
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
			if (date.getFullYear() === year && date.getMonth() === Number(match[2]) - 1 && date.getDate() === Number(match[1])) {
				return date;
			}
			return null;
		}
		match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
		if (match) {
			date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
			if (date.getFullYear() === Number(match[1]) && date.getMonth() === Number(match[2]) - 1 && date.getDate() === Number(match[3])) {
				return date;
			}
		}
		return null;
	}

	function validDate(value) {
		return !!parseDate(value);
	}

	function dateDiffDays(start, end) {
		var startDate = parseDate(start);
		var endDate = parseDate(end);
		if (!startDate || !endDate) {
			return 0;
		}
		return Math.floor((endDate.getTime() - startDate.getTime()) / 86400000);
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

	function attrAt(node, index) {
		var attr = node && node.attributes && node.attributes.item(index);
		return trim(attr ? attr.nodeValue || attr.value || "" : "");
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

	function modalArgs() {
		ensureCompat();
		return window["dialog" + "Arguments"] || null;
	}

	function getHeaderNode() {
		var nodes = elementChildren(rootNode);
		for (var i = 0; i < nodes.length; i += 1) {
			if (attrAt(nodes[i], 0) === trim(lineNo) && attrAt(nodes[i], 1) === trim(pickNo)) {
				return nodes[i];
			}
		}
		return null;
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
		if (!control) {
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

	function setDatePickerBound(name, value) {
		var control = getDatePicker();
		var methodName = name === "min" ? "SetMinDate" : "SetMaxDate";
		var fallbackName = name === "min" ? "setMinDate" : "setMaxDate";
		var date = parseDate(value);
		if (!control || !value) {
			return;
		}
		if (typeof control[methodName] === "function") {
			control[methodName](value);
		} else if (typeof control[fallbackName] === "function") {
			control[fallbackName](value);
		} else if (date) {
			control[name] = date.getFullYear() + "-" + pad2(date.getMonth() + 1) + "-" + pad2(date.getDate());
		}
	}

	function focusOrSelect(control) {
		if (!control) {
			return;
		}
		if (typeof control.select === "function") {
			control.select();
		} else if (typeof control.focus === "function") {
			control.focus();
		}
	}

	function selectedDateMode() {
		if (radio(0) && radio(0).checked) {
			return "M";
		}
		if (radio(1) && radio(1).checked) {
			return "S";
		}
		return "";
	}

	function syncSingleDateRows() {
		var count = toNumber(field("hiCtr") && field("hiCtr").value);
		var value = getDatePickerValue();
		var dateField;
		for (var i = 1; i <= count; i += 1) {
			dateField = field("txtDate" + i);
			if (dateField) {
				dateField.value = value;
				dateField.readOnly = true;
			}
		}
	}

	function applySelectedDateMode() {
		var mode = selectedDateMode();
		var count = toNumber(field("hiCtr") && field("hiCtr").value);
		var dateField;
		for (var i = 1; i <= count; i += 1) {
			dateField = field("txtDate" + i);
			if (!dateField) {
				continue;
			}
			if (mode === "S") {
				dateField.value = getDatePickerValue();
				dateField.readOnly = true;
			} else if (mode === "M") {
				dateField.readOnly = false;
			}
		}
	}

	function setReturnValue() {
		if (!rootNode) {
			rootNode = xmlRoot(objTemp);
		}
		if (!rootNode) {
			return;
		}
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(rootNode);
		} else {
			window["return" + "Value"] = rootNode;
			window.returnvalue = rootNode;
		}
	}

	function closeWithReturn() {
		setReturnValue();
		window.close();
	}

	function removeExistingSerialHeader(headerNode) {
		var serialHeaders = elementChildren(headerNode, "SERIALHEADER");
		for (var i = 0; i < serialHeaders.length; i += 1) {
			headerNode.removeChild(serialHeaders[i]);
		}
	}

	function requestedQuantity() {
		var quantityElement = document.getElementById("idQty");
		return toNumber(quantityElement ? quantityElement.textContent || quantityElement.innerText : 0);
	}

	function validateManualDate(dateField, todaysdate) {
		var dateValue = trim(dateField && dateField.value);
		if (!dateValue) {
			alert("Enter Date");
			focusOrSelect(dateField);
			return false;
		}
		if (!validDate(dateValue)) {
			alert("Invalid date");
			focusOrSelect(dateField);
			return false;
		}
		if (dateDiffDays(scheduleMinDate, dateValue) < 0) {
			alert("Date should be greater than or equal to Schedule Date");
			focusOrSelect(dateField);
			return false;
		}
		if (dateDiffDays(todaysdate, dateValue) > 0) {
			alert("Date should be less than or equal to Today's Date");
			focusOrSelect(dateField);
			return false;
		}
		return true;
	}

	function populateExistingSerialDetails(headerNode) {
		var serialHeaders = elementChildren(headerNode, "SERIALHEADER");
		var details;
		var serialHeader;
		var qtyField;
		var dateField;
		var i;

		if (!serialHeaders.length) {
			return;
		}

		serialHeader = serialHeaders[0];
		if (attrAt(serialHeader, 0) === "M") {
			if (radio(0)) {
				radio(0).checked = true;
			}
		} else if (radio(1)) {
			radio(1).checked = true;
		}
		setDatePickerValue(attrAt(serialHeader, 1));

		details = elementChildren(serialHeader, "SERIALDETAILS");
		for (i = 0; i < details.length; i += 1) {
			qtyField = field("txtQty" + (i + 1));
			dateField = field("txtDate" + (i + 1));
			if (qtyField) {
				qtyField.value = attrAt(details[i], 1);
			}
			if (dateField) {
				dateField.value = attrAt(details[i], 2);
			}
		}
		applySelectedDateMode();
	}

	window.ChangeDet = function (radValue) {
		var count = toNumber(field("hiCtr") && field("hiCtr").value);
		var dateField;
		for (var i = 1; i <= count; i += 1) {
			dateField = field("txtDate" + i);
			if (!dateField) {
				continue;
			}
			dateField.value = "";
			if (radValue === "M") {
				dateField.readOnly = false;
			} else {
				dateField.value = getDatePickerValue();
				dateField.readOnly = true;
			}
		}
	};

	window.fnInit = function (arg, maxDate) {
		var parts;
		var headerNode;
		var datePicker = getDatePicker();

		if (trim(field("hiCtr") && field("hiCtr").value) === "0") {
			return;
		}

		objTemp = modalArgs();
		rootNode = xmlRoot(objTemp);
		parts = String(arg || "").split("`");

		lineNo = parts[1] || "";
		pickNo = parts[2] || "";
		lotNo = parts[3] || "";
		storeCode = parts[4] || "";
		binCode = parts[5] || "";

		headerNode = getHeaderNode();
		if (headerNode) {
			scheduleMinDate = attrAt(headerNode, 7);
			scheduleMaxDate = attrAt(headerNode, 8);
		}

		setDatePickerBound("max", maxDate);
		setDatePickerBound("min", scheduleMinDate);

		if (datePicker) {
			datePicker.addEventListener("change", function () {
				if (selectedDateMode() === "S") {
					syncSingleDateRows();
				}
			});
			datePicker.addEventListener("blur", function () {
				if (selectedDateMode() === "S") {
					syncSingleDateRows();
				}
			});
		}

		if (headerNode) {
			populateExistingSerialDetails(headerNode);
		}
	};

	window.CheckSubmit = function (todaysdate) {
		var doc = xmlDocument(objTemp);
		var count = toNumber(field("hiCtr") && field("hiCtr").value);
		var mode = selectedDateMode();
		var totalQty = 0;
		var headerNode;
		var serialHeader;
		var detail;
		var qtyField;
		var stockQtyField;
		var dateField;
		var serialField;
		var qty;
		var stockQty;
		var dateValue;
		var i;

		if (count === 0) {
			return;
		}

		if (!mode) {
			alert("Select Date");
			if (radio(0)) {
				radio(0).focus();
			}
			return;
		}

		if (mode === "S") {
			dateValue = getDatePickerValue();
			if (dateDiffDays(scheduleMinDate, dateValue) < 0) {
				alert("Date should be greater than or equal to Schedule Date");
				return;
			}
			if (dateDiffDays(todaysdate, dateValue) > 0) {
				alert("Date should be less than or equal to Today's Date");
				return;
			}
			syncSingleDateRows();
		}

		for (i = 1; i <= count; i += 1) {
			qtyField = field("txtQty" + i);
			stockQtyField = field("txtStQty" + i);
			dateField = field("txtDate" + i);

			if (!trim(qtyField && qtyField.value)) {
				alert("Enter Quantity");
				focusOrSelect(qtyField);
				return;
			}
			if (!checkNumbers(qtyField.value)) {
				alert("Enter Numerals Only");
				focusOrSelect(qtyField);
				return;
			}

			qty = toNumber(qtyField.value);
			stockQty = toNumber(stockQtyField && stockQtyField.value);

			if (mode === "M" && qty > 0 && !validateManualDate(dateField, todaysdate)) {
				return;
			}
			if (qty > stockQty) {
				alert("Quantity Issue should be equal to or less than Stock Quantity");
				focusOrSelect(qtyField);
				return;
			}

			totalQty += qty;
		}

		headerNode = getHeaderNode();
		if (!doc || !headerNode) {
			return;
		}
		removeExistingSerialHeader(headerNode);

		serialHeader = doc.createElement("SERIALHEADER");
		serialHeader.setAttribute("DATETYPE", mode);
		serialHeader.setAttribute("DATEVALUE", getDatePickerValue());

		for (i = 1; i <= count; i += 1) {
			serialField = field("hSerial" + i);
			qtyField = field("txtQty" + i);
			dateField = field("txtDate" + i);
			detail = doc.createElement("SERIALDETAILS");
			detail.setAttribute("SERIALNO", trim(serialField && serialField.value));
			detail.setAttribute("QTY", trim(qtyField && qtyField.value));
			detail.setAttribute("TDATE", trim(dateField && dateField.value));
			serialHeader.appendChild(detail);
		}
		headerNode.appendChild(serialHeader);

		if (totalQty > requestedQuantity()) {
			alert("Total Quantity Issue should be equal to or less than Quantity Requested");
			return;
		}

		headerNode.setAttribute("ISSQTY", String(totalQty));
		closeWithReturn();
	};

	window.window_onunload = setReturnValue;
	window.addEventListener("beforeunload", setReturnValue);
}(window, document));
