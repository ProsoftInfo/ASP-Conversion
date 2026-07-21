(function (window, document) {
	"use strict";

	var objTemp = null;
	var rootNode = null;
	var itemCode = "";
	var classCode = "";
	var requiredQty = 0;
	var entryNo = "";
	var scheduleNodeName = "Schedule";

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

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function checkNumbers(value, allowDecimal) {
		return (allowDecimal ? /^[0-9.]+$/ : /^[0-9]+$/).test(String(value || ""));
	}

	function pad2(value) {
		return value < 10 ? "0" + value : String(value);
	}

	function parseDate(value) {
		var text = trim(value);
		var match;
		var year;
		if (!text) {
			return null;
		}
		match = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
		if (match) {
			year = Number(match[3]);
			if (year < 100) {
				year += 2000;
			}
			return new Date(year, Number(match[2]) - 1, Number(match[1]));
		}
		match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
		if (match) {
			return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
		}
		return null;
	}

	function validDate(value) {
		var date = parseDate(value);
		return !!date && !isNaN(date.getTime()) && pad2(date.getDate()) + "/" + pad2(date.getMonth() + 1) + "/" + date.getFullYear() === trim(value).replace(/[.-]/g, "/");
	}

	function currentDate(todaysdate) {
		return parseDate(todaysdate) || new Date();
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

	function selectedScheduleType() {
		var select = field("selSchtype");
		return select && select.selectedIndex > 0 ? select.options[select.selectedIndex].value : "";
	}

	function selectedScheduleIndex() {
		var select = field("selSchtype");
		return select ? select.selectedIndex : 0;
	}

	function getHeaderNode() {
		var nodes = elementChildren(rootNode);
		for (var i = 0; i < nodes.length; i += 1) {
			if (attrAt(nodes[i], 0) === trim(entryNo) && attrAt(nodes[i], 2) === trim(classCode) && attrAt(nodes[i], 1) === trim(itemCode)) {
				return nodes[i];
			}
		}
		return null;
	}

	function setReturnValue() {
		if (!rootNode) {
			rootNode = xmlRoot(objTemp);
		}
		window.returnValue = rootNode;
		window.returnvalue = rootNode;
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(rootNode);
		}
	}

	function closeWithReturn() {
		setReturnValue();
		window.close();
	}

	function validateNeedAndQuantity(row, todaysdate, seen) {
		var dateField = field("txtD" + row);
		var qtyField = field("txtQ" + row);
		var need = trim(dateField && dateField.value);
		var qty = trim(qtyField && qtyField.value);
		var typeIndex = selectedScheduleIndex();
		var today = currentDate(todaysdate);
		var month;
		var week;
		var year;

		if (!need) {
			return 0;
		}

		if (typeIndex === 1) {
			if (!validDate(need)) {
				alert("Invalid Date");
				dateField.select();
				return false;
			}
			if (dateDiffDays(todaysdate, need) < 0) {
				alert("Date should be greater or equal to Today's Date");
				dateField.select();
				return false;
			}
		} else if (typeIndex === 2) {
			if (need.length !== 6 || !checkNumbers(need, false)) {
				alert("Invalid Format (Month 2 and Year 4 characters)");
				dateField.select();
				return false;
			}
			month = Number(need.substr(0, 2));
			year = Number(need.substr(2));
			if (month < 1 || month > 12) {
				alert("Invalid Month");
				dateField.select();
				return false;
			}
			if (year < today.getFullYear()) {
				alert("Year should be greater than or equal to Current Year");
				dateField.select();
				return false;
			}
			if (month < today.getMonth() + 1 && year === today.getFullYear()) {
				alert("Month should be greater than or equal to Current month");
				dateField.select();
				return false;
			}
		} else if (typeIndex === 3) {
			if (need.length !== 6 || !checkNumbers(need, false)) {
				alert("Invalid Format (Week 2 and Year 4 characters)");
				dateField.select();
				return false;
			}
			week = Number(need.substr(0, 2));
			year = Number(need.substr(2));
			if (week < 1 || week > 52) {
				alert("Invalid Week");
				dateField.select();
				return false;
			}
			if (year < today.getFullYear()) {
				alert("Year should be greater than or equal to Current Year");
				dateField.select();
				return false;
			}
		} else if (typeIndex === 4) {
			if (need.length !== 8 || !checkNumbers(need, false)) {
				alert("Invalid Format (Month 2, Week 2 and Year 4 characters)");
				dateField.select();
				return false;
			}
			month = Number(need.substr(0, 2));
			week = Number(need.substr(2, 2));
			year = Number(need.substr(4));
			if (month < 1 || month > 12) {
				alert("Invalid Month");
				dateField.select();
				return false;
			}
			if (week < 1 || week > 4) {
				alert("Invalid Week");
				dateField.select();
				return false;
			}
			if (year < today.getFullYear()) {
				alert("Year should be greater than or equal to Current Year");
				dateField.select();
				return false;
			}
			if (month < today.getMonth() + 1 && year === today.getFullYear()) {
				alert("Month should be greater than or equal to Current month");
				dateField.select();
				return false;
			}
		}

		if (!qty) {
			alert("Enter Quantity");
			qtyField.select();
			return false;
		}
		if (!checkNumbers(qty, true)) {
			alert("Enter Numerals Only");
			qtyField.select();
			return false;
		}
		if (seen[need]) {
			alert("No duplication date allowed");
			dateField.select();
			return false;
		}
		seen[need] = true;
		return toNumber(qty);
	}

	window.fnInit = function (sItem, sClass, sQty, sEntNo, sScheduleNodeName) {
		itemCode = sItem;
		classCode = sClass;
		requiredQty = toNumber(sQty);
		entryNo = sEntNo;
		scheduleNodeName = trim(sScheduleNodeName) || "Schedule";
		objTemp = window.dialogArguments;
		rootNode = xmlRoot(objTemp);
	};

	window.setMax = function (objp) {
		var maxLength = 10;
		if (objp && (objp.selectedIndex === 2 || objp.selectedIndex === 3)) {
			maxLength = 6;
		} else if (objp && objp.selectedIndex === 4) {
			maxLength = 8;
		}
		for (var i = 1; i <= 12; i += 1) {
			if (field("txtD" + i)) {
				field("txtD" + i).value = "";
				field("txtD" + i).maxLength = maxLength;
			}
			if (field("txtQ" + i)) {
				field("txtQ" + i).value = "";
			}
		}
	};

	window.CheckSubmit = function (todaysdate) {
		var doc = xmlDocument(objTemp);
		var headerNode;
		var schedules;
		var detail;
		var total = 0;
		var seen = {};
		var result;
		var i;
		var sType = selectedScheduleType();

		if (!sType) {
			alert("Select Schedule Type");
			if (field("selSchtype")) {
				field("selSchtype").focus();
			}
			return;
		}

		for (i = 1; i <= 12; i += 1) {
			result = validateNeedAndQuantity(i, todaysdate, seen);
			if (result === false) {
				return;
			}
			total += result;
		}

		if (total !== requiredQty) {
			alert("Total Schedule Quantity should be equal to Quantity Required");
			return;
		}

		headerNode = getHeaderNode();
		if (!doc || !headerNode) {
			return;
		}
		schedules = elementChildren(headerNode, scheduleNodeName);
		if (!schedules.length) {
			schedules = [doc.createElement(scheduleNodeName)];
			headerNode.appendChild(schedules[0]);
		}

		for (var s = 0; s < schedules.length; s += 1) {
			for (i = 1; i <= 12; i += 1) {
				detail = doc.createElement("ScheduleDetails");
				detail.setAttribute("SNO", i);
				detail.setAttribute("NEED", field("txtD" + i) ? field("txtD" + i).value : "");
				detail.setAttribute("QTY", field("txtQ" + i) ? field("txtQ" + i).value : "");
				detail.setAttribute("TYPE", sType);
				schedules[s].appendChild(detail);
			}
		}
		closeWithReturn();
	};

	window.window_onunload = setReturnValue;
	window.addEventListener("beforeunload", setReturnValue);
}(window, document));
