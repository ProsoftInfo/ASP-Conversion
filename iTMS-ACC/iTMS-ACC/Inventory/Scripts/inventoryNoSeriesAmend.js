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
		return document.getElementById(name) || (frm && frm.elements ? frm.elements[name] : null);
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

	function focusField(name) {
		var item = field(name);
		if (item && item.focus) {
			item.focus();
		}
	}

	function xmlFromResponse(xhr) {
		if (xhr.responseXML && xhr.responseXML.documentElement) {
			return xhr.responseXML;
		}
		return new DOMParser().parseFromString(xhr.responseText || "<ROOT/>", "application/xml");
	}

	function attr(node, name) {
		return node && node.getAttribute ? node.getAttribute(name) || "" : "";
	}

	function appendCell(row, value, className, align, width) {
		var cell = row.insertCell();
		cell.textContent = value == null ? "" : String(value);
		cell.className = className || "";
		if (align) {
			cell.align = align;
		}
		if (width) {
			cell.width = width;
		}
		return cell;
	}

	function appendTextInput(row, name, value, className, readOnly) {
		var cell = row.insertCell();
		var input = document.createElement("input");
		input.type = "text";
		input.name = name;
		input.size = 12;
		input.value = value || "";
		input.maxLength = 10;
		input.className = className || "FormElem";
		input.style.textAlign = "left";
		input.readOnly = !!readOnly;
		cell.appendChild(input);
		cell.width = "10";
		cell.className = readOnly ? "ExcelDisplayCell" : "ExcelInputCell";
		return cell;
	}

	function CheckPeriod(sDate) {
		var today = new Date();
		var month = today.getMonth() + 1;
		var year = today.getFullYear();
		var fromDate;
		var toDate;
		if (month <= 3) {
			fromDate = String(year - 1) + "04";
			toDate = String(year) + "03";
		} else {
			fromDate = String(year) + "04";
			toDate = String(year + 1) + "03";
		}
		return Number(sDate) >= Number(fromDate) && Number(sDate) <= Number(toDate);
	}

	function ClearTable() {
		var table = document.getElementById("tblBook");
		var row;
		if (!table) {
			return;
		}
		while (table.rows.length) {
			table.deleteRow(0);
		}
		row = table.insertRow(0);
		appendCell(row, "S.No.", "ExcelHeaderCell", "center", 10);
		appendCell(row, "Period", "ExcelHeaderCell", "center", 100);
		appendCell(row, "Start No", "ExcelHeaderCell", "center", 50);
		appendCell(row, "Prefix", "ExcelHeaderCell", "center", 100);
		appendCell(row, "Suffix", "ExcelHeaderCell", "center", 100);
	}

	function DisplayTable() {
		var unit;
		var activity;
		var categoryCode;
		var classCode;
		var xhr;
		var xml;
		var root;
		var nodes;
		var row;
		var node;
		var editableCount = 0;
		var serial = 0;
		ClearTable();
		if (field("selUnit").selectedIndex === 0) {
			alert("Select Unit");
			focusField("selUnit");
			return false;
		}
		if (field("selActType").selectedIndex === 0) {
			alert("Select Activity");
			focusField("selActType");
			return false;
		}
		unit = valueOf("selUnit");
		activity = valueOf("selActType");
		categoryCode = valueOf("hCatCode");
		classCode = valueOf("hClassCode");
		if (trim(categoryCode) === "" || trim(categoryCode) === "0") {
			categoryCode = valueOf("selCategory");
		}
		xhr = new XMLHttpRequest();
		xhr.open("GET", "XMLNoSeriesSelection.asp?sUnit=" + encodeURIComponent(unit) + "&sItem=&sActivity=" + encodeURIComponent(activity) + "&CatCode=" + encodeURIComponent(categoryCode) + "&ClassCode=" + encodeURIComponent(classCode), false);
		xhr.send();
		if (trim(xhr.responseText) === "") {
			setValue("hSeriesLen", "");
			alert("No Number Series Allocated");
			return false;
		}
		xml = xmlFromResponse(xhr);
		root = xml.documentElement;
		nodes = root ? root.childNodes : [];
		for (var i = 0; i < nodes.length; i += 1) {
			node = nodes[i];
			if (node.nodeType !== 1) {
				continue;
			}
			serial += 1;
			row = document.getElementById("tblBook").insertRow(document.getElementById("tblBook").rows.length);
			if (CheckPeriod(attr(node, "PERIOD"))) {
				editableCount += 1;
				appendCell(row, serial, "ExcelSerial", "center");
				appendTextInput(row, "txtPeriod" + editableCount, attr(node, "PERIOD"), "FormElemRead", true);
				appendTextInput(row, "txtStartNo" + editableCount, attr(node, "NUMBER"), "FormElem", false);
				appendTextInput(row, "txtPrefix" + editableCount, attr(node, "PREFIX"), "FormElem", false);
				appendTextInput(row, "txtSuffix" + editableCount, attr(node, "SUFFIX"), "FormElem", false);
			} else {
				appendCell(row, serial, "ExcelSerial", "center");
				appendCell(row, attr(node, "PERIOD"), "ExcelDisplayCell", "left");
				appendCell(row, attr(node, "NUMBER"), "ExcelDisplayCell", "center");
				appendCell(row, attr(node, "PREFIX"), "ExcelDisplayCell", "left");
				appendCell(row, attr(node, "SUFFIX"), "ExcelDisplayCell", "left");
			}
		}
		setValue("hSeriesLen", editableCount || "");
		return false;
	}

	function ClearAll() {
		if (field("selActType")) {
			field("selActType").selectedIndex = 0;
		}
		if (field("selUnit")) {
			field("selUnit").selectedIndex = 0;
		}
		ClearTable();
		return false;
	}

	function validateForm() {
		if (field("selUnit").selectedIndex === 0) {
			alert("Select Unit");
			focusField("selUnit");
			return false;
		}
		if (field("selActType").selectedIndex === 0) {
			alert("Select Activity");
			focusField("selActType");
			return false;
		}
		if (trim(valueOf("hSeriesLen")) === "") {
			alert("No Details found for this Activity");
			focusField("selActType");
			return false;
		}
		form().submit();
		return false;
	}

	function parseClassification(returnData) {
		var parts = String(returnData || "").split("*****");
		var classParts;
		var nameParts;
		var classCodes = [];
		var categoryCodes = [];
		var classNames = [];
		var select;
		var classItem;
		var valueItem;
		if (parts[0] === "-1" || !parts[0]) {
			return;
		}
		classParts = parts[0].split("|");
		for (var i = 0; i < classParts.length; i += 1) {
			valueItem = classParts[i].split(":");
			if (valueItem.length > 1) {
				classCodes.push(valueItem[valueItem.length - 1]);
				categoryCodes.push(valueItem[1]);
			} else {
				categoryCodes.push(valueItem[0].substring(3));
			}
		}
		nameParts = (parts[1] || "").split("|||");
		for (var j = 0; j < nameParts.length; j += 1) {
			classItem = nameParts[j].split(":");
			if (trim(classItem[classItem.length - 1]) !== "") {
				classNames.push(classItem[classItem.length - 1]);
			}
		}
		if (document.getElementById("txtClass")) {
			document.getElementById("txtClass").textContent = classNames.join(",");
		}
		setValue("hClassCode", classCodes.join(","));
		setValue("hCatCode", categoryCodes.join(","));
		select = field("selCategory");
		if (select && categoryCodes.length) {
			for (var k = 0; k < select.options.length; k += 1) {
				if (trim(select.options[k].value) === trim(categoryCodes.join(","))) {
					select.selectedIndex = k;
					select.disabled = true;
					break;
				}
			}
		}
	}

	function SelectClassifcation() {
		var orgId = valueOf("selUnit");
		var url = "/include/ClassificationSelectPop.asp?sIType=1&sOrgID=" + encodeURIComponent(orgId) + "&sITypename=&SelMode=M";
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, "Classification", "dialogHeight:500px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No", parseClassification);
		}
		return false;
	}

	window.DisplayTable = DisplayTable;
	window.CheckPeriod = CheckPeriod;
	window.ClearAll = ClearAll;
	window.ClearTable = ClearTable;
	window.validateForm = validateForm;
	window.SelectClassifcation = SelectClassifcation;
}(window, document));
