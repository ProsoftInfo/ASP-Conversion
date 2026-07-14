(function (window, document) {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.formname || document.forms.formname || document.forms[0] || null;
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] : document.getElementById(name);
	}

	function fields(name) {
		var item = field(name);
		if (!item) {
			return [];
		}
		if (item.length != null && !item.tagName) {
			return Array.prototype.slice.call(item);
		}
		return [item];
	}

	function selectedText(select) {
		return select && select.options && select.selectedIndex >= 0 ? select.options[select.selectedIndex].text : "";
	}

	function selectedValue(select) {
		return select && select.options && select.selectedIndex >= 0 ? select.options[select.selectedIndex].value : select && select.value || "";
	}

	function setValue(name, value) {
		var item = field(name);
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function setDisabled(name, disabled) {
		var item = field(name);
		if (item) {
			item.disabled = !!disabled;
		}
	}

	function setText(id, value) {
		var item = document.getElementById(id) || field(id);
		if (item) {
			item.textContent = value == null ? "" : String(value);
		}
	}

	function focusField(name) {
		var item = field(name);
		if (item && item.focus) {
			item.focus();
		}
	}

	function radioChecked(name, index) {
		var list = fields(name);
		return !!(list[index] && list[index].checked);
	}

	function setRadio(name, index, checked) {
		var list = fields(name);
		if (list[index]) {
			list[index].checked = !!checked;
		}
	}

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function xmlObject(name) {
		ensureCompat();
		return window[name] || document[name] || document.getElementById(name);
	}

	function xmlDocument(name) {
		var object = xmlObject(name);
		return object && object.XMLDocument || object && object._doc || object && object.nodeType === 9 && object || null;
	}

	function xmlRoot(name) {
		var object = xmlObject(name);
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || object && object._doc && object._doc.documentElement || null;
	}

	function serializeXml(name) {
		var doc = xmlDocument(name);
		var root = xmlRoot(name);
		if (doc) {
			return new XMLSerializer().serializeToString(doc);
		}
		return root ? new XMLSerializer().serializeToString(root) : "";
	}

	function loadXml(name, text) {
		var object = xmlObject(name);
		var xmlText = trim(text) || "<Root/>";
		var doc = new DOMParser().parseFromString(xmlText, "application/xml");
		if (object && typeof object.loadXML === "function") {
			object.loadXML(xmlText);
		} else if (object) {
			object._doc = doc;
		}
		return object || doc;
	}

	function childElements(node, nodeName) {
		var wanted = nodeName ? String(nodeName).toLowerCase() : "";
		return Array.prototype.slice.call(node && node.childNodes || []).filter(function (child) {
			return child.nodeType === 1 && (!wanted || String(child.nodeName).toLowerCase() === wanted);
		});
	}

	function attr(node, nameOrIndex) {
		var item;
		if (!node || !node.attributes) {
			return "";
		}
		if (typeof nameOrIndex === "number") {
			item = node.attributes.item(nameOrIndex);
			return item ? item.nodeValue : "";
		}
		return node.getAttribute(nameOrIndex) || "";
	}

	function attrAny(node, names, index) {
		var value = "";
		names.some(function (name) {
			value = attr(node, name);
			return trim(value) !== "";
		});
		return trim(value) !== "" ? value : attr(node, index);
	}

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function xpathLiteral(value) {
		var text = String(value == null ? "" : value);
		if (text.indexOf("'") === -1) {
			return "'" + text + "'";
		}
		if (text.indexOf('"') === -1) {
			return '"' + text + '"';
		}
		return "concat('" + text.replace(/'/g, "',\"'\",'") + "')";
	}

	function selectNodes(context, expression) {
		var doc;
		var found;
		var nodes = [];
		if (!context) {
			return nodes;
		}
		if (typeof context.selectNodes === "function") {
			return Array.prototype.slice.call(context.selectNodes(expression));
		}
		doc = context.nodeType === 9 ? context : context.ownerDocument;
		if (!doc || !doc.evaluate) {
			return nodes;
		}
		try {
			found = doc.evaluate(expression, context, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
			for (var index = 0; index < found.snapshotLength; index += 1) {
				nodes.push(found.snapshotItem(index));
			}
		} catch (ignore) {}
		return nodes;
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
		try {
			xhr.setRequestHeader("Content-Type", "text/xml");
		} catch (ignore) {}
		xhr.send(body || "");
		return xhr;
	}

	function table(id) {
		return document.getElementById(id);
	}

	function insertCell(row, mode, name, value, className, align, valign, size, maxLength, colspan) {
		if (typeof window.InsertCell === "function") {
			return window.InsertCell(row, mode, name || "", value == null ? "" : String(value), className || "ExcelDisplayCell", align || "", valign || "", size || 0, maxLength || 0, colspan || 0, 0, "");
		}
		var cell = row.insertCell(-1);
		cell.className = className || "ExcelDisplayCell";
		if (align) {
			cell.align = align;
		}
		if (colspan) {
			cell.colSpan = colspan;
		}
		if (mode === 2) {
			var input = document.createElement("input");
			input.type = "text";
			input.name = name || "";
			input.value = value == null ? "" : String(value);
			input.className = "Formelem";
			if (size) {
				input.size = size;
			}
			if (maxLength) {
				input.maxLength = maxLength;
			}
			cell.appendChild(input);
		} else {
			cell.innerHTML = value == null ? "" : String(value);
		}
		return cell;
	}

	function clearTable() {
		var target = table("tblBook");
		var row;
		if (!target) {
			return false;
		}
		while (target.rows.length) {
			target.deleteRow(0);
		}
		row = target.insertRow(0);
		row.removeAttribute("data-agent-header");
		insertCell(row, 1, "", "S.No", "ExcelSerial", "Center");
		insertCell(row, 1, "", "Period", "ExcelHeaderCell", "left");
		insertCell(row, 1, "", "StartNo", "ExcelHeaderCell", "left");
		insertCell(row, 1, "", "Prefix", "ExcelHeaderCell", "left");
		insertCell(row, 1, "", "Suffix", "ExcelHeaderCell", "left");
		return false;
	}

	function addAgentHeader() {
		var target = table("tblBook");
		var row = target && target.rows[0];
		var cell;
		if (!row || row.getAttribute("data-agent-header") === "1") {
			return;
		}
		cell = row.insertCell(-1);
		cell.className = "ExcelHeaderCell";
		cell.innerHTML = "Agent Name ";
		row.setAttribute("data-agent-header", "1");
	}

	function clearSeriesTable() {
		var target = table("tblVoucher");
		var display = document.getElementById("DisVoucher");
		var row;
		if (display) {
			display.style.height = "1px";
			display.style.visibility = "hidden";
		}
		if (!target) {
			return false;
		}
		while (target.rows.length) {
			target.deleteRow(0);
		}
		row = target.insertRow(0);
		insertCell(row, 1, "", "S.No", "ExcelSerial", "Center");
		insertCell(row, 1, "", "&nbsp;", "ExcelHeaderCell", "left");
		insertCell(row, 1, "", "Number For", "ExcelHeaderCell", "left");
		insertCell(row, 1, "", "Item Type", "ExcelHeaderCell", "left");
		insertCell(row, 1, "", "Invoice Type", "ExcelHeaderCell", "left");
		insertCell(row, 1, "", "Sale Type", "ExcelHeaderCell", "left");
		insertCell(row, 1, "", "Agent", "ExcelHeaderCell", "left");
		insertCell(row, 1, "", "Series Type", "ExcelHeaderCell", "left");
		insertCell(row, 1, "", "Agent Code", "ExcelHeaderCell", "left");
		insertCell(row, 1, "", "Item Values", "ExcelHeaderCell", "left");
		insertCell(row, 1, "", "Invoice Values", "ExcelHeaderCell", "left");
		insertCell(row, 1, "", "Sale Values", "ExcelHeaderCell", "left");
		return false;
	}

	function periodLabel(periodType, period) {
		if (periodType === "M") {
			return "Month-" + period;
		}
		if (periodType === "Q") {
			return "Quater-" + period;
		}
		if (periodType === "Y") {
			return "Yearly";
		}
		return period;
	}

	function loadNumberSeries(url, targetName) {
		var xhr = syncGet(url);
		if (xhr.responseXML && xhr.responseXML.documentElement) {
			loadXml(targetName, new XMLSerializer().serializeToString(xhr.responseXML));
		} else if (trim(xhr.responseText)) {
			loadXml(targetName, xhr.responseText);
		}
		return xhr;
	}

	function seriesNode(seriesNo) {
		return selectNodes(xmlRoot("SeriesNoData"), "//Series[@No=" + xpathLiteral(seriesNo) + "]")[0];
	}

	function templateEntries(seriesNo) {
		return selectNodes(xmlRoot("SeriesNoData"), "//Series[@No=" + xpathLiteral(seriesNo) + "]/Entry");
	}

	function setTemplateValues(template) {
		if (template) {
			setValue("hSeriesType", attr(template, 2));
			setValue("hSeriesLen", attr(template, 4));
		}
	}

	function commissionAgentEnabled() {
		return radioChecked("optComm", 0);
	}

	function appendExistingSeriesRow(entry, rowIndex, readOnly) {
		var target = table("tblBook");
		var row = target.insertRow(rowIndex);
		var entryNo = attr(entry, 9) || rowIndex;
		insertCell(row, 1, "", attr(entry, 0) || rowIndex, "ExcelSerial", "Center");
		insertCell(row, 1, "", attr(entry, 1), "ExcelDisplayCell", "left");
		if (readOnly) {
			insertCell(row, 1, "", attr(entry, 2), "ExcelDisplayCell", "", "", 7, 6);
			insertCell(row, 1, "", attr(entry, 3), "ExcelDisplayCell", "", "", 11, 10);
			insertCell(row, 1, "", attr(entry, 4), "ExcelDisplayCell", "", "", 11, 10);
		} else {
			insertCell(row, 2, "txtStartNo" + entryNo, attr(entry, 2), "ExcelInputCell", "", "", 8, 8);
			insertCell(row, 2, "txtPrefix" + entryNo, attr(entry, 3), "ExcelInputCell", "", "", 11, 10);
			insertCell(row, 2, "txtSuffix" + entryNo, attr(entry, 4), "ExcelInputCell", "", "", 11, 10);
		}
		if (commissionAgentEnabled()) {
			insertCell(row, 1, "", field("hTempName") && field("hTempName").value || "", "ExcelDisplayCell", "", "", 11, 10);
		}
	}

	function appendTemplateRow(entry, template, rowIndex, agentCode, agentName) {
		var target = table("tblBook");
		var row = target.insertRow(rowIndex);
		var entryNo = attr(entry, 0);
		var numberLength = Number(attr(template, 4)) || 0;
		var prefix = agentCode ? "txt" + agentCode : "txt";
		insertCell(row, 1, "", rowIndex, "ExcelSerial", "Center");
		insertCell(row, 1, "", periodLabel(attr(template, 3), attr(entry, 1)), "ExcelDisplayCell", "left");
		insertCell(row, 2, prefix + "StartNo" + entryNo, attr(entry, 2), "ExcelInputCell", "", "", numberLength + 2 || 8, numberLength || 8);
		insertCell(row, 2, prefix + "Prefix" + entryNo, attr(entry, 3), "ExcelInputCell", "", "", 11, 10);
		insertCell(row, 2, prefix + "Suffix" + entryNo, attr(entry, 4), "ExcelInputCell", "", "", 11, 10);
		if (agentCode) {
			insertCell(row, 2, "txt" + agentCode + entryNo, agentName, "ExcelInputCell", "", "", 11, 10);
		}
	}

	function displayTable() {
		var select = field("selNoSeries");
		var seriesNo = selectedValue(select);
		var template;
		var entries;
		var unit;
		var existingRows;
		var usedRows;
		var listNode;
		var rowIndex = 1;
		clearTable();
		if (!select || select.selectedIndex === 0) {
			return false;
		}
		if (commissionAgentEnabled()) {
			addAgentHeader();
		}
		template = seriesNode(seriesNo);
		entries = templateEntries(seriesNo);
		setTemplateValues(template);
		if (entries.length) {
			setValue("hTotEntNo", entries.length);
		}
		unit = selectedValue(field("selActType")) === "QUT" ? "010101" : selectedValue(field("selUnit"));
		loadNumberSeries("XMLGetNoSeries.asp?sVal=" + encodeURIComponent([field("hSeriesNo") && field("hSeriesNo").value || "", field("hSeriesCode") && field("hSeriesCode").value || "", unit].join(":")), "NoSeries");
		existingRows = selectNodes(xmlRoot("NoSeries"), "//NumSeries");
		if (existingRows.length) {
			if (trim(field("hTransNo") && field("hTransNo").value) !== "") {
				listNode = selectNodes(xmlRoot("SeriesList"), "//NumSeriesList[@TransNo=" + xpathLiteral(field("hTransNo").value) + "]")[0];
				setValue("hDispCheck", listNode ? attr(listNode, "NoUsed") : field("hDispCheck").value);
			}
			existingRows.forEach(function (entry) {
				var readOnly = field("hDispCheck") && field("hDispCheck").value === "Y";
				appendExistingSeriesRow(entry, rowIndex, readOnly);
				setValue("hEntryNo", attr(existingRows[existingRows.length - 1], 9));
				setValue("hFinFrom", attr(entry, 5));
				setValue("hFinTo", attr(entry, 6));
				setValue("hSeriesNo", attr(entry, 7));
				setValue("hSeriesCode", attr(entry, 8));
				rowIndex += 1;
			});
			usedRows = selectNodes(xmlRoot("NoSeries"), "//NumSeries[@StartNo != '1' and @StartNo != 1]");
			setDisabled("btnSubmit", usedRows.length > 0);
			return false;
		}
		setValue("hFinFrom", "");
		setValue("hFinTo", "");
		setValue("hSeriesNo", "");
		setValue("hSeriesCode", "");
		setValue("hEntryNo", "0");
		if (template && trim(field("hAgentCode") && field("hAgentCode").value) !== "") {
			String(field("hAgentCode").value).split(":").forEach(function (code, agentIndex) {
				var names = String(field("hAgentName") && field("hAgentName").value || "").split(":");
				entries.forEach(function (entry) {
					appendTemplateRow(entry, template, rowIndex, code, names[agentIndex] || "");
					rowIndex += 1;
				});
			});
		} else if (template) {
			entries.forEach(function (entry) {
				appendTemplateRow(entry, template, rowIndex, "", "");
				rowIndex += 1;
			});
		}
		return false;
	}

	function init() {
		if (field("selUnit")) {
			field("selUnit").selectedIndex = 0;
		}
		return false;
	}

	function popSeriesNo() {
		var select = field("selNoSeries");
		var root = xmlRoot("SeriesNoData");
		if (!select) {
			return false;
		}
		select.options.length = 0;
		select.add(new Option("Select Number Series", "0"));
		childElements(root).forEach(function (header) {
			if (attr(header, 3) === "M") {
				select.add(new Option(attr(header, 1), attr(header, 0)));
			}
		});
		return false;
	}

	function activityChange(obj) {
		formReset();
		if (obj && obj.value === "QUT") {
			setDisabled("selUnit", true);
			getSeriesList("U");
		} else {
			setDisabled("selUnit", false);
		}
		return false;
	}

	function optclick(obj) {
		setValue("hOptval", obj && obj.value || "");
		if (field("selItmType") && field("selAgent")) {
			if (obj && obj.value === "I") {
				field("selItmType").disabled = false;
				field("selAgent").selectedIndex = 0;
				field("selAgent").disabled = true;
			} else if (obj && obj.value === "A") {
				field("selItmType").selectedIndex = 0;
				field("selItmType").disabled = true;
				field("selAgent").disabled = false;
			} else {
				field("selItmType").selectedIndex = 0;
				field("selAgent").selectedIndex = 0;
				field("selItmType").disabled = true;
				field("selAgent").disabled = true;
			}
		}
		return false;
	}

	function numForName() {
		if (radioChecked("chkFor", 0) && radioChecked("chkFor", 1)) {
			return "Both";
		}
		if (radioChecked("chkFor", 0)) {
			return "Domestic";
		}
		return "Export";
	}

	function numForCode() {
		return numForName().charAt(0);
	}

	function itemTypeName() {
		return "All";
	}

	function invTypeName() {
		return radioChecked("optInv", 0) ? "All" : "Specific";
	}

	function saleTypeName() {
		return radioChecked("optSale", 0) ? "All" : "Specific";
	}

	function agentTypeName() {
		return commissionAgentEnabled() ? "Yes" : "No";
	}

	function createElement(name) {
		var doc = xmlDocument("SeriesList");
		return doc ? doc.createElement(name) : document.implementation.createDocument("", "", null).createElement(name);
	}

	function setSeriesAttributes(node, editCheck) {
		setAttr(node, "NumFor", numForName());
		setAttr(node, "ItemTy", itemTypeName());
		setAttr(node, "InvTy", invTypeName());
		setAttr(node, "SaleTy", saleTypeName());
		setAttr(node, "AgentTy", agentTypeName());
		setAttr(node, "ItemValue", field("hItemValue") && field("hItemValue").value || "");
		setAttr(node, "InvValue", field("hInvValue") && field("hInvValue").value || "");
		setAttr(node, "SaleValue", field("hSaleValue") && field("hSaleValue").value || "");
		setAttr(node, "AgentCode", field("hAgentCode") && field("hAgentCode").value || "");
		setAttr(node, "AgentName", field("hAgentName") && field("hAgentName").value || "");
		setAttr(node, "ItemDesc", "");
		setAttr(node, "InvDesc", field("txtInv") && field("txtInv").value || "");
		setAttr(node, "SaleDesc", field("txtSale") && field("txtSale").value || "");
		setAttr(node, "EditCheck", editCheck);
	}

	function createXML() {
		var root = xmlRoot("SeriesList");
		var node = createElement("NumSeriesList");
		if (!root) {
			return false;
		}
		setAttr(node, "TransNo", "0");
		setAttr(node, "SeriesNo", "0");
		setAttr(node, "SeriesCode", "0");
		setAttr(node, "EntryNo", "1");
		setSeriesAttributes(node, "E");
		root.appendChild(node);
		return true;
	}

	function rowSuffixes() {
		var suffixes = [];
		Array.prototype.slice.call(document.querySelectorAll('input[name^="txtStartNo"]')).forEach(function (input) {
			var suffix = input.name.replace(/^txtStartNo/, "");
			if (suffix) {
				suffixes.push(suffix);
			}
		});
		if (!suffixes.length) {
			for (var index = 1, count = Number(field("hTotEntNo") && field("hTotEntNo").value) || 0; index <= count; index += 1) {
				suffixes.push(String(index));
			}
		}
		return suffixes;
	}

	function updateXML() {
		var node = selectNodes(xmlRoot("SeriesList"), "//NumSeriesList[@TransNo=" + xpathLiteral(field("hTransNo") && field("hTransNo").value || "") + "]")[0];
		if (!node) {
			return false;
		}
		setSeriesAttributes(node, "Y");
		childElements(node, "NoList").forEach(function (child) {
			node.removeChild(child);
		});
		rowSuffixes().forEach(function (suffix) {
			var noList = createElement("NoList");
			setAttr(noList, "StartNo", field("txtStartNo" + suffix) ? field("txtStartNo" + suffix).value : "");
			setAttr(noList, "Prefix", field("txtPrefix" + suffix) ? field("txtPrefix" + suffix).value : "");
			setAttr(noList, "Suffix", field("txtSuffix" + suffix) ? field("txtSuffix" + suffix).value : "");
			node.appendChild(noList);
		});
		return true;
	}

	function normalizeSpecific(value) {
		var text = trim(value);
		if (text === "All" || text === "A") {
			return "A";
		}
		if (text === "Specific" || text === "S") {
			return "S";
		}
		return text;
	}

	function checkMatch(fullValue, specificValue) {
		var full = String(fullValue || "").split(":").map(trim).filter(Boolean);
		var specific = String(specificValue || "").split(":").map(trim).filter(Boolean);
		if (!full.length || !specific.length) {
			return true;
		}
		return !specific.some(function (value) {
			return full.indexOf(value) !== -1;
		});
	}

	function typeOverlap(currentType, oldType, currentValue, oldValue) {
		var old = normalizeSpecific(oldType);
		if (currentType === "A" && old === "A") {
			return true;
		}
		if (currentType === "S" && old === "S") {
			return !checkMatch(currentValue, oldValue);
		}
		if (currentType === "A" && old === "S") {
			return true;
		}
		if (currentType === "S" && old === "A") {
			return true;
		}
		return false;
	}

	function numForOverlap(currentCode, oldValue) {
		var oldCode = String(oldValue || "").charAt(0);
		return (currentCode === "B" && oldCode === "B") ||
			(currentCode === "D" && oldCode === "B") ||
			(currentCode === "E" && oldCode === "B") ||
			(currentCode === "B" && oldCode === "D") ||
			(currentCode === "B" && oldCode === "E");
	}

	function splitCodes(value) {
		return String(value || "").split(":").map(trim).filter(Boolean);
	}

	function agentOverlap(currentType, oldType, oldAgentCode) {
		var currentCodes;
		var oldCode;
		if (currentType === "Y" && oldType === "Yes") {
			currentCodes = splitCodes(field("hAgentCode") && field("hAgentCode").value);
			oldCode = trim(oldAgentCode);
			if (oldCode.length === 1) {
				oldCode = "0" + oldCode;
			}
			return currentCodes.indexOf(oldCode) !== -1 || currentCodes.indexOf(trim(oldAgentCode)) !== -1;
		}
		return currentType === "N" && oldType === "No";
	}

	function rowCombinationMatches(node, includeClass) {
		var currentNumFor = numForCode();
		var currentItemTy = "A";
		var currentInvTy = radioChecked("optInv", 0) ? "A" : "S";
		var currentSaleTy = radioChecked("optSale", 0) ? "A" : "S";
		var currentAgentTy = commissionAgentEnabled() ? "Y" : "N";
		var classMatches = true;
		if (includeClass) {
			classMatches = trim(field("hClassCode") && field("hClassCode").value) === trim(attr(node, "SelClass"));
		}
		return numForOverlap(currentNumFor, attrAny(node, ["NumFor"], 1)) &&
			typeOverlap(currentItemTy, attrAny(node, ["ItemTy"], 5), field("hItemValue") && field("hItemValue").value || "", attrAny(node, ["ItemValue"], 9)) &&
			typeOverlap(currentInvTy, attrAny(node, ["InvTy"], 6), field("hInvValue") && field("hInvValue").value || "", attrAny(node, ["InvValue"], 10)) &&
			typeOverlap(currentSaleTy, attrAny(node, ["SaleTy"], 7), field("hSaleValue") && field("hSaleValue").value || "", attrAny(node, ["SaleValue"], 11)) &&
			agentOverlap(currentAgentTy, attrAny(node, ["AgentTy"], 8), attrAny(node, ["AgentCode"], 12)) &&
			classMatches;
	}

	function hasDuplicate(skipTransNo, includeClass) {
		var duplicate = false;
		selectNodes(xmlRoot("SeriesList"), "//NumSeriesList").some(function (node) {
			if (skipTransNo && attrAny(node, ["TransNo"], 0) === skipTransNo) {
				return false;
			}
			duplicate = rowCombinationMatches(node, includeClass);
			return duplicate;
		});
		return duplicate;
	}

	function checkDuplicate() {
		return !hasDuplicate("", true);
	}

	function amendDuplicate() {
		return !hasDuplicate(field("hTransNo") && field("hTransNo").value || "", false);
	}

	function formReset() {
		if (field("selUnit")) {
			field("selUnit").selectedIndex = 0;
		}
		if (field("selNoSeries")) {
			field("selNoSeries").disabled = false;
			field("selNoSeries").selectedIndex = 0;
		}
		setRadio("chkFor", 0, true);
		setRadio("chkFor", 1, true);
		setRadio("optInv", 0, true);
		setRadio("optInv", 1, false);
		setRadio("optSale", 0, true);
		setRadio("optSale", 1, false);
		setRadio("optComm", 0, false);
		setRadio("optComm", 1, true);
		setValue("txtInv", "");
		setValue("txtSale", "");
		setValue("hAgentCode", "");
		setValue("hAgentName", "");
		setValue("hItemType", "");
		setValue("hInvType", "");
		setValue("hSaleType", "");
		setValue("hItemValue", "");
		setValue("hInvValue", "");
		setValue("hSaleValue", "");
		setValue("hTransNo", "");
		setValue("hSeriesNo", "0");
		setValue("hSeriesCode", "0");
		setValue("hEditCheck", "N");
		setDisabled("btnUpdate", true);
		setDisabled("btnSubmit", false);
		return false;
	}

	function postSeriesAndSubmit() {
		var xhr = syncPost("XMLSave.asp?Name=NoSeries&Mod=SA", serializeXml("SeriesList"));
		if (trim(xhr.responseText) !== "") {
			alert(xhr.responseText);
			return false;
		}
		form().submit();
		return true;
	}

	function validateForm(callType) {
		var mustValidate;
		if (callType === "U" && field("hDispCheck") && field("hDispCheck").value === "Y") {
			clearSeriesTable();
			getSeriesList("A");
			formReset();
			clearTable();
			return false;
		}
		mustValidate = (field("hEditCheck") && field("hEditCheck").value === "N" && callType === "S") || callType === "U";
		if (mustValidate) {
			if (selectedValue(field("selActType")) !== "QUT" && field("selUnit") && field("selUnit").selectedIndex === 0) {
				alert("Select Unit");
				focusField("selUnit");
				return false;
			}
			if (field("selActType") && field("selActType").selectedIndex === 0) {
				alert("Select Activity");
				focusField("selActType");
				return false;
			}
			if (!radioChecked("chkFor", 0) && !radioChecked("chkFor", 1)) {
				alert("Select Number Series For ");
				focusField("chkFor");
				return false;
			}
			if (radioChecked("optInv", 1) && trim(field("hInvValue") && field("hInvValue").value) === "") {
				alert("Select any Invoice Type ");
				focusField("optInv");
				return false;
			}
			if (radioChecked("optSale", 1) && trim(field("hSaleValue") && field("hSaleValue").value) === "") {
				alert("Select any Sale Type ");
				focusField("optSale");
				return false;
			}
			if (field("selNoSeries") && field("selNoSeries").selectedIndex === 0) {
				alert("Select Number Series");
				focusField("selNoSeries");
				return false;
			}
			setValue("hActivityName", selectedText(field("selActType")));
			setValue("hNumFor", numForCode());
		}
		if (callType === "U") {
			if (!amendDuplicate()) {
				alert("Combination already gets Matches, Choose Different one ");
				return false;
			}
			updateXML();
			clearSeriesTable();
			getSeriesList("A");
			clearTable();
			return postSeriesAndSubmit();
		}
		if (callType === "S" && field("hEditCheck") && field("hEditCheck").value === "N") {
			if (!checkDuplicate()) {
				alert("Selected Combination gets Matches with Existing one. Select a different Combination ");
				return false;
			}
			if (field("selNoSeries") && field("selNoSeries").selectedIndex !== 0) {
				createXML();
				return postSeriesAndSubmit();
			}
		}
		return postSeriesAndSubmit();
	}

	function popupDet(obj, type) {
		var selected = obj && obj.value;
		var value;
		var url;
		if (selected === "S" || selected === "Y") {
			if (type === "I") {
				setValue("hItemType", selected);
			} else if (type === "V") {
				setValue("hInvType", selected);
			} else if (type === "S") {
				setValue("hSaleType", selected);
			}
			value = type !== "A" ? selectedValue(field("selUnit")) + ":" + type + ":0" : selectedValue(field("selUnit")) + ":" + type + ":" + (field("hItemValue") && field("hItemValue").value || "");
			url = "NoSeriesItemSel.asp?Value=" + encodeURIComponent(value);
			if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
				window.ITMSModernCompat.openModalDialog(url, "A", "dialogHeight:200px;dialogWidth:350px;center:Yes;help:No;resizable:No;status:No", function (result) {
					var parts;
					var names;
					if (trim(result) === "") {
						return;
					}
					parts = String(result || "").split("``");
					names = String(parts[1] || "").replace(/:/g, ",");
					if (type === "I") {
						setValue("hItemValue", parts[0] || "");
					} else if (type === "V") {
						setValue("txtInv", names);
						setValue("hInvValue", parts[0] || "");
					} else if (type === "S") {
						setValue("txtSale", names);
						setValue("hSaleValue", parts[0] || "");
					} else if (type === "A") {
						setValue("hAgentCode", parts[0] || "");
						popAgent(result);
					}
				});
			}
		} else if (type === "I") {
			setValue("hItemValue", "");
		} else if (type === "V") {
			setValue("txtInv", "");
			setValue("hInvValue", "");
		} else if (type === "S") {
			setValue("txtSale", "");
			setValue("hSaleValue", "");
		}
		return false;
	}

	function popAgent(value) {
		var parts = String(value || "").split("``");
		setValue("hAgentCode", parts[0] || "");
		setValue("hAgentName", parts[2] || "");
		addAgentHeader();
		if (field("selNoSeries") && field("selNoSeries").selectedIndex !== 0) {
			displayTable();
		}
		return false;
	}

	function getSeriesList(callType) {
		var unit;
		var activity = selectedValue(field("selActType"));
		var display = document.getElementById("DisVoucher");
		clearSeriesTable();
		unit = activity === "QUT" ? "010101" : selectedValue(field("selUnit"));
		if (callType === "U") {
			loadNumberSeries("GetNumberSeriesList.asp?sVal=" + encodeURIComponent(unit + ":" + activity), "SeriesList");
		}
		childElements(xmlRoot("SeriesList"), "NumSeriesList").forEach(function (node, index) {
			var seriesNo = attrAny(node, ["SeriesNo"], 2);
			var template = seriesNode(seriesNo);
			var row = table("tblVoucher").insertRow(-1);
			if (display) {
				display.style.height = "200px";
				display.style.visibility = "visible";
			}
			insertCell(row, 1, "", index + 1, "ExcelSerial", "Center", "top");
			insertCell(row, 1, "", '<a class="ExcelDisplaylink" href="#" onclick="EditEntry(\'' + attrAny(node, ["TransNo"], 0) + '\'); return false;">Edit</a>', "ExcelDisplayCell", "Center", "top");
			insertCell(row, 1, "", attrAny(node, ["NumFor"], 1), "ExcelDisplayCell", "Center", "top");
			insertCell(row, 1, "", attrAny(node, ["ItemTy"], 5), "ExcelDisplayCell", "Center", "top");
			insertCell(row, 1, "", attrAny(node, ["InvTy"], 6), "ExcelDisplayCell", "Center", "top");
			insertCell(row, 1, "", attrAny(node, ["SaleTy"], 7), "ExcelDisplayCell", "Center", "top");
			insertCell(row, 1, "", attrAny(node, ["AgentTy"], 8), "ExcelDisplayCell", "Center", "top");
			insertCell(row, 1, "", template ? attr(template, "Description") || attr(template, 1) : "", "ExcelDisplayCell", "Center", "top");
			insertCell(row, 1, "", attrAny(node, ["AgentName"], 13), "ExcelDisplayCell", "Center", "top");
			insertCell(row, 1, "", attrAny(node, ["ItemDesc"], 14), "ExcelDisplayCell", "Center", "top");
			insertCell(row, 1, "", attrAny(node, ["InvDesc"], 15), "ExcelDisplayCell", "Center", "top");
			insertCell(row, 1, "", attrAny(node, ["SaleDesc"], 16), "ExcelDisplayCell", "Center", "top");
		});
		return false;
	}

	function setCategorySelection(categoryCode) {
		var category = field("selCategory");
		if (!category) {
			return;
		}
		for (var index = 0; index < category.options.length; index += 1) {
			if (trim(category.options[index].value) === trim(categoryCode)) {
				category.selectedIndex = index;
				category.disabled = true;
				break;
			}
		}
	}

	function editEntry(transNo) {
		var node = selectNodes(xmlRoot("SeriesList"), "//NumSeriesList[@TransNo=" + xpathLiteral(transNo) + "]")[0];
		var select = field("selNoSeries");
		var seriesNo;
		var className;
		if (!node) {
			return false;
		}
		setValue("hEditCheck", "Y");
		setDisabled("btnUpdate", false);
		if (select) {
			select.disabled = true;
		}
		setValue("hTransNo", transNo);
		seriesNo = attrAny(node, ["SeriesNo"], 2);
		setValue("hSeriesNo", seriesNo);
		setValue("hSeriesCode", attrAny(node, ["SeriesCode"], 3));
		setValue("hTempName", attrAny(node, ["AgentName"], 13));
		if (select) {
			for (var index = 0; index < select.options.length; index += 1) {
				if (select.options[index].value === seriesNo) {
					select.selectedIndex = index;
					break;
				}
			}
		}
		if (attrAny(node, ["NumFor"], 1) === "Both") {
			setRadio("chkFor", 0, true);
			setRadio("chkFor", 1, true);
		} else if (attrAny(node, ["NumFor"], 1) === "Domestic") {
			setRadio("chkFor", 0, true);
			setRadio("chkFor", 1, false);
		} else {
			setRadio("chkFor", 0, false);
			setRadio("chkFor", 1, true);
		}
		if (attrAny(node, ["InvTy"], 6) === "Specific") {
			setRadio("optInv", 0, false);
			setRadio("optInv", 1, true);
			setValue("hInvValue", attrAny(node, ["InvValue"], 10));
		} else {
			setRadio("optInv", 0, true);
			setRadio("optInv", 1, false);
		}
		if (attrAny(node, ["SaleTy"], 7) === "Specific") {
			setRadio("optSale", 0, false);
			setRadio("optSale", 1, true);
			setValue("hSaleValue", attrAny(node, ["SaleValue"], 11));
		} else {
			setRadio("optSale", 0, true);
			setRadio("optSale", 1, false);
		}
		if (attrAny(node, ["AgentTy"], 8) === "Yes") {
			setRadio("optComm", 0, true);
			setRadio("optComm", 1, false);
		} else {
			setRadio("optComm", 0, false);
			setRadio("optComm", 1, true);
		}
		setValue("txtInv", attrAny(node, ["InvDesc"], 15));
		setValue("txtSale", attrAny(node, ["SaleDesc"], 16));
		setValue("hAgentCode", attrAny(node, ["AgentCode"], 12));
		setValue("hAgentName", attrAny(node, ["AgentName"], 13));
		setValue("hClassCode", attr(node, "SelClass"));
		setValue("hCatCode", attr(node, "SelCat"));
		setCategorySelection(field("hCatCode") && field("hCatCode").value);
		className = attr(node, "SelClassName") || attr(node, "SelCatName");
		setText("txtClass", className);
		displayTable();
		setDisabled("btnSubmit", true);
		return false;
	}

	function parseClassificationReturn(returnData) {
		var parts = String(returnData || "").split("*****");
		var classCodes = [];
		var categoryCodes = [];
		var classNames = [];
		if (parts[0] === "-1" || trim(parts[0]) === "") {
			return null;
		}
		String(parts[0] || "").split("|").forEach(function (value) {
			var split = value.split(":");
			if (split.length > 1) {
				classCodes.push(split[split.length - 1]);
				categoryCodes.push(split[1]);
			} else if (split[0]) {
				categoryCodes.push(split[0].substring(3));
			}
		});
		String(parts[1] || "").split("|||").forEach(function (value) {
			var split = value.split(":");
			if (trim(split[split.length - 1]) !== "") {
				classNames.push(split[split.length - 1]);
			}
		});
		return {
			classCode: classCodes.join(","),
			categoryCode: categoryCodes.join(","),
			className: classNames.join(",")
		};
	}

	function applyClassification(returnData) {
		var result = parseClassificationReturn(returnData);
		if (!result) {
			return;
		}
		setText("txtClass", result.className);
		setValue("hClassCode", result.classCode);
		setValue("hCatCode", result.categoryCode);
		setCategorySelection(result.categoryCode);
	}

	function selectClassifcation() {
		var orgId = selectedValue(field("selUnit"));
		var url = "/include/ClassificationSelectPop.asp?sIType=1&sOrgID=" + encodeURIComponent(orgId) + "&sITypename=&SelMode=M";
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, "Classification", "dialogHeight:500px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No", applyClassification);
		}
		return false;
	}

	window.Activity_change = activityChange;
	window.Optclick = optclick;
	window.DisplayTable = displayTable;
	window.init = init;
	window.popSeriesNo = popSeriesNo;
	window.ClearTable = clearTable;
	window.ClearSeriesTable = clearSeriesTable;
	window.validateForm = validateForm;
	window.FormReset = formReset;
	window.UpdateXML = updateXML;
	window.UpDateXML = updateXML;
	window.CreateXML = createXML;
	window.CheckDuplicate = checkDuplicate;
	window.AmendDuplicate = amendDuplicate;
	window.PopupDet = popupDet;
	window.PopAgent = popAgent;
	window.GetSeriesList = getSeriesList;
	window.EditEntry = editEntry;
	window.CheckMatch = checkMatch;
	window.SelectClassifcation = selectClassifcation;
	window.MatchStr = function (text, pattern) {
		return new RegExp(pattern, "i").test(text);
	};
}(window, document));
