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

	function setText(id, value) {
		var item = document.getElementById(id) || field(id);
		if (item) {
			item.textContent = value == null ? "" : String(value);
		}
	}

	function setDisabled(name, disabled) {
		var item = field(name);
		if (item) {
			item.disabled = !!disabled;
		}
	}

	function setChecked(name, index, checked) {
		var list = fields(name);
		if (list[index]) {
			list[index].checked = !!checked;
		}
	}

	function isChecked(name, index) {
		var list = fields(name);
		return !!(list[index] && list[index].checked);
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
		var doc = new DOMParser().parseFromString(trim(text) || "<Root/>", "application/xml");
		if (object && typeof object.loadXML === "function") {
			object.loadXML(trim(text) || "<Root/>");
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

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
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
			input.className = className || "ExcelInputCell";
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
		insertCell(row, 1, "", "S.No", "ExcelSerial", "Center");
		insertCell(row, 1, "", "Period", "ExcelHeaderCell", "left");
		insertCell(row, 1, "", "StartNo", "ExcelHeaderCell", "left");
		insertCell(row, 1, "", "Prefix", "ExcelHeaderCell", "left");
		insertCell(row, 1, "", "Suffix", "ExcelHeaderCell", "left");
		return false;
	}

	function clearSeriesTable() {
		var target = table("tblVoucher");
		var row;
		if (!target) {
			return false;
		}
		while (target.rows.length) {
			target.deleteRow(0);
		}
		row = target.insertRow(0);
		insertCell(row, 1, "", "S.No", "ExcelSerial", "Center");
		insertCell(row, 1, "", "&nbsp;", "ExcelHeaderCell", "left");
		insertCell(row, 1, "", "Manual<br>Numbering", "ExcelHeaderCell", "center");
		insertCell(row, 1, "", "Item Type", "ExcelHeaderCell", "center");
		insertCell(row, 1, "", "Product Wise", "ExcelHeaderCell", "center");
		insertCell(row, 1, "", "Packing Type", "ExcelHeaderCell", "center");
		insertCell(row, 1, "", "Series Type", "ExcelHeaderCell", "center");
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

	function appendSeriesRow(entry, rowIndex, readOnly) {
		var target = table("tblBook");
		var row = target.insertRow(rowIndex);
		var entryNo = attr(entry, 9) || attr(entry, 0) || rowIndex;
		insertCell(row, 1, "", attr(entry, 0) || rowIndex, "ExcelSerial", "Center");
		insertCell(row, 1, "", attr(entry, 1), "ExcelDisplayCell", "left");
		if (readOnly) {
			insertCell(row, 1, "", attr(entry, 2), "ExcelDisplayCell");
			insertCell(row, 1, "", attr(entry, 3), "ExcelDisplayCell");
			insertCell(row, 1, "", attr(entry, 4), "ExcelDisplayCell");
		} else {
			insertCell(row, 2, "txtStartNo" + entryNo, attr(entry, 2), "ExcelInputCell", "", "", 8, 8);
			insertCell(row, 2, "txtPrefix" + entryNo, attr(entry, 3), "ExcelInputCell", "", "", 11, 10);
			insertCell(row, 2, "txtSuffix" + entryNo, attr(entry, 4), "ExcelInputCell", "", "", 11, 10);
		}
	}

	function appendTemplateRow(entry, header, rowIndex) {
		var target = table("tblBook");
		var row = target.insertRow(rowIndex);
		var entryNo = attr(entry, 0);
		insertCell(row, 1, "", rowIndex, "ExcelSerial", "Center");
		insertCell(row, 1, "", periodLabel(attr(header, 3), attr(entry, 1)), "ExcelDisplayCell", "left");
		insertCell(row, 2, "txtStartNo" + entryNo, attr(entry, 2), "ExcelInputCell", "", "", 8, 8);
		insertCell(row, 2, "txtPrefix" + entryNo, attr(entry, 3), "ExcelInputCell", "", "", 11, 10);
		insertCell(row, 2, "txtSuffix" + entryNo, attr(entry, 4), "ExcelInputCell", "", "", 11, 10);
	}

	function seriesByNo(seriesNo) {
		var root = xmlRoot("SeriesNoData");
		return selectNodes(root, "//Series[@No=" + seriesNo + "]")[0];
	}

	function entriesForSeries(seriesNo) {
		var root = xmlRoot("SeriesNoData");
		return selectNodes(root, "//Series[@No=" + seriesNo + "]/Entry");
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

	function loadNumberSeries(url, targetName) {
		var xhr = syncGet(url);
		if (xhr.responseXML && xhr.responseXML.documentElement) {
			loadXml(targetName, new XMLSerializer().serializeToString(xhr.responseXML));
		} else if (trim(xhr.responseText)) {
			loadXml(targetName, xhr.responseText);
		}
		return xhr;
	}

	function displayTable() {
		var select = field("selNoSeries");
		var seriesCode = select && select.value;
		var root = xmlRoot("SeriesNoData");
		var template = seriesByNo(seriesCode);
		var existingRows;
		var rowIndex = 1;
		var unit = selectedValue(field("selUnit"));
		var storedKey = [field("hSeriesNo") && field("hSeriesNo").value, field("hSeriesCode") && field("hSeriesCode").value, unit].join(":");
		clearTable();
		if (!select || select.selectedIndex === 0) {
			return false;
		}
		if (template) {
			setValue("hSeriesType", attr(template, 2));
			setValue("hSeriesLen", attr(template, 4));
			setValue("hTotEntNo", childElements(template, "Entry").length);
		}
		if (unit && unit !== "select") {
			loadNumberSeries("XMLGetNoSeriesPR.asp?sVal=" + encodeURIComponent(storedKey), "NoSeries");
			existingRows = selectNodes(xmlRoot("NoSeries"), "//NumSeries");
			if (existingRows.length) {
				if (trim(field("hTransNo") && field("hTransNo").value) !== "") {
					var parts = String(field("hTransNo").value).split("Z");
					var listNode = selectNodes(xmlRoot("SeriesList"), "//NumSeriesList[@SeriesNo=" + parts[0] + " and @SeriesCode=" + parts[1] + "]")[0];
					setValue("hDispCheck", listNode ? attr(listNode, "NoUsed") : field("hDispCheck").value);
				}
				existingRows.forEach(function (entry) {
					var readOnly = field("hDispCheck") && field("hDispCheck").value === "Y";
					appendSeriesRow(entry, rowIndex, readOnly);
					setValue("hEntryNo", attr(existingRows[existingRows.length - 1], 9));
					setValue("hFinFrom", attr(entry, 5));
					setValue("hFinTo", attr(entry, 6));
					setValue("hSeriesNo", attr(entry, 7));
					setValue("hSeriesCode", attr(entry, 8));
					rowIndex += 1;
				});
				setDisabled("btnSubmit", !!selectNodes(xmlRoot("NoSeries"), "//NumSeries[@StartNo !=1]").length);
				return false;
			}
			setValue("hFinFrom", "");
			setValue("hFinTo", "");
			setValue("hSeriesNo", "");
			setValue("hSeriesCode", "");
			setValue("hEntryNo", "0");
		}
		if (template) {
			childElements(template, "Entry").forEach(function (entry) {
				appendTemplateRow(entry, template, rowIndex);
				rowIndex += 1;
			});
		}
		return false;
	}

	function formReset() {
		if (field("selNoSeries")) {
			field("selNoSeries").disabled = false;
			field("selNoSeries").selectedIndex = 0;
		}
		setValue("hTransNo", "");
		setValue("hSeriesNo", "0");
		setValue("hSeriesCode", "0");
		setDisabled("btnUpdate", true);
		setDisabled("btnSubmit", false);
		return false;
	}

	function disableNoSeries() {
		if (selectedValue(field("selUnit")) === "select") {
			alert("Select the unit from the list");
			if (field("selUnit")) {
				field("selUnit").focus();
			}
			setChecked("radManual", 1, true);
			return false;
		}
		if (selectedValue(field("selNumType")) === "select") {
			alert("Select the numbering type from the list");
			if (field("selNumType")) {
				field("selNumType").focus();
			}
			setChecked("radManual", 1, true);
			return false;
		}
		if (field("selNoSeries")) {
			field("selNoSeries").selectedIndex = 0;
			field("selNoSeries").disabled = true;
		}
		clearTable();
		return false;
	}

	function enableNoSeries() {
		if (isChecked("radManual", 1) && field("selNoSeries")) {
			field("selNoSeries").disabled = false;
		}
		return false;
	}

	function disableChk(obj) {
		if (field("chkPacking")) {
			field("chkPacking").checked = false;
			field("chkPacking").disabled = obj && obj.value !== "P";
		}
		if (field("chkProductWise")) {
			field("chkProductWise").checked = false;
			field("chkProductWise").disabled = obj && obj.value !== "C";
		}
		formReset();
		clearTable();
		getSeriesList("U");
		return false;
	}

	function createXmlElement(name) {
		var doc = xmlDocument("SeriesList");
		return doc ? doc.createElement(name) : document.implementation.createDocument("", "", null).createElement(name);
	}

	function createXML() {
		var root = xmlRoot("SeriesList");
		var node = createXmlElement("NumSeriesList");
		if (!root) {
			return;
		}
		setAttr(node, "TransNo", "0");
		setAttr(node, "SeriesNo", "0");
		setAttr(node, "SeriesCode", "0");
		setAttr(node, "EntryNo", "1");
		setAttr(node, "EditCheck", "E");
		root.appendChild(node);
	}

	function updateXML() {
		var key = String(field("hTransNo") && field("hTransNo").value || "");
		var parts = key.split("Z");
		var node = selectNodes(xmlRoot("SeriesList"), "//NumSeriesList[@SeriesNo=" + parts[0] + " and @SeriesCode=" + parts[1] + "]")[0];
		var count = Number(field("hTotEntNo") && field("hTotEntNo").value) || 0;
		if (!node) {
			return false;
		}
		setAttr(node, "EditCheck", "Y");
		for (var index = 1; index <= count; index += 1) {
			var noList = createXmlElement("NoList");
			setAttr(noList, "StartNo", field("txtStartNo" + index) ? field("txtStartNo" + index).value : "");
			setAttr(noList, "Prefix", field("txtPrefix" + index) ? field("txtPrefix" + index).value : "");
			setAttr(noList, "Suffix", field("txtSuffix" + index) ? field("txtSuffix" + index).value : "");
			node.appendChild(noList);
		}
		return true;
	}

	function checkDuplicate() {
		var rows = selectNodes(xmlRoot("SeriesList"), "//NumSeriesList");
		var type = selectedValue(field("selNumType"));
		if (rows.length > 0 && (type === "U" || type === "P")) {
			return false;
		}
		return true;
	}

	function amendDuplicate() {
		return true;
	}

	function postSeriesAndSubmit() {
		var xhr = syncPost("XMLSave.asp?Name=NoSeries&Mod=PR", serializeXml("SeriesList"));
		if (trim(xhr.responseText) !== "") {
			alert(xhr.responseText);
			return false;
		}
		form().submit();
		return true;
	}

	function validateForm(callType) {
		if (callType === "U" && field("hDispCheck") && field("hDispCheck").value === "Y") {
			clearSeriesTable();
			getSeriesList("A");
			formReset();
			clearTable();
			return false;
		}
		if (field("selUnit") && field("selUnit").selectedIndex === 0) {
			alert("Select Unit");
			field("selUnit").focus();
			return false;
		}
		if (field("selNumType") && field("selNumType").selectedIndex === 0) {
			alert("Select Numbering Type");
			field("selNumType").focus();
			return false;
		}
		if (isChecked("radManual", 1) && field("selNoSeries") && field("selNoSeries").selectedIndex === 0) {
			alert("Select Number Series");
			field("selNoSeries").focus();
			return false;
		}
		setValue("hActivityName", selectedText(field("selNumType")));
		if (callType === "U") {
			if (!amendDuplicate()) {
				alert("Combination already gets Matches, Choose Different one ");
				return false;
			}
			updateXML();
			return postSeriesAndSubmit();
		}
		if (callType === "S" && field("hEditCheck") && field("hEditCheck").value === "N") {
			if (!checkDuplicate()) {
				alert("Selected Combination gets Matches with Existing one. Select a different Combination ");
				return false;
			}
			if (field("selNoSeries") && field("selNoSeries").selectedIndex !== 0) {
				createXML();
			}
		}
		return postSeriesAndSubmit();
	}

	function getSeriesList(callType) {
		var unit = selectedValue(field("selUnit"));
		var type = selectedValue(field("selNumType"));
		var existingType = "";
		var root;
		var seriesTemplateRoot;
		clearSeriesTable();
		if (callType === "U") {
			loadNumberSeries("GetNumberSeriesListPR.asp?sVal=" + encodeURIComponent(unit + ":"), "SeriesList");
			childElements(xmlRoot("SeriesList"), "NumSeriesList").forEach(function (node) {
				existingType = trim(attr(node, "NumberingType"));
			});
			if (existingType && existingType !== type) {
				alert("No. Series is already defined");
				return false;
			}
			loadNumberSeries("GetNumberSeriesListPR.asp?sVal=" + encodeURIComponent(unit + ":" + type), "SeriesList");
		}
		root = xmlRoot("SeriesList");
		seriesTemplateRoot = xmlRoot("SeriesNoData");
		childElements(root, "NumSeriesList").forEach(function (node, index) {
			var display = document.getElementById("DisVoucher");
			var seriesNo = attr(node, "SeriesNo") || attr(node, 0);
			var seriesCode = attr(node, "SeriesCode") || attr(node, 1);
			var template = selectNodes(seriesTemplateRoot, "//Series[@No=" + seriesNo + "]")[0];
			var row = table("tblVoucher").insertRow(-1);
			if (display) {
				display.style.height = "200px";
				display.style.visibility = "visible";
			}
			insertCell(row, 1, "", index + 1, "ExcelSerial", "Center", "top");
			insertCell(row, 1, "", '<a class="ExcelDisplaylink" href="#" onclick="EditEntry(\'' + seriesNo + "Z" + seriesCode + '\'); return false;">Edit</a>', "ExcelDisplayCell", "Center", "top");
			insertCell(row, 1, "", (attr(node, "ManualNumbering") || attr(node, 2)) === "N" ? "No" : "Yes", "ExcelDisplayCell", "Center", "top");
			insertCell(row, 1, "", trim(attr(node, "ItemType") || attr(node, 3)) === "0" ? "N/A" : trim(attr(node, 10)), "ExcelDisplayCell", "Center", "top");
			insertCell(row, 1, "", trim(attr(node, "ProductWise") || attr(node, 4)) === "0" ? "N/A" : "Applicable", "ExcelDisplayCell", "Center", "top");
			insertCell(row, 1, "", trim(attr(node, "PackingType") || attr(node, 5)) === "0" ? "N/A" : "Applicable", "ExcelDisplayCell", "Center", "top");
			insertCell(row, 1, "", template ? attr(template, "Description") || attr(template, 1) : "", "ExcelDisplayCell", "Center", "top");
		});
		return false;
	}

	function editEntry(value) {
		var parts = String(value || "").split("Z");
		var node = selectNodes(xmlRoot("SeriesList"), "//NumSeriesList[@SeriesNo=" + parts[0] + " and @SeriesCode=" + parts[1] + "]")[0];
		var select = field("selNoSeries");
		if (!node) {
			return false;
		}
		setValue("hEditCheck", "Y");
		setDisabled("btnUpdate", false);
		if (select) {
			select.disabled = true;
			for (var index = 0; index < select.options.length; index += 1) {
				if (select.options[index].value === parts[0]) {
					select.selectedIndex = index;
					break;
				}
			}
		}
		setValue("hTransNo", value);
		setValue("hSeriesNo", parts[0]);
		setValue("hSeriesCode", parts[1]);
		if ((attr(node, "ManualNumbering") || attr(node, 2)) === "Y") {
			setChecked("radManual", 0, true);
			setChecked("radManual", 1, false);
		} else {
			setChecked("radManual", 0, false);
			setChecked("radManual", 1, true);
		}
		fields("radManual").forEach(function (radio) {
			radio.disabled = true;
		});
		if (field("chkProductWise")) {
			field("chkProductWise").checked = (attr(node, "ProductWise") || attr(node, 4)) !== "0";
		}
		if (field("chkPacking")) {
			field("chkPacking").checked = (attr(node, "PackingType") || attr(node, 5)) !== "0";
		}
		setValue("hClassCode", attr(node, "ClassCode"));
		displayTable();
		setDisabled("btnSubmit", true);
		return false;
	}

	function popupDet(select, type) {
		var selected = select && select.value;
		var value;
		var url;
		if (selected !== "S" && selected !== "Y") {
			return false;
		}
		value = selectedValue(field("selUnit")) + ":" + type + ":" + (type === "A" ? field("hItemValue") && field("hItemValue").value || "" : "0");
		url = "NoSeriesItemSel.asp?Value=" + encodeURIComponent(value);
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, "A", "dialogHeight:200px;dialogWidth:350px;center:Yes;help:No;resizable:No;status:No", function (result) {
				var parts = String(result || "").split("``");
				var names = String(parts[1] || "").replace(/:/g, ",");
				if (type === "I") {
					setValue("txtItem", names);
					setValue("hItemValue", parts[0]);
				} else if (type === "V") {
					setValue("txtInv", names);
					setValue("hInvValue", parts[0]);
				} else if (type === "S") {
					setValue("txtSale", names);
					setValue("hSaleValue", parts[0]);
				}
			});
		}
		return false;
	}

	window.DisableNoSeries = disableNoSeries;
	window.EnableNoSeries = enableNoSeries;
	window.DisableChk = disableChk;
	window.DisplayTable = displayTable;
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
	window.GetSeriesList = getSeriesList;
	window.EditEntry = editEntry;
	window.CheckMatch = function (full, part) {
		return String(full || "").split(":").indexOf(String(part || "")) === -1;
	};
	window.MatchStr = function (text, pattern) {
		return new RegExp(pattern, "i").test(text);
	};
}(window, document));
