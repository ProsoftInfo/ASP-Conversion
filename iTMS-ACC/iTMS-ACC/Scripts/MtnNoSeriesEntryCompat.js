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

	function getText(id) {
		var item = document.getElementById(id) || field(id);
		return item ? trim(item.textContent || "") : "";
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

	function insertCellReadOnly(row, mode, name, value, className, align, valign, size, maxLength, colspan, rowspan, options) {
		return insertCell(row, mode, name, value, className, align, valign, size, maxLength, colspan, rowspan, options);
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
		insertCell(row, 1, "", "Work Group", "ExcelHeaderCell", "left");
		insertCell(row, 1, "", "Work Group Desc", "ExcelHeaderCell", "left");
		insertCell(row, 1, "", "Series Type", "ExcelHeaderCell", "left");
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

	function appendExistingSeriesRow(entry, rowIndex, readOnly, numberLength) {
		var target = table("tblBook");
		var row = target.insertRow(rowIndex);
		var entryNo = attr(entry, 9) || rowIndex;
		var inputSize = (Number(numberLength) || 0) + 5;
		insertCell(row, 1, "", attr(entry, 0) || rowIndex, "ExcelSerial", "Center");
		insertCell(row, 1, "", attr(entry, 1), "ExcelDisplayCell", "left");
		if (readOnly) {
			insertCellReadOnly(row, 1, "", attr(entry, 2), "ExcelDisplayCell", "", "", 7, 6, 0, 0, "");
			insertCellReadOnly(row, 1, "", attr(entry, 3), "ExcelDisplayCell", "", "", 11, 10, 0, 0, "");
			insertCellReadOnly(row, 1, "", attr(entry, 4), "ExcelDisplayCell", "", "", 11, 10, 0, 0, "");
		} else {
			insertCell(row, 2, "txtStartNo" + entryNo, attr(entry, 2), "ExcelInputCell", "", "", inputSize, inputSize);
			insertCell(row, 2, "txtPrefix" + entryNo, attr(entry, 3), "ExcelInputCell", "", "", 11, 10);
			insertCell(row, 2, "txtSuffix" + entryNo, attr(entry, 4), "ExcelInputCell", "", "", 11, 10);
		}
	}

	function appendTemplateRow(entry, template, rowIndex, numberLength) {
		var target = table("tblBook");
		var row = target.insertRow(rowIndex);
		var entryNo = attr(entry, 0);
		var inputSize = (Number(numberLength) || 0) + 2;
		insertCell(row, 1, "", rowIndex, "ExcelSerial", "Center");
		insertCell(row, 1, "", periodLabel(attr(template, 3), attr(entry, 1)), "ExcelDisplayCell", "left");
		insertCell(row, 2, "txtStartNo" + entryNo, attr(entry, 2), "ExcelInputCell", "", "", inputSize, Number(numberLength) || 0);
		insertCell(row, 2, "txtPrefix" + entryNo, attr(entry, 3), "ExcelInputCell", "", "", 11, 10);
		insertCell(row, 2, "txtSuffix" + entryNo, attr(entry, 4), "ExcelInputCell", "", "", 11, 10);
	}

	function displayTable() {
		var select = field("selNoSeries");
		var seriesNo = selectedValue(select);
		var template;
		var entries;
		var numberLength;
		var unit;
		var existingRows;
		var usedRows;
		var listNode;
		var rowIndex = 1;
		clearTable();
		if (!select || select.selectedIndex === 0) {
			return false;
		}
		template = seriesNode(seriesNo);
		entries = templateEntries(seriesNo);
		numberLength = template ? attr(template, 4) : "";
		setTemplateValues(template);
		if (entries.length) {
			setValue("hTotEntNo", entries.length);
		}
		unit = selectedValue(field("selActType")) === "QUT" ? "010101" : selectedValue(field("selUnit"));
		loadNumberSeries("XMLGetNoSeries_Mtn.asp?sVal=" + encodeURIComponent([field("hSeriesNo") && field("hSeriesNo").value || "", field("hSeriesCode") && field("hSeriesCode").value || "", unit].join(":")), "NoSeries");
		existingRows = selectNodes(xmlRoot("NoSeries"), "//NumSeries");
		if (existingRows.length) {
			if (trim(field("hTransNo") && field("hTransNo").value) !== "") {
				listNode = selectNodes(xmlRoot("SeriesList"), "//NumSeriesList[@TransNo=" + xpathLiteral(field("hTransNo").value) + "]")[0];
				setValue("hDispCheck", listNode ? attr(listNode, "NoUsed") : field("hDispCheck").value);
			}
			existingRows.forEach(function (entry) {
				var readOnly = field("hDispCheck") && field("hDispCheck").value === "Y";
				appendExistingSeriesRow(entry, rowIndex, readOnly, numberLength);
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
		if (template) {
			entries.forEach(function (entry) {
				appendTemplateRow(entry, template, rowIndex, numberLength);
				rowIndex += 1;
			});
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

	function workGroupTypeName() {
		return radioChecked("OptWorkGroup", 0) ? "All" : "Specific";
	}

	function workGroupTypeCode() {
		return radioChecked("OptWorkGroup", 0) ? "A" : "S";
	}

	function createElement(name) {
		var doc = xmlDocument("SeriesList");
		return doc ? doc.createElement(name) : document.implementation.createDocument("", "", null).createElement(name);
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
		setAttr(node, "WorkGroupType", workGroupTypeName());
		setAttr(node, "WorkGroupValue", field("hWorkGroupValue") && field("hWorkGroupValue").value || "");
		setAttr(node, "WorkGroupDesc", getText("SpnWorkGroup"));
		setAttr(node, "EditCheck", "E");
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
		setAttr(node, "WorkGroupType", workGroupTypeName());
		setAttr(node, "WorkGroupValue", field("hWorkGroupValue") && field("hWorkGroupValue").value || "");
		setAttr(node, "WorkGroupDesc", getText("SpnWorkGroup"));
		setAttr(node, "EditCheck", "Y");
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

	function workGroupValue(node) {
		return attrAny(node, ["WorkGroupValue"], 6);
	}

	function workGroupType(node) {
		return attrAny(node, ["WorkGroupType"], 5);
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

	function hasDuplicate(skipTransNo) {
		var selectedType = workGroupTypeCode();
		var currentValue = field("hWorkGroupValue") && field("hWorkGroupValue").value || "";
		var duplicate = false;
		selectNodes(xmlRoot("SeriesList"), "//NumSeriesList").some(function (node) {
			var oldType;
			if (skipTransNo && attrAny(node, ["TransNo"], 0) === skipTransNo) {
				return false;
			}
			oldType = workGroupType(node);
			if (selectedType === "A" && oldType === "All") {
				duplicate = true;
			} else if (selectedType === "S" && oldType === "Specific") {
				duplicate = !checkMatch(currentValue, workGroupValue(node));
			} else if (selectedType === "A" && oldType === "Specific") {
				duplicate = true;
			} else if (selectedType === "S" && oldType === "All") {
				duplicate = true;
			}
			return duplicate;
		});
		return duplicate;
	}

	function checkDuplicate() {
		return !hasDuplicate("");
	}

	function amendDuplicate() {
		return !hasDuplicate(field("hTransNo") && field("hTransNo").value || "");
	}

	function resetSeries() {
		if (field("selNoSeries")) {
			field("selNoSeries").disabled = false;
		}
		return false;
	}

	function formReset() {
		if (field("selNoSeries")) {
			field("selNoSeries").disabled = false;
			field("selNoSeries").selectedIndex = 0;
		}
		setRadio("OptWorkGroup", 0, true);
		setRadio("OptWorkGroup", 1, false);
		setText("SpnWorkGroup", "");
		setValue("hWorkGroupValue", "");
		setValue("hTransNo", "");
		setValue("hSeriesNo", "0");
		setValue("hSeriesCode", "0");
		setValue("hEditCheck", "N");
		setDisabled("btnUpdate", true);
		setDisabled("btnSubmit", false);
		return false;
	}

	function setSeriesNo(transNo) {
		var node = selectNodes(xmlRoot("SeriesList"), "//NumSeriesList[@TransNo=" + xpathLiteral(transNo) + "]")[0];
		var select = field("selNoSeries");
		var seriesNo;
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
		if (select) {
			for (var index = 0; index < select.options.length; index += 1) {
				if (select.options[index].value === seriesNo) {
					select.selectedIndex = index;
					break;
				}
			}
		}
		if (workGroupType(node) === "Specific") {
			setRadio("OptWorkGroup", 0, false);
			setRadio("OptWorkGroup", 1, true);
			setValue("hWorkGroupValue", workGroupValue(node));
		} else {
			setRadio("OptWorkGroup", 0, true);
			setRadio("OptWorkGroup", 1, false);
			setValue("hWorkGroupValue", "");
		}
		setText("SpnWorkGroup", attrAny(node, ["WorkGroupDesc"], 7));
		displayTable();
		setDisabled("btnSubmit", true);
		return false;
	}

	function getSeriesList(callType) {
		var unit = selectedValue(field("selUnit"));
		var activity = selectedValue(field("selActType"));
		var display = document.getElementById("DisVoucher");
		clearSeriesTable();
		clearTable();
		if (callType === "U") {
			loadNumberSeries("GetNumberSeriesList_Mtn.asp?sVal=" + encodeURIComponent(unit + ":" + activity), "SeriesList");
		}
		childElements(xmlRoot("SeriesList"), "NumSeriesList").forEach(function (node, index) {
			var seriesNo = attrAny(node, ["SeriesNo"], 2);
			var template = seriesNode(seriesNo);
			var workValue = workGroupValue(node);
			var row = table("tblVoucher").insertRow(-1);
			if (display) {
				display.style.height = "200px";
				display.style.visibility = "visible";
			}
			if (workValue === "0") {
				setAttr(node, "WorkGroupValue", "All");
			}
			insertCell(row, 1, "", index + 1, "ExcelSerial", "Center", "top");
			insertCell(row, 1, "", '<a class="ExcelDisplaylink" href="#" onclick="SetSeriesNo(\'' + attrAny(node, ["TransNo"], 0) + '\'); return false;">Edit</a>', "ExcelDisplayCell", "Center", "top");
			insertCell(row, 1, "", workGroupType(node), "ExcelDisplayCell", "Center", "top");
			insertCell(row, 1, "", attrAny(node, ["WorkGroupDesc"], 7), "ExcelDisplayCell", "Center", "top");
			insertCell(row, 1, "", template ? attr(template, "Description") || attr(template, 1) : "", "ExcelDisplayCell", "Center", "top");
		});
		return false;
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
			if (field("selActType") && field("selActType").selectedIndex === 0) {
				alert("Select Activity");
				focusField("selActType");
				return false;
			}
			if (field("selUnit") && field("selUnit").selectedIndex === 0) {
				alert("Select Unit");
				focusField("selUnit");
				return false;
			}
			if (!radioChecked("OptWorkGroup", 0) && !radioChecked("OptWorkGroup", 1)) {
				alert("Select Number Series For ");
				focusField("OptWorkGroup");
				return false;
			}
			if (radioChecked("OptWorkGroup", 1) && trim(field("hWorkGroupValue") && field("hWorkGroupValue").value) === "") {
				alert("Select any Work Group");
				focusField("OptWorkGroup");
				return false;
			}
			if (field("selNoSeries") && field("selNoSeries").selectedIndex === 0) {
				alert("Select Number Series");
				focusField("selNoSeries");
				return false;
			}
			setValue("hActivityName", selectedText(field("selActType")));
		}
		if (callType === "U") {
			if (!amendDuplicate()) {
				alert("Combination already gets Matches, Choose Different one ");
				return false;
			}
			updateXML();
			return postSeriesAndSubmit(true);
		}
		if (callType === "S" && field("hEditCheck") && field("hEditCheck").value === "N") {
			if (!checkDuplicate()) {
				alert("Selected Combination gets Matches with Existing one. Select a different Combination ");
				return false;
			}
			if (field("selNoSeries") && field("selNoSeries").selectedIndex !== 0) {
				createXML();
				return postSeriesAndSubmit(true);
			}
		}
		return false;
	}

	function postSeriesAndSubmit(resetOnError) {
		var xhr = syncPost("XMLSave.asp?Name=NoSeries&Mod=MTN", serializeXml("SeriesList"));
		if (trim(xhr.responseText) !== "") {
			alert(xhr.responseText);
			if (resetOnError) {
				clearSeriesTable();
				getSeriesList("A");
				formReset();
				clearTable();
			}
			return false;
		}
		form().submit();
		return true;
	}

	function popupDet(obj) {
		var selected = obj && obj.value;
		var url;
		resetSeries();
		if (selected === "S") {
			url = "NoSeriesWorkGroupSelect.asp?Value=" + encodeURIComponent(selectedValue(field("selUnit")) + ":" + (field("hWorkGroupValue") && field("hWorkGroupValue").value || ""));
			if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
				window.ITMSModernCompat.openModalDialog(url, "A", "dialogHeight:200px;dialogWidth:350px;center:Yes;help:No;resizable:No;status:No", function (result) {
					var parts;
					if (trim(result) === "") {
						return;
					}
					parts = String(result || "").split("``");
					setValue("hWorkGroupValue", parts[0] || "");
					setText("SpnWorkGroup", String(parts[1] || "").replace(/:/g, ","));
				});
			}
		} else {
			setText("SpnWorkGroup", "");
			setValue("hWorkGroupValue", "");
		}
		return false;
	}

	window.InsertCellReadOnly = insertCellReadOnly;
	window.SetSeriesNo = setSeriesNo;
	window.GetSeriesList = getSeriesList;
	window.ClearSeriesTable = clearSeriesTable;
	window.DisplayTable = displayTable;
	window.popSeriesNo = popSeriesNo;
	window.ClearTable = clearTable;
	window.validateForm = validateForm;
	window.ResetSeries = resetSeries;
	window.FormReset = formReset;
	window.UpdateXML = updateXML;
	window.UpDateXML = updateXML;
	window.CreateXML = createXML;
	window.CheckDuplicate = checkDuplicate;
	window.AmendDuplicate = amendDuplicate;
	window.PopupDet = popupDet;
	window.CheckMatch = checkMatch;
	window.MatchStr = function (text, pattern) {
		return new RegExp(pattern, "i").test(text);
	};
}(window, document));
