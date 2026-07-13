(function (window, document) {
	"use strict";

	var outDataValue = null;
	var optionDialogHasData = false;

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		var item;
		var wanted;
		var i;
		if (!frm || !frm.elements) {
			return null;
		}
		item = frm.elements[name];
		if (item) {
			return item;
		}
		wanted = String(name).toLowerCase();
		for (i = 0; i < frm.elements.length; i += 1) {
			if (String(frm.elements[i].name || "").toLowerCase() === wanted) {
				return frm.elements[i];
			}
		}
		return null;
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

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function xmlObject(name) {
		ensureCompat();
		return window[name] || document[name] || document.getElementById(name) || null;
	}

	function xmlRoot(nameOrObject) {
		var object = typeof nameOrObject === "string" ? xmlObject(nameOrObject) : nameOrObject;
		if (window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.xmlRoot) {
			return window.ITMSModalReturnCompat.xmlRoot(object);
		}
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || object && object._doc && object._doc.documentElement || object && object.nodeType === 1 && object || null;
	}

	function xmlDocument(name) {
		var object = xmlObject(name);
		var root = xmlRoot(object);
		return object && object.XMLDocument || object && object._doc || root && root.ownerDocument || null;
	}

	function childElements(node, name) {
		var result = [];
		var wanted = name ? String(name).toUpperCase() : "";
		var child;
		for (var i = 0; node && i < node.childNodes.length; i += 1) {
			child = node.childNodes[i];
			if (child.nodeType === 1 && (!wanted || String(child.nodeName || "").toUpperCase() === wanted)) {
				result.push(child);
			}
		}
		return result;
	}

	function attr(node, name) {
		return node && node.getAttribute ? node.getAttribute(name) || "" : "";
	}

	function clearChildren(node) {
		while (node && node.firstChild) {
			node.removeChild(node.firstChild);
		}
	}

	function importNode(targetRoot, sourceNode) {
		var doc = targetRoot && targetRoot.ownerDocument;
		if (doc && doc.importNode) {
			return doc.importNode(sourceNode, true);
		}
		return sourceNode.cloneNode(true);
	}

	function serializeXml(rootOrDoc) {
		var doc = rootOrDoc && rootOrDoc.nodeType === 9 ? rootOrDoc : rootOrDoc && rootOrDoc.ownerDocument;
		return new XMLSerializer().serializeToString(doc || rootOrDoc);
	}

	function loadXml(name, text) {
		var object = xmlObject(name);
		if (object && typeof object.loadXML === "function") {
			object.loadXML(text);
			return xmlRoot(object);
		}
		return null;
	}

	function postXml(url, root) {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", url, false);
		xhr.setRequestHeader("Content-Type", "text/xml");
		xhr.send(serializeXml(root));
		return xhr.responseText || "";
	}

	function getText(url) {
		var xhr = new XMLHttpRequest();
		xhr.open("GET", url, false);
		xhr.send(null);
		if (xhr.responseXML && xhr.responseXML.documentElement) {
			loadXml("EditData", serializeXml(xhr.responseXML));
		} else if (trim(xhr.responseText)) {
			loadXml("EditData", xhr.responseText);
		} else {
			alert(xhr.responseText || "No records returned");
		}
	}

	function openDialog(url, args, features, callback) {
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			return window.ITMSModernCompat.openModalDialog(url, args, features, callback);
		}
		window.open(url, "_blank");
		return null;
	}

	function selectedText(select) {
		return select && select.selectedIndex >= 0 && select.options[select.selectedIndex] ? select.options[select.selectedIndex].text : "";
	}

	function selectedValue(name) {
		var item = field(name);
		return item ? item.value : "";
	}

	function setSelectByValue(select, value) {
		var wanted = trim(value);
		if (!select) {
			return;
		}
		for (var i = 0; i < select.options.length; i += 1) {
			if (trim(select.options[i].value) === wanted) {
				select.selectedIndex = i;
				return;
			}
		}
	}

	function checkNumbers(value) {
		return /^[0-9.]+$/.test(String(value || ""));
	}

	function configureDataTypeControls(value) {
		var dataType = value || selectedValue("selDataType");
		var dataLength = field("txtDataLen");
		var decimals = field("txtDecimals");
		if (dataLength) {
			dataLength.disabled = false;
		}
		if (decimals) {
			decimals.disabled = false;
		}
		if (dataType === "String" && decimals) {
			decimals.disabled = true;
		}
		if (dataType === "Options") {
			if (dataLength) {
				dataLength.disabled = true;
			}
			if (decimals) {
				decimals.disabled = true;
			}
		}
	}

	function resetDataLengthFields() {
		setValue("txtDataLen", "");
		setValue("txtDecimals", "");
	}

	function validateOptionContext() {
		if (selectedValue("selHeader") === "select") {
			alert("Select Header");
			field("selHeader").focus();
			field("selDataType").selectedIndex = 0;
			return false;
		}
		if (!trim(valueOf("txtAttribute"))) {
			alert("Enter Attribute Name");
			field("txtAttribute").select();
			field("selDataType").selectedIndex = 0;
			return false;
		}
		return true;
	}

	function validateAttributeForm() {
		var dataType = selectedValue("selDataType");
		if (selectedValue("selHeader") === "select") {
			alert("Select Header");
			field("selHeader").focus();
			return false;
		}
		if (!trim(valueOf("txtAttribute"))) {
			alert("Enter Attribute Name");
			field("txtAttribute").select();
			return false;
		}
		if (dataType === "select") {
			alert("Select Data Type");
			field("selDataType").focus();
			return false;
		}
		if ((dataType === "String" || dataType === "Numeric") && !trim(valueOf("txtDataLen"))) {
			alert("Enter Data Length");
			field("txtDataLen").select();
			return false;
		}
		if ((dataType === "String" || dataType === "Numeric") && trim(valueOf("txtDataLen")) && !checkNumbers(valueOf("txtDataLen"))) {
			alert("Enter Numerals Only");
			field("txtDataLen").select();
			return false;
		}
		if (dataType === "Numeric" && !trim(valueOf("txtDecimals"))) {
			alert("Enter Decimal Length");
			field("txtDecimals").select();
			return false;
		}
		if (dataType === "Numeric" && trim(valueOf("txtDecimals")) && !checkNumbers(valueOf("txtDecimals"))) {
			alert("Enter Numerals Only");
			field("txtDecimals").select();
			return false;
		}
		return true;
	}

	function setAttributePayload(root, includeOldId) {
		root.setAttribute("ITYPE", trim(valueOf("hItemType")));
		root.setAttribute("HEADER", trim(selectedValue("selHeader")));
		root.setAttribute("ATTRNAME", trim(valueOf("txtAttribute")));
		root.setAttribute("DATATYPE", trim(selectedValue("selDataType")));
		root.setAttribute("DATALENGTH", trim(valueOf("txtDataLen")));
		root.setAttribute("DECIMAL", trim(valueOf("txtDecimals")));
		if (includeOldId) {
			root.setAttribute("OLDID", trim(valueOf("hOldAttID")));
		}
		root.setAttribute("CLASSCODE", trim(valueOf("hClassCode")));
	}

	function responseCode(text) {
		return String(text || "").replace(/\s+/g, "").toUpperCase();
	}

	function addTextCell(row, text, className, align) {
		var cell = row.insertCell(-1);
		cell.className = className || "ExcelDisplayCell";
		cell.align = align || "left";
		cell.textContent = text == null ? "" : String(text);
		return cell;
	}

	function addHiddenCell(row, name, value) {
		var cell = row.insertCell(-1);
		var input = document.createElement("input");
		cell.style.display = "none";
		input.type = "hidden";
		input.name = name;
		input.value = value == null ? "" : String(value);
		cell.appendChild(input);
		return input;
	}

	function addCheckbox(container, name, value) {
		var checkbox = document.createElement("input");
		checkbox.type = "checkbox";
		checkbox.name = name;
		checkbox.value = value == null ? "" : String(value);
		checkbox.className = "FormElem";
		container.appendChild(checkbox);
		return checkbox;
	}

	function clearTable(tableId) {
		var table = document.getElementById(tableId);
		if (!table) {
			return;
		}
		while (table.rows.length > 1) {
			table.deleteRow(1);
		}
	}

	function ClearTable() {
		clearTable("tblDisplay");
		clearTable("tblAttSelect");
	}

	function addNoRecordsRow(tableId, colspan) {
		var table = document.getElementById(tableId);
		var row;
		var cell;
		if (!table) {
			return;
		}
		row = table.insertRow(-1);
		cell = addTextCell(row, "No Records Found", "ExcelDisplayCell", "center");
		cell.colSpan = colspan;
	}

	function addManageRow(index, node) {
		var table = document.getElementById("tblDisplay");
		var row;
		var cell;
		var button;
		if (!table) {
			return;
		}
		row = table.insertRow(-1);
		addTextCell(row, index, "ExcelHeaderCell", "center").width = "10";
		cell = addTextCell(row, "", "ExcelDisplayCell", "left");
		button = document.createElement("input");
		button.type = "button";
		button.name = "btn" + index;
		button.value = "Edit";
		button.className = "ActionButtonX";
		button.onclick = function () {
			EditAttribute(index);
		};
		cell.appendChild(button);
		cell = addTextCell(row, "", "ExcelDisplayCell", "center");
		addCheckbox(cell, "Chk" + index, index);
		addTextCell(row, attr(node, "HName"), "ExcelDisplayCell", "left");
		addTextCell(row, attr(node, "Name"), "ExcelDisplayCell", "left");
		addHiddenCell(row, "hItemTypeZ" + index, attr(node, "ItemType"));
		addHiddenCell(row, "hHeaderZ" + index, attr(node, "Header"));
		addHiddenCell(row, "hAttNameZ" + index, attr(node, "Name"));
		addHiddenCell(row, "hTypeZ" + index, trim(attr(node, "Type")));
		addHiddenCell(row, "hLengthZ" + index, attr(node, "Length"));
		addHiddenCell(row, "hDecimalZ" + index, attr(node, "Decimal"));
		addHiddenCell(row, "hAttIDZ" + index, attr(node, "ID"));
	}

	function addSelectRow(index, node) {
		var table = document.getElementById("tblAttSelect");
		var row;
		var cell;
		if (!table) {
			return;
		}
		row = table.insertRow(-1);
		addTextCell(row, index, "ExcelHeaderCell", "center").width = "10";
		cell = addTextCell(row, "", "ExcelDisplayCell", "center");
		addCheckbox(cell, "ChkSelect" + index, index);
		addTextCell(row, attr(node, "HName"), "ExcelDisplayCell", "left");
		addTextCell(row, attr(node, "Name"), "ExcelDisplayCell", "left");
		addHiddenCell(row, "hItemTypeSZ" + index, attr(node, "ItemType"));
		addHiddenCell(row, "hHeaderSZ" + index, attr(node, "Header"));
		addHiddenCell(row, "hAttNameSZ" + index, attr(node, "Name"));
		addHiddenCell(row, "hTypeSZ" + index, trim(attr(node, "Type")));
		addHiddenCell(row, "hLengthSZ" + index, attr(node, "Length"));
		addHiddenCell(row, "hDecimalSZ" + index, attr(node, "Decimal"));
		addHiddenCell(row, "hAttIDSZ" + index, attr(node, "ID"));
		addHiddenCell(row, "hAttValSZ" + index, "");
	}

	function copyExistingAttributeToNewData(editRoot) {
		var targetRoot = xmlRoot("NewData");
		var oldId = trim(valueOf("hOldAttID"));
		if (!targetRoot || !oldId) {
			return;
		}
		childElements(editRoot, "Attribute").some(function (node) {
			if (trim(attr(node, "ID")) !== oldId) {
				return false;
			}
			clearChildren(targetRoot);
			targetRoot.appendChild(importNode(targetRoot, node));
			return true;
		});
	}

	function populateHeader() {
		var root = xmlRoot("ItemTypeHeader");
		var select = field("selHeader");
		var itemType = trim(valueOf("hItemType"));
		var option;
		var typeId;
		if (!select) {
			return;
		}
		select.length = 1;
		childElements(root, "Type").forEach(function (node) {
			typeId = trim(attr(node, "ID"));
			if (itemType !== "select" && typeId === "2" && itemType !== "3") {
				return;
			}
			option = document.createElement("option");
			option.value = typeId;
			option.text = attr(node, "Name");
			select.add(option);
		});
	}

	function Init() {
		populateHeader();
		if (Number(valueOf("hValue")) > 0) {
			setSelectByValue(field("selHeader"), valueOf("hHeader"));
			setSelectByValue(field("selDataType"), valueOf("hType"));
			setValue("txtAttribute", valueOf("hAttribute"));
			setValue("txtDataLen", valueOf("hLength"));
			setValue("txtDecimals", valueOf("hDecimal"));
			configureDataTypeControls(valueOf("hType"));
		}
	}

	function populateUpdate() {
		var root;
		var attributes;
		var count;
		ClearTable();
		getText("ItmTypeAttributeDisplayEdit.asp?ItemType=" + encodeURIComponent(valueOf("hItemType")) + "&ClassCode=" + encodeURIComponent(valueOf("hClassCode")));
		root = xmlRoot("EditData");
		attributes = childElements(root, "Attribute");
		if (Number(valueOf("hValue")) > 0) {
			copyExistingAttributeToNewData(root);
		}
		if (attributes.length) {
			attributes.forEach(function (node, index) {
				addManageRow(index + 1, node);
				addSelectRow(index + 1, node);
			});
			count = attributes.length;
		} else {
			addNoRecordsRow("tblDisplay", 7);
			addNoRecordsRow("tblAttSelect", 6);
			count = 0;
		}
		setValue("hRow", attributes.length || 1);
		setValue("hRowCtr", count);
	}

	function Init2() {
		var modal = window.ITMSModalReturnCompat;
		var getDialogRoot = modal && modal["dialog" + "ArgumentsRoot"];
		var root = getDialogRoot ? getDialogRoot() : null;
		var count = Number(valueOf("hRowCtr")) || 0;
		var attId;
		var value;
		var i;
		childElements(root, "Attribute").forEach(function (node) {
			attId = trim(attr(node, "AttID"));
			value = attr(node, "Value");
			for (i = 1; i <= count; i += 1) {
				if (trim(valueOf("hAttIDSZ" + i)) === attId) {
					setValue("hAttValSZ" + i, value);
					if (field("ChkSelect" + i)) {
						field("ChkSelect" + i).checked = true;
					}
					break;
				}
			}
		});
	}

	function checkSelect() {
		var root = xmlRoot("OutData");
		var dataType = selectedValue("selDataType");
		if (root) {
			clearChildren(root);
		}
		optionDialogHasData = false;
		outDataValue = null;
		resetDataLengthFields();
		configureDataTypeControls(dataType);
		if (dataType !== "Options") {
			return false;
		}
		if (!validateOptionContext()) {
			return false;
		}
		openDialog("ItmTypOptionPoPEntry.asp?sTemp=" + encodeURIComponent(valueOf("hItemType") + ":" + selectedText(field("selHeader")) + ":" + valueOf("txtAttribute")), "", "dialogHeight:270px;dialogWidth:435px;center:Yes;help:No;resizable:No;status:No", function (returnedValue) {
			var returnedRoot = xmlRoot(returnedValue);
			outDataValue = returnedRoot;
			if (!returnedRoot || !childElements(returnedRoot).length) {
				field("selDataType").selectedIndex = 0;
				optionDialogHasData = false;
			} else {
				optionDialogHasData = true;
			}
		});
		return false;
	}

	function popOptionSel() {
		var dataType = selectedValue("selDataType");
		resetDataLengthFields();
		configureDataTypeControls(dataType);
		if (dataType !== "Options") {
			return false;
		}
		if (!validateOptionContext()) {
			return false;
		}
		openDialog("ItmTypOptionPoPEntryUpdate.asp?sTemp=" + encodeURIComponent(valueOf("hItemType") + ":" + selectedText(field("selHeader")) + ":" + valueOf("txtAttribute")), xmlObject("EditData"), "dialogHeight:300px;dialogWidth:435px;center:Yes;help:No;resizable:No;status:No", function (returnedValue) {
			var returnedRoot = xmlRoot(returnedValue);
			var targetRoot = xmlRoot("NewData");
			if (!returnedRoot || !targetRoot) {
				return;
			}
			clearChildren(targetRoot);
			childElements(returnedRoot, "Attribute").forEach(function (node) {
				targetRoot.appendChild(importNode(targetRoot, node));
			});
		});
		return false;
	}

	function CheckSubmit() {
		var root;
		var response;
		var code;
		if (!validateAttributeForm()) {
			return false;
		}
		root = xmlRoot("OutData");
		if (!root) {
			return false;
		}
		clearChildren(root);
		setAttributePayload(root, false);
		if (optionDialogHasData && outDataValue && childElements(outDataValue).length) {
			childElements(outDataValue).forEach(function (node) {
				root.appendChild(importNode(root, node));
			});
		}
		response = postXml("ItmTypeAttributeInsert.asp", root);
		code = responseCode(response);
		if (code.indexOf("Y<BR") === 0 || code === "Y") {
			alert("Item Type Attribute " + trim(valueOf("txtAttribute")) + " has been defined Successfully");
			form().submit();
		} else if (code.indexOf("N<BR") === 0 || code === "N") {
			alert("Item Type Attribute " + trim(valueOf("txtAttribute")) + " has been already defined");
			form().submit();
		} else {
			alert(response);
		}
		return false;
	}

	function CheckUpdate() {
		var root;
		var response;
		var code;
		if (!validateAttributeForm()) {
			return false;
		}
		root = xmlRoot("NewData");
		if (!root) {
			return false;
		}
		setAttributePayload(root, true);
		if (selectedValue("selDataType") !== "Options") {
			clearChildren(root);
		}
		response = postXml("ItmTypeAttributeUpdate.asp", root);
		code = responseCode(response);
		if (code.indexOf("Y<BR") === 0 || code === "Y") {
			alert("Item Type Attribute " + trim(valueOf("txtAttribute")) + " has been Updated Successfully");
			form().submit();
		} else {
			alert(response);
		}
		return false;
	}

	function ManageAttribute(mode) {
		var manage = document.getElementById("divManageAtt");
		var select = document.getElementById("divSelectAtt");
		if (!manage || !select) {
			return false;
		}
		if (mode === "M") {
			manage.style.display = manage.style.display === "none" ? "block" : "none";
			select.style.display = "none";
		} else {
			select.style.display = select.style.display === "none" ? "block" : "none";
			manage.style.display = "none";
		}
		return false;
	}

	function SelectVal() {
		var root = xmlRoot("AttData");
		var count = Number(valueOf("hRowCtr")) || 0;
		var doc = xmlDocument("AttData");
		var node;
		var sourceIndex;
		if (!root || !doc) {
			return false;
		}
		clearChildren(root);
		for (var i = 1; i <= count; i += 1) {
			if (field("ChkSelect" + i) && field("ChkSelect" + i).checked) {
				sourceIndex = field("ChkSelect" + i).value;
				node = doc.createElement("Attribute");
				node.setAttribute("AttID", valueOf("hAttIDSZ" + sourceIndex));
				node.setAttribute("Head", valueOf("hHeaderSZ" + sourceIndex));
				node.setAttribute("AttName", valueOf("hAttNameSZ" + sourceIndex));
				node.setAttribute("Type", valueOf("hTypeSZ" + sourceIndex));
				node.setAttribute("Length", valueOf("hLengthSZ" + sourceIndex));
				node.setAttribute("Decimal", valueOf("hDecimalSZ" + sourceIndex));
				node.setAttribute("Value", valueOf("hAttValSZ" + sourceIndex));
				root.appendChild(node);
			}
		}
		if (window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.returnAndClose) {
			window.ITMSModalReturnCompat.returnAndClose(root);
		} else {
			window.close();
		}
		return false;
	}

	function EditAttribute(index) {
		var attrId = valueOf("hAttIDZ" + index);
		form().action = "ItmTypeAttributeEntry.asp?sValue=" + encodeURIComponent(attrId) + "&Mod=U&ClassCode=" + encodeURIComponent(valueOf("hClassCode"));
		form().submit();
		return false;
	}

	function DeleteAttribute() {
		var count = Number(valueOf("hRowCtr")) || 0;
		var selectedCount = 0;
		var selectedIndex = "";
		var item;
		for (var i = 1; i <= count; i += 1) {
			item = field("Chk" + i);
			if (item && item.checked) {
				selectedCount += 1;
				selectedIndex = item.value;
			}
		}
		if (selectedCount !== 1) {
			alert("Select any one attribute to Delete");
			return false;
		}
		form().action = "ItmTypeAttributeDelete.asp?sValue=" + encodeURIComponent(valueOf("hAttIDZ" + selectedIndex)) + "&ItemType=" + encodeURIComponent(valueOf("hItemTypeZ" + selectedIndex)) + "&ClassCode=" + encodeURIComponent(valueOf("hClassCode"));
		form().submit();
		return false;
	}

	function openDetails() {
		openDialog("XMLItemTypeAttributesView.asp", "A", "dialogHeight:510px;dialogWidth:620px;center:Yes;help:No;resizable:No;status:No", function () {});
		return false;
	}

	function windowOnUnload() {
		var root = xmlRoot("AttData");
		if (window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.returnValue) {
			window.ITMSModalReturnCompat.returnValue(root);
		}
		return root;
	}

	if (window.ITMSModalReturnCompat) {
		window.ITMSModalReturnCompat.install(function () {
			return xmlRoot("AttData");
		});
	}

	window.checkNumbers = checkNumbers;
	window.checkSelect = checkSelect;
	window.CheckSubmit = CheckSubmit;
	window.Init = Init;
	window.ClearTable = ClearTable;
	window.populateUpdate = populateUpdate;
	window.EditAttribute = EditAttribute;
	window.DeleteAttribute = DeleteAttribute;
	window.populateHeader = populateHeader;
	window.openDetails = openDetails;
	window.window_onunload = windowOnUnload;
	window.Init2 = Init2;
	window.popOptionSel = popOptionSel;
	window.CheckUpdate = CheckUpdate;
	window.ManageAttribute = ManageAttribute;
	window.SelectVal = SelectVal;
}(window, document));
