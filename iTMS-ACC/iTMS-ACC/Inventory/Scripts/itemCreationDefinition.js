(function (window, document) {
	"use strict";

	var state = {
		j: 0,
		iClass: "0"
	};

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function isSimplePage() {
		return /(?:ItemCreateSimple|ItemEditSimple)\.asp/i.test(window.location.pathname);
	}

	function isEditPage() {
		return /(?:ItemEditSimple|ItmEditEntry)\.asp/i.test(window.location.pathname);
	}

	function trimValue(value) {
		return String(value == null ? "" : value).replace(/\u00a0/g, " ").replace(/^\s+|\s+$/g, "");
	}

	function num(value) {
		var parsed = parseFloat(trimValue(value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function formatNumber(value, decimals) {
		return num(value).toFixed(decimals == null ? 2 : decimals);
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
			element.textContent = String(value == null ? "" : value);
		}
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] : null;
	}

	function fieldValue(name, defaultValue) {
		var item = field(name);
		return item ? item.value : defaultValue || "";
	}

	function selectedText(select) {
		return select && select.options && select.selectedIndex >= 0 ? select.options[select.selectedIndex].text : "";
	}

	function selectedValue(select) {
		return select && select.options && select.selectedIndex >= 0 ? select.options[select.selectedIndex].value : select && select.value || "";
	}

	function setSelectValue(select, value) {
		if (!select || !select.options) {
			return false;
		}
		for (var i = 0; i < select.options.length; i += 1) {
			if (trimValue(select.options[i].value) === trimValue(value)) {
				select.selectedIndex = i;
				return true;
			}
		}
		return false;
	}

	function focusControl(control, selectText) {
		if (!control) {
			return;
		}
		if (selectText && control.select) {
			control.select();
		} else if (control.focus) {
			control.focus();
		}
	}

	function controlArray(name) {
		var control = field(name);
		if (!control) {
			return [];
		}
		if (typeof control.length === "number" && control.tagName !== "SELECT") {
			return Array.prototype.slice.call(control);
		}
		return [control];
	}

	function radio(name, index) {
		return controlArray(name)[index] || null;
	}

	function checkedRadioValue(name, defaultValue) {
		var radios = controlArray(name);
		for (var i = 0; i < radios.length; i += 1) {
			if (radios[i] && radios[i].checked) {
				return radios[i].value;
			}
		}
		return defaultValue || "";
	}

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function xmlIsland(name) {
		ensureCompat();
		return window[name] || document[name] || byId(name);
	}

	function xmlDocument(name) {
		var island = xmlIsland(name);
		return island && (island.XMLDocument || island._doc || island.ownerDocument && island.nodeType === 1 && island.ownerDocument || island.nodeType === 9 && island) || null;
	}

	function xmlRoot(nameOrObject) {
		var object = typeof nameOrObject === "string" ? xmlIsland(nameOrObject) : nameOrObject;
		if (!object) {
			return null;
		}
		if (object.documentElement) {
			return object.documentElement;
		}
		if (object.XMLDocument && object.XMLDocument.documentElement) {
			return object.XMLDocument.documentElement;
		}
		if (object._doc && object._doc.documentElement) {
			return object._doc.documentElement;
		}
		if (object.nodeType === 1) {
			return object;
		}
		return null;
	}

	function docFor(name) {
		return xmlDocument(name) || document.implementation.createDocument("", "Root", null);
	}

	function serializeXml(nameOrObject) {
		var doc = typeof nameOrObject === "string" ? xmlDocument(nameOrObject) : xmlDocumentFromRoot(nameOrObject);
		if (!doc) {
			return "";
		}
		return new XMLSerializer().serializeToString(doc);
	}

	function xmlDocumentFromRoot(root) {
		if (!root) {
			return null;
		}
		if (root.nodeType === 9) {
			return root;
		}
		if (root.XMLDocument) {
			return root.XMLDocument;
		}
		if (root.documentElement) {
			return root;
		}
		return root.ownerDocument || null;
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

	function removeDirectChildren(root, name) {
		var children = childElements(root, name);
		for (var i = 0; i < children.length; i += 1) {
			root.removeChild(children[i]);
		}
	}

	function appendImported(parent, node) {
		var imported;
		if (!parent || !node) {
			return null;
		}
		imported = node;
		if (node.ownerDocument && node.ownerDocument !== parent.ownerDocument) {
			imported = parent.ownerDocument.importNode(node, true);
		}
		return parent.appendChild(imported);
	}

	function setAttr(node, name, value) {
		if (node) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function getAttr(node, name) {
		return node && node.getAttribute ? node.getAttribute(name) || "" : "";
	}

	function selectNodes(root, expression, tagName) {
		var nodes;
		if (!root) {
			return [];
		}
		if (root.selectNodes) {
			nodes = root.selectNodes(expression);
			return Array.prototype.slice.call(nodes || []);
		}
		return Array.prototype.slice.call(root.getElementsByTagName(tagName || "*"));
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
		if (body != null) {
			xhr.setRequestHeader("Content-Type", "text/xml");
		}
		xhr.send(body == null ? null : body);
		return xhr;
	}

	function responseXmlText(xhr) {
		if (xhr.responseXML && xhr.responseXML.documentElement && window.XMLSerializer) {
			return new XMLSerializer().serializeToString(xhr.responseXML);
		}
		return xhr.responseText || "";
	}

	function loadXmlFromResponse(islandName, xhr) {
		var island = xmlIsland(islandName);
		var text = responseXmlText(xhr);
		if (island && island.loadXML && trimValue(text) !== "") {
			island.loadXML(text);
		}
		return xmlRoot(islandName);
	}

	function openDialog(url, args, features, callback) {
		ensureCompat();
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			return window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
		}
		return window.open(url, "_blank", "width=650,height=500,resizable=yes,scrollbars=yes");
	}

	function AssaignValue() {
		CalcValue("MP");
		CalcValue("OP");
	}

	function CalcValue(callFrom) {
		var frm = form();
		var rate = num(fieldValue("txtPurRate"));
		var percentage;
		if (rate === 0) {
			return;
		}
		if (callFrom === "MP") {
			percentage = fieldValue("txtMarPer");
			if (trimValue(percentage) === "") {
				return;
			}
			if (num(percentage) < 0) {
				alert("Enter Numerals Only");
				frm.txtMarPer.value = 0;
				return;
			}
			frm.txtMarVal.value = formatNumber(rate * (num(percentage) / 100), 2);
		} else if (callFrom === "MV") {
			frm.txtMarPer.value = Math.round((num(frm.txtMarVal.value) * 1000) / rate) / 10;
		} else if (callFrom === "OP") {
			percentage = fieldValue("txtChaPer");
			if (trimValue(percentage) === "") {
				return;
			}
			if (num(percentage) < 0) {
				alert("Enter Numerals Only");
				frm.txtChaVal.value = 0;
				return;
			}
			frm.txtChaVal.value = formatNumber(rate * (num(percentage) / 100), 2);
		} else if (callFrom === "OV") {
			frm.txtChaPer.value = Math.round((num(frm.txtChaVal.value) * 1000) / rate) / 10;
		}
		frm.txtTotPrice.value = num(frm.txtPurRate.value) + num(frm.txtMarVal.value) + num(frm.txtChaVal.value);
	}

	function ItemSpecPop() {
		var frm = form();
		var className = fieldValue("txtClass");
		var itemType = selectedValue(frm.selIType);
		var classCode = fieldValue("hClassCode");
		if (isEditPage()) {
			return EditItemSpecPop();
		}
		if (trimValue(className) === "") {
			alert("Select the Category");
			return false;
		}
		openDialog("ItmSpecsPop.asp?ItemType=" + encodeURIComponent(itemType) + "&ClassCode=" + encodeURIComponent(classCode), xmlIsland("OutData"), "dialogHeight:300px;dialogWeight:250px;Status:No", function (returnedValue) {
			var outRoot = xmlRoot("OutData");
			var returnedRoot = xmlRoot(returnedValue);
			removeDirectChildren(outRoot, "ItemSpecs");
			if (returnedRoot) {
				appendImported(outRoot, returnedRoot);
			}
		});
		return true;
	}

	function ChangeBOM(obj) {
		var radios = controlArray("radBOM").length ? controlArray("radBOM") : controlArray("radBoM");
		var value = selectedValue(obj);
		if (!radios.length) {
			return;
		}
		if (value === "4") {
			radios[0].checked = true;
		} else if (radios[1]) {
			radios[1].checked = true;
		}
	}

	function ChangeAttribute() {
		return true;
	}

	function UploadImage() {
		if (isEditPage()) {
			return EditUploadImage();
		}
		alert("Item details need to be saved before uploading the image ");
	}

	function DeleteImage() {
		var itemCode;
		var response;
		if (!isEditPage()) {
			return false;
		}
		itemCode = fieldValue("hItemCode");
		if (!confirm("Do you want to remove the image?")) {
			return false;
		}
		response = syncPost("ItemImageDelete.asp?ItemCode=" + encodeURIComponent(itemCode));
		if (trimValue(response.responseText) === "") {
			form().submit();
		} else {
			alert(response.responseText);
		}
		return true;
	}

	function Init() {
		var frm = form();
		if (isEditPage()) {
			return EditInit();
		}
		if (window.LetIType) {
			window.LetIType(frm.selIType);
		}
		if (window.ChangeLabel) {
			window.ChangeLabel(frm.selIType);
		}
	}

	function CheckNoSeries() {
		return true;
	}

	function SelectAccHead() {
		var frm = form();
		var url = "ItmAccHead.asp?OAH=" + encodeURIComponent(fieldValue("hOAH")) + "&CAH=" + encodeURIComponent(fieldValue("hCAH"));
		openDialog(url, "", "dialogHeight:400px;dialogWidth:500px;Status:No", function (returnedValue) {
			var root = xmlRoot(returnedValue);
			var nodes = childElements(root, "AccHead");
			var node;
			for (var i = 0; i < nodes.length; i += 1) {
				node = nodes[i];
				frm.hOAH.value = getAttr(node, "OAHV");
				setText("spanOpenAccHead", getAttr(node, "OAHN"));
				frm.hCAH.value = getAttr(node, "CAHV");
				setText("spanCloseAccHead", getAttr(node, "CAHN"));
			}
		});
	}

	function ManageAttribute() {
		var frm = form();
		var itemType;
		var classCode;
		if (!isSimplePage() && frm.selIType.selectedIndex === 0) {
			alert("Select Item Type");
			frm.selIType.focus();
			return false;
		}
		if (trimValue(fieldValue("txtClass")) === "") {
			alert("Select Category");
			return false;
		}
		itemType = selectedValue(frm.selIType);
		classCode = fieldValue("hClassCode");
		openDialog("ItmTypeAttributeEntry.asp?ItemType=" + encodeURIComponent(itemType) + "&ClassCode=" + encodeURIComponent(classCode), xmlIsland("ItemAttData"), "dialogHeight:400px;dialogWidth:500px;Status:No", function (returnedValue) {
			var itemAttRoot = xmlRoot("ItemAttData");
			var returnedRoot = xmlRoot(returnedValue);
			var attrs = childElements(returnedRoot, "Attribute");
			removeDirectChildren(itemAttRoot, "Attribute");
			for (var i = 0; i < attrs.length; i += 1) {
				appendImported(itemAttRoot, attrs[i]);
			}
			DisplayAttTable();
		});
		return true;
	}

	function ClearAttTable() {
		var table = byId("tblAttribute");
		while (table && table.rows.length) {
			table.deleteRow(0);
		}
	}

	function createOptionInput(attribute) {
		var select = document.createElement("select");
		var options = childElements(xmlRoot("OptionData"), "Option");
		var attId = getAttr(attribute, "AttID");
		var value = getAttr(attribute, "Value");
		var option;
		select.name = "selAttZ" + attId;
		select.className = "FormElem";
		for (var i = 0; i < options.length; i += 1) {
			if (getAttr(options[i], "AttID") !== trimValue(attId)) {
				continue;
			}
			option = document.createElement("option");
			option.text = getAttr(options[i], "Name");
			option.value = getAttr(options[i], "ID");
			select.add(option);
			if (trimValue(option.value) === trimValue(value)) {
				select.selectedIndex = select.options.length - 1;
			}
		}
		return select;
	}

	function createAttributeControl(attribute) {
		var type = trimValue(getAttr(attribute, "Type"));
		var input;
		if (type === "OPTIONS") {
			return createOptionInput(attribute);
		}
		input = document.createElement("input");
		input.type = "text";
		input.name = "txtAtt" + getAttr(attribute, "AttID");
		input.className = "FormElem";
		input.maxLength = String(num(getAttr(attribute, "Length")) + (type === "NUMERIC" ? 1 : 0));
		input.value = getAttr(attribute, "Value");
		return input;
	}

	function DisplayAttTable() {
		var table = byId("tblAttribute");
		var headers = childElements(xmlRoot("HeadData"), "Header");
		var attrs = childElements(xmlRoot("ItemAttData"), "Attribute");
		var row;
		var cell;
		var nested;
		var nestedRow;
		var nestedCell;
		var headId;
		var firstForHead;
		ClearAttTable();
		if (!table || !headers.length) {
			return;
		}
		row = table.insertRow(table.rows.length);
		for (var i = 0; i < headers.length; i += 1) {
			headId = getAttr(headers[i], "ID");
			cell = row.insertCell(row.cells.length);
			cell.className = "FieldCell";
			cell.vAlign = "top";
			nested = document.createElement("table");
			nested.id = "tblHead" + headId;
			nested.border = "0";
			nested.width = "100%";
			nested.className = "BodyTable";
			cell.appendChild(nested);
			firstForHead = true;
			for (var jAttr = 0; jAttr < attrs.length; jAttr += 1) {
				if (getAttr(attrs[jAttr], "Head") !== headId) {
					continue;
				}
				if (firstForHead) {
					nestedRow = nested.insertRow(nested.rows.length);
					nestedCell = nestedRow.insertCell(0);
					nestedCell.innerHTML = getAttr(headers[i], "Name");
					nestedCell.className = "ExcelHeaderCell";
					nestedCell.align = "center";
					if (headId === "3") {
						nestedCell.colSpan = 2;
					}
				}
				nestedRow = nested.insertRow(nested.rows.length);
				nestedCell = nestedRow.insertCell(0);
				nestedCell.innerHTML = getAttr(attrs[jAttr], "AttName");
				nestedCell.className = "FieldCellSub";
				if (headId === "3") {
					nestedCell = nestedRow.insertCell(1);
					nestedCell.className = "FieldCellSub";
					nestedCell.appendChild(createAttributeControl(attrs[jAttr]));
				}
				firstForHead = false;
			}
		}
	}

	function ClearTable() {
		var table = byId("tblData");
		while (table && table.rows.length > 1) {
			table.deleteRow(1);
		}
		state.j = 0;
	}

	function parseClassificationReturn(returnedValue) {
		var parts = String(returnedValue || "").split("*****");
		var values;
		var names;
		var categories = [];
		var classes = [];
		var classNames = [];
		var item;
		if (parts[0] === "-1" || !parts[0]) {
			return null;
		}
		values = parts[0].split("|");
		for (var i = 0; i < values.length; i += 1) {
			item = values[i].split(":");
			if (item.length) {
				classes.push(item[item.length - 1]);
				categories.push(item[1] || "");
			}
		}
		names = String(parts[1] || "").split("|||");
		for (var jName = 0; jName < names.length; jName += 1) {
			item = names[jName].split(":");
			if (item.length) {
				classNames.push(item[item.length - 1]);
			}
		}
		return {
			categories: categories,
			classes: classes,
			names: classNames
		};
	}

	function loadCategoryNames(categoryCodes) {
		var root;
		var nodes;
		var names = [];
		if (!categoryCodes.length) {
			return "";
		}
		root = loadXmlFromResponse("CategoryData", syncGet("GetCategoryCode.asp?Code=" + encodeURIComponent(categoryCodes.join(","))));
		nodes = childElements(root, "CATEGORY");
		for (var i = 0; i < nodes.length; i += 1) {
			names.push(getAttr(nodes[i], "NAME"));
		}
		return names.join(",");
	}

	function popClass() {
		var frm = form();
		var orgId = fieldValue("hUnitID");
		var itemType = fieldValue("selIType");
		var itemTypeName = selectedText(frm.selIType);
		var url = "/include/ClassificationSelectPop.asp?sIType=" + encodeURIComponent(itemType) + "&sOrgID=" + encodeURIComponent(orgId) + "&sITypename=" + encodeURIComponent(itemTypeName);
		openDialog(url, "Classification", "dialogHeight:500px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No", function (returnedValue) {
			var parsed = parseClassificationReturn(returnedValue);
			var root = xmlRoot("OutData");
			var node;
			var categoryName;
			var classCode;
			var response;
			if (!parsed) {
				return;
			}
			categoryName = loadCategoryNames(parsed.categories);
			if (trimValue(categoryName) !== "") {
				setText("spanCategory", categoryName);
			}
			removeDirectChildren(root, "CLASSIFICATION");
			for (var i = 0; i < parsed.classes.length; i += 1) {
				classCode = String(parsed.classes[i]);
				node = docFor("OutData").createElement("CLASSIFICATION");
				setAttr(node, "CODE", classCode);
				setAttr(node, "NAME", parsed.names[i] || "");
				setAttr(node, "CATEGORY", parsed.categories[i] || "");
				root.appendChild(node);
			}
			frm.txtClass.value = parsed.names.join(", ");
			frm.hClassCode.value = parsed.classes[parsed.classes.length - 1] || "";
			frm.hCategory.value = parsed.categories.join(",");
			response = syncGet("../../Include/GetClassAttribute.asp?ClassCode=" + encodeURIComponent(frm.hClassCode.value));
			if (trimValue(response.responseText) === "Y") {
				ManageAttribute();
			}
		});
	}

	function GetDetails() {
		var frm = form();
		var root;
		var bom;
		if (radio("radBoM", 1) && radio("radBoM", 1).checked) {
			return false;
		}
		if (!isSimplePage() && frm.selIType.selectedIndex === 0) {
			alert("Select Item Type");
			frm.selIType.focus();
			return false;
		}
		if (trimValue(fieldValue("txtItmDesc")) === "") {
			alert("Enter Item Description");
			focusControl(frm.txtItmDesc, true);
			return false;
		}
		root = xmlRoot("OutData");
		if (!childElements(root, "BOM").length) {
			bom = docFor("OutData").createElement("BOM");
			root.appendChild(bom);
		}
		openDialog("ItmBoMEntry.asp?sTemp=" + encodeURIComponent(fieldValue("txtItmDesc") + ":" + fieldValue("txtItmDesc") + ":" + fieldValue("selIType")), xmlIsland("OutData"), "dialogHeight:450px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No", function () {});
		return true;
	}

	function GetAttr() {
		if (isEditPage()) {
			return GetEditAttr();
		}
		var root = xmlRoot("AttrData");
		var select = field("selAttr");
		var classCode = fieldValue("hClassCode");
		var nodes = selectNodes(root, "//ROOT/ATTRIBUTES [@ClassCode='" + classCode + "']", "ATTRIBUTES");
		var option;
		if (!select) {
			return;
		}
		select.length = 0;
		for (var i = 0; i < nodes.length; i += 1) {
			if (getAttr(nodes[i], "ClassCode") !== classCode) {
				continue;
			}
			option = document.createElement("option");
			option.text = getAttr(nodes[i], "ATTRNAME");
			option.value = getAttr(nodes[i], "ATTRID");
			select.add(option);
		}
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

	function createTextInput(name, value, size, className, keyScript, disabled) {
		var input = document.createElement("input");
		input.type = "text";
		input.name = name;
		input.value = value || "";
		input.size = size;
		input.className = className || "FormElem";
		input.style.textAlign = "right";
		if (keyScript) {
			input.setAttribute("onkeypress", keyScript);
		}
		if (disabled) {
			input.disabled = true;
		}
		return input;
	}

	function GetStockDet() {
		if (isEditPage()) {
			return GetEditStockDet();
		}
		var root = xmlRoot("OutData");
		var doc = docFor("OutData");
		var table = byId("tblData");
		var details;
		var itemDetail;
		var storageNodes;
		var storage;
		var unitStore;
		var orgId;
		var loc;
		var bin;
		var storeName;
		var row;
		var button;
		var classCode = fieldValue("hClassCode") || state.iClass || "0";
		var date = fieldValue("hFinFrom");
		var newStorage;
		var displayIndex = 0;
		ClearTable();
		if (!root || !table) {
			return;
		}
		details = doc.createElement("Details");
		root.appendChild(details);
		itemDetail = doc.createElement("ItemDetail");
		setAttr(itemDetail, "ItemCode", "0");
		details.appendChild(itemDetail);
		storageNodes = selectNodes(root, "//STOREDET", "STOREDET");
		for (var i = 0; i < storageNodes.length; i += 1) {
			storage = storageNodes[i];
			unitStore = trimValue(getAttr(storage, "UNITSTORE")).split("-");
			orgId = unitStore[0] || "";
			loc = unitStore[1] || "";
			bin = unitStore[2] || "";
			storeName = getAttr(storage, "STORE");
			displayIndex += 1;
			if (displayIndex === 1) {
				setText("idStore", storeName);
			}
			row = table.insertRow(state.j + 1);
			appendCell(row, "ExcelSerial", "center", displayIndex);
			appendCell(row, "ExcelDisplayCell", "left", storeName);
			appendCell(row, "ExcelInputCell", "", createTextInput("txtMYX" + loc + "X" + bin, date, 10, isSimplePage() ? "Formelem" : "FormElem", "DoKeyPress('N',6,0)", true)).width = "8";
			appendCell(row, "ExcelInputCell", "", createTextInput("txtQTYX" + loc + "X" + bin, "", 12, isSimplePage() ? "Formelem" : "FormElem", "DoKeyPress('Y',7,3)")).width = "10";
			appendCell(row, "ExcelInputCell", "", createTextInput("txtVALX" + loc + "X" + bin, "", 12, isSimplePage() ? "Formelem" : "FormElem", "DoKeyPress('Y',10,2)")).width = "10";
			button = document.createElement("input");
			button.type = "button";
			button.value = " Yes ";
			button.name = "btn:" + classCode + ":" + orgId + ":" + loc + ":" + bin;
			button.size = 12;
			button.maxLength = 10;
			button.className = "AddButtonX";
			button.onclick = (function (name) {
				return function () {
					GetLot(this, name);
				};
			}(storeName));
			appendCell(row, "ExcelFieldCell", "center", button);
			newStorage = doc.createElement("STORAGE");
			setAttr(newStorage, "STORE", trimValue(loc));
			setAttr(newStorage, "BIN", trimValue(bin));
			setAttr(newStorage, "MONTHYEAR", "");
			setAttr(newStorage, "QTY", "");
			setAttr(newStorage, "STORAGEVALUE", "");
			setAttr(newStorage, "CLASSIFICATION", "");
			setAttr(newStorage, "UNIT", orgId);
			itemDetail.appendChild(newStorage);
			state.j += 1;
		}
		if (field("hOpeningStockUnit")) {
			field("hOpeningStockUnit").value = orgId || "";
		}
	}

	function updateAttributeInputValues() {
		var itemRoot = xmlRoot("ItemAttData");
		var attrs = childElements(itemRoot, "Attribute");
		var control;
		var value;
		for (var i = 0; i < attrs.length; i += 1) {
			if (getAttr(attrs[i], "Head") !== "3") {
				continue;
			}
			if (trimValue(getAttr(attrs[i], "Type")) === "OPTIONS") {
				control = field("selAttZ" + getAttr(attrs[i], "AttID"));
				value = selectedValue(control);
			} else {
				control = field("txtAtt" + getAttr(attrs[i], "AttID"));
				value = control ? control.value : "";
			}
			if (trimValue(value) === "") {
				alert("Please enter value for " + getAttr(attrs[i], "AttName"));
				focusControl(control, false);
				return false;
			}
			setAttr(attrs[i], "Value", value);
		}
		return true;
	}

	function itemAttributeList(forLotDialog) {
		var attrs = childElements(xmlRoot("ItemAttData"), "Attribute");
		var ids = [];
		for (var i = 0; i < attrs.length; i += 1) {
			if (!forLotDialog || isSimplePage() || getAttr(attrs[i], "Head") === "2") {
				ids.push(getAttr(attrs[i], "AttID"));
			}
		}
		return ids.length ? ids.join(",") : "NULL";
	}

	function updateReturnedStorageFields(root) {
		var storageNodes = selectNodes(root, "//STORAGE", "STORAGE");
		var store;
		var bin;
		for (var i = 0; i < storageNodes.length; i += 1) {
			store = getAttr(storageNodes[i], "STORE");
			bin = getAttr(storageNodes[i], "BIN");
			if (field("txtQTYX" + store + "X" + bin)) {
				field("txtQTYX" + store + "X" + bin).value = getAttr(storageNodes[i], "QTY");
			}
			if (field("txtVALX" + store + "X" + bin)) {
				field("txtVALX" + store + "X" + bin).value = getAttr(storageNodes[i], "STORAGEVALUE");
			}
		}
	}

	function GetLot(obj, storeName) {
		if (isEditPage()) {
			return GetEditLot(obj, storeName);
		}
		var frm = form();
		var parts = String(obj && obj.name || "").split(":");
		var classCode = parts[1] || "0";
		var orgId = parts[2] || "";
		var loc = parts[3] || "";
		var bin = parts[4] || "";
		var qty = field("txtQTYX" + loc + "X" + bin);
		var value = field("txtVALX" + loc + "X" + bin);
		var attrList;
		var tempValues;
		var url = "";
		if (frm.selRecNum.selectedIndex === 0) {
			alert("Select Receipt Numbering");
			frm.selRecNum.focus();
			return false;
		}
		if (!qty || trimValue(qty.value) === "") {
			alert("Enter Quantity");
			focusControl(qty, true);
			return false;
		}
		if (!value || trimValue(value.value) === "") {
			alert("Enter Value");
			focusControl(value, true);
			return false;
		}
		if (!updateAttributeInputValues()) {
			return false;
		}
		attrList = itemAttributeList(true);
		tempValues = [
			frm.selRecNum.value,
			"0",
			classCode,
			orgId,
			loc,
			bin,
			qty.value,
			String(storeName || "").replace(/&/g, "and"),
			frm.selUoMStores.value,
			value.value,
			trimValue(frm.txtItmDesc.value).replace(/'/g, "~~"),
			frm.selIType.value,
			attrList
		].join("``");
		if (frm.selRecNum.value === "N" && attrList !== "" && attrList !== "NULL") {
			url = "../../Common/PackingLotSerialDetailsForNone.asp?sTemp=" + encodeURIComponent(tempValues);
		} else if (frm.selRecNum.value !== "N") {
			url = "../../Common/PackingLotSerialDetails.asp?sTemp=" + encodeURIComponent(tempValues);
		}
		if (!url) {
			return true;
		}
		openDialog(url, xmlIsland("OutData"), "dialogHeight:580px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No", function (returnedValue) {
			var root = xmlRoot(returnedValue) || xmlRoot("OutData");
			updateReturnedStorageFields(root);
			syncPost("XMLSave.asp?SessionFlag=False&Value=LotData&Folder=Master", serializeXml("OutData"));
		});
		return true;
	}

	function left(value, count) {
		return String(value || "").substring(0, count);
	}

	function right(value, count) {
		var text = String(value || "");
		return text.substring(Math.max(0, text.length - count));
	}

	function validateStorageRows(finFromOriginal, finToOriginal) {
		var frm = form();
		var finFromP = String(num(right(finFromOriginal, 4)) - 1) + left(finFromOriginal, 2);
		var finFrom = right(finFromOriginal, 4) + left(finFromOriginal, 2);
		var finTo = right(finToOriginal, 4) + left(finToOriginal, 2);
		var elements = frm.elements;
		var control;
		var monthYear;
		for (var i = 10; i < elements.length; i += 1) {
			control = elements[i];
			if (!control || control.type !== "text" || String(control.name || "").indexOf("txtMYX") === -1) {
				continue;
			}
			monthYear = trimValue(control.value);
			if (monthYear.length !== 6) {
				alert("Invalid Format (Month 2 and Year 4 characters)");
				focusControl(control, true);
				return false;
			}
			if (num(left(monthYear, 2)) < 0 || num(left(monthYear, 2)) > 12) {
				alert("Invalid Month");
				focusControl(control, true);
				return false;
			}
			if (!(num(right(monthYear, 4) + left(monthYear, 2)) >= num(finFrom) && num(right(monthYear, 4) + left(monthYear, 2)) <= num(finTo)) && fieldValue("hFinYear") === "1") {
				alert("Month Year should be in current Financial Year");
				focusControl(control, true);
				return false;
			}
			if (!(num(right(monthYear, 4) + left(monthYear, 2)) <= num(finFrom) && num(right(monthYear, 4) + left(monthYear, 2)) >= num(finFromP)) && fieldValue("hFinYear") === "0") {
				alert("Month Year should be in last Financial Year");
				focusControl(control, true);
				return false;
			}
			if (!elements[i + 1] || trimValue(elements[i + 1].value) === "") {
				alert("Enter Quantity");
				focusControl(elements[i + 1], true);
				return false;
			}
			if (!elements[i + 2] || trimValue(elements[i + 2].value) === "") {
				alert("Enter Value");
				focusControl(elements[i + 2], true);
				return false;
			}
		}
		return true;
	}

	function appendAttributeData(root) {
		var itemRoot = xmlRoot("ItemAttData");
		var attrs = childElements(itemRoot, "Attribute");
		var doc = xmlDocumentFromRoot(root) || docFor("OutData");
		var attrRoot;
		var detail;
		removeDirectChildren(root, "ATTRIBUTE");
		if (!attrs.length) {
			return;
		}
		attrRoot = doc.createElement("ATTRIBUTE");
		root.appendChild(attrRoot);
		for (var i = 0; i < attrs.length; i += 1) {
			detail = doc.createElement("ATTDET");
			setAttr(detail, "AttID", getAttr(attrs[i], "AttID"));
			setAttr(detail, "Head", getAttr(attrs[i], "Head"));
			setAttr(detail, "AttName", getAttr(attrs[i], "AttName"));
			setAttr(detail, "Type", getAttr(attrs[i], "Type"));
			setAttr(detail, "Length", getAttr(attrs[i], "Length"));
			setAttr(detail, "Decimal", getAttr(attrs[i], "Decimal"));
			setAttr(detail, "Value", getAttr(attrs[i], "Value"));
			attrRoot.appendChild(detail);
		}
	}

	function checkedValue(name, index, defaultValue) {
		var item = radio(name, index);
		return item && item.checked ? item.value : defaultValue || "";
	}

	function effectiveDateValue() {
		var ctl = field("ctlEffDate") || window.ctlEffDate;
		if (!ctl) {
			return "";
		}
		if (ctl.getDate) {
			return ctl.getDate();
		}
		if (ctl.GetDate) {
			return ctl.GetDate();
		}
		return ctl.value || "";
	}

	function appendDetailsAndControls(root, attrList) {
		var frm = form();
		var doc = docFor("OutData");
		var detail = doc.createElement("DETAILS");
		var controls = doc.createElement("CONTROLS");
		var account;
		var modVat = radio("radMod", 0) && radio("radMod", 0).checked ? radio("radMod", 0).value : radio("radMod", 1) && radio("radMod", 1).value || "";
		var bomYes = radio("radBoM", 1);
		var bomNo = radio("radBoM", 0);
		setAttr(detail, "ITYPE", fieldValue("selIType"));
		setAttr(detail, "ICODE", trimValue(fieldValue("txtitmCode")));
		setAttr(detail, "DESC", trimValue(fieldValue("txtItmDesc")));
		setAttr(detail, "SHDESC", "");
		setAttr(detail, "CATALOUGE", "");
		setAttr(detail, "DRAWVER", "");
		setAttr(detail, "VARIANT", trimValue(fieldValue("txtVariant")));
		setAttr(detail, "ADDDESC", trimValue(fieldValue("txtItmAddDesc")));
		setAttr(detail, "UOM", fieldValue("selUoMStores"));
		setAttr(detail, "CATEGORY", fieldValue("hCategory"));
		setAttr(detail, "ATTRIBUTES", attrList);
		setAttr(detail, "GROUP", fieldValue("hGroup"));
		setAttr(detail, "LEVEL", fieldValue("hLevel"));
		setAttr(detail, "UNIT", fieldValue("selUnit"));
		setAttr(detail, "OPSTOCKUNIT", fieldValue("hOpeningStockUnit"));
		setAttr(detail, "MODVAT", modVat);
		setAttr(detail, "PURTAX", selectedValue(frm.selPurTaxType));
		setAttr(detail, "SALTAX", selectedValue(frm.selSalTaxType));
		root.appendChild(detail);

		setAttr(controls, "RECNUM", fieldValue("selRecNum"));
		setAttr(controls, "ROUTING", fieldValue("selRecRout"));
		setAttr(controls, "ACCOUNTING", fieldValue("selAcc"));
		setAttr(controls, "MODVAT", "0");
		setAttr(controls, "REORDERLEVEL", isSimplePage() ? fieldValue("txtReLvl", "0") : "0");
		setAttr(controls, "REORDERQTY", isSimplePage() ? fieldValue("txtReQty", "0") : "0");
		setAttr(controls, "ECOORDERQTY", isSimplePage() ? fieldValue("txtEcQty", "0") : "0");
		if (bomYes && bomYes.checked) {
			setAttr(controls, "BOMAPPLICABLE", bomYes.value);
		} else if (bomNo && bomNo.checked) {
			setAttr(controls, "BOMAPPLICABLE", bomNo.value);
		}
		root.appendChild(controls);

		account = doc.createElement("ACCHEAD");
		setAttr(account, "Name", getText("spanOpenAccHead"));
		setAttr(account, "Type", "O");
		setAttr(account, "Value", fieldValue("hOAH"));
		controls.appendChild(account);

		account = doc.createElement("ACCHEAD");
		setAttr(account, "Name", getText("spanCloseAccHead"));
		setAttr(account, "Type", "C");
		setAttr(account, "Value", fieldValue("hCAH"));
		controls.appendChild(account);
	}

	function updateStorageXml(root) {
		var storageNodes = selectNodes(root, "//STORAGE", "STORAGE");
		var store;
		var bin;
		for (var i = 0; i < storageNodes.length; i += 1) {
			store = getAttr(storageNodes[i], "STORE");
			bin = getAttr(storageNodes[i], "BIN");
			setAttr(storageNodes[i], "MONTHYEAR", trimValue(fieldValue("txtMYX" + store + "X" + bin)));
			setAttr(storageNodes[i], "QTY", trimValue(fieldValue("txtQTYX" + store + "X" + bin)));
			setAttr(storageNodes[i], "STORAGEVALUE", trimValue(fieldValue("txtVALX" + store + "X" + bin)));
			setAttr(storageNodes[i], "CLASSIFICATION", fieldValue("hClassCode"));
		}
	}

	function storageHasLotSerial(root) {
		var storageNodes = selectNodes(root, "//STORAGE", "STORAGE");
		var store;
		var bin;
		var qty;
		var lots;
		if (field("selRecNum") && field("selRecNum").selectedIndex === 4) {
			return true;
		}
		for (var i = 0; i < storageNodes.length; i += 1) {
			store = getAttr(storageNodes[i], "STORE");
			bin = getAttr(storageNodes[i], "BIN");
			qty = num(getAttr(storageNodes[i], "QTY"));
			lots = childElements(storageNodes[i], "LotSerial");
			if (lots.length <= 0 && qty > 0) {
				alert("Enter Lot / Serial Details");
				if (field("txtQTYX" + store + "X" + bin)) {
					field("txtQTYX" + store + "X" + bin).focus();
				}
				return false;
			}
		}
		return true;
	}

	function appendPricing(root) {
		var doc;
		var pricing;
		if (!isSimplePage()) {
			return;
		}
		doc = xmlDocumentFromRoot(root) || docFor("OutData");
		pricing = doc.createElement("PRICING");
		setAttr(pricing, "PURRATE", fieldValue("txtPurRate"));
		setAttr(pricing, "PURRATEPER", fieldValue("txtPurRatePer"));
		setAttr(pricing, "CHARPER", fieldValue("txtChaPer"));
		setAttr(pricing, "CHARVAL", fieldValue("txtChaVal"));
		setAttr(pricing, "MARPER", fieldValue("txtMarPer"));
		setAttr(pricing, "MARVAL", fieldValue("txtMarVal"));
		setAttr(pricing, "TOTPRICE", fieldValue("txtTotPrice"));
		setAttr(pricing, "EFFFROM", effectiveDateValue());
		root.appendChild(pricing);
	}

	function CheckSubmitDetails(finFrom, finTo) {
		if (isEditPage()) {
			return CheckEditSubmitDetails(finFrom, finTo);
		}
		var frm = form();
		var root;
		var attrList;
		var response;
		var parts;
		if (trimValue(fieldValue("hCategory")) === "") {
			alert("Select Category");
			return false;
		}
		if (!isSimplePage() && frm.selIType.selectedIndex === 0) {
			alert("Select Item Type");
			frm.selIType.focus();
			return false;
		}
		if (trimValue(fieldValue("txtClass")) === "") {
			alert("Select Classification");
			popClass();
			return false;
		}
		if (trimValue(fieldValue("txtitmCode")) === "") {
			alert("Enter Item Code");
			focusControl(frm.txtitmCode, true);
			return false;
		}
		if (trimValue(fieldValue("txtItmDesc")) === "") {
			alert("Enter Item Description");
			focusControl(frm.txtItmDesc, true);
			return false;
		}
		if (frm.selUoMStores.selectedIndex === 0) {
			alert("Select Unit Of Measurement");
			frm.selUoMStores.focus();
			return false;
		}
		if (frm.selRecNum.value === "select") {
			alert("Select Receipt Numbering");
			frm.selRecNum.focus();
			return false;
		}
		if (frm.selRecRout.value === "select") {
			alert("Select Receipt Routing");
			frm.selRecRout.focus();
			return false;
		}
		if (frm.selAcc.value === "select") {
			alert("Select Accounting Type");
			frm.selAcc.focus();
			return false;
		}
		if (getText("idStore") === "") {
			alert("Select Storage");
			GetStore();
			return false;
		}
		frm.selUnit.value = frm.chkAllUnit && frm.chkAllUnit.checked ? frm.chkAllUnit.value : frm.hUnitID.value;
		if (state.j > 0 && !validateStorageRows(finFrom, finTo)) {
			return false;
		}
		root = xmlRoot("OutData");
		removeDirectChildren(root, "DETAILS");
		removeDirectChildren(root, "CONTROLS");
		if (!updateAttributeInputValues()) {
			return false;
		}
		attrList = itemAttributeList(false);
		appendAttributeData(root);
		appendDetailsAndControls(root, attrList);
		updateStorageXml(root);
		if (!storageHasLotSerial(root)) {
			return false;
		}
		appendPricing(root);
		response = syncPost("ItmCreationDefinitionInsert.asp", serializeXml("OutData"));
		if (response.responseText.substring(0, 13) === "ItemClassCode") {
			alert("Item has been created and defined Successfully.");
			parts = response.responseText.split(":");
			window.location.href = (isSimplePage() ? "ItemEditSimple.asp" : "ItmEditEntry.asp") + "?hItemCode=" + encodeURIComponent(parts[1] || "") + "&ClassCode=" + encodeURIComponent(parts[2] || "");
		} else if (response.responseText === "N") {
			alert("Item Name or Code Already Exists");
			if (frm.B2) {
				frm.B2.disabled = false;
			}
		} else {
			alert(response.responseText);
			if (frm.B2) {
				frm.B2.disabled = false;
			}
		}
		return true;
	}

	function GetStore() {
		if (isEditPage()) {
			return GetEditStore();
		}
		var frm = form();
		if (trimValue(fieldValue("hCategory")) === "") {
			alert("Select Category");
			return false;
		}
		if (!isSimplePage() && frm.selIType.selectedIndex === 0) {
			alert("Select Item Type");
			frm.selIType.focus();
			return false;
		}
		if (trimValue(fieldValue("txtitmCode")) === "") {
			alert("Enter Item Code");
			focusControl(frm.txtitmCode, true);
			return false;
		}
		if (trimValue(fieldValue("txtItmDesc")) === "") {
			alert("Enter Item Description");
			focusControl(frm.txtItmDesc, true);
			return false;
		}
		if (frm.selUoMStores.selectedIndex === 0) {
			alert("Select Unit Of Measurement");
			frm.selUoMStores.focus();
			return false;
		}
		openDialog("StorageSelectPop.asp?sUnit=" + encodeURIComponent(fieldValue("hUnitID")), xmlIsland("OutData"), "dialogHeight:375px;dialogWidth:375px;center:Yes;help:No;resizable:No;status:No", function () {
			GetStockDet();
		});
		return true;
	}

	function UoMDetails() {
		if (isEditPage()) {
			return EditUoMDetails();
		}
		var frm = form();
		var root = xmlRoot("OutData");
		var node;
		if (frm.selUoMStores.selectedIndex === 0) {
			alert("Select Unit of Measurement");
			frm.selUoMStores.focus();
			return false;
		}
		node = docFor("OutData").createElement("UOMDETAILS");
		setAttr(node, "PUR", frm.selUoMStores.value);
		setAttr(node, "MAN", "select");
		setAttr(node, "SAL", frm.selUoMStores.value);
		setAttr(node, "PURFAC", "1");
		setAttr(node, "PUROPE", "0");
		setAttr(node, "SALFAC", "1");
		setAttr(node, "SALOPE", "0");
		setAttr(node, "MANFAC", "0");
		setAttr(node, "MANOPE", "select");
		root.appendChild(node);
		openDialog("ItmOpUoMPop.asp?UOM=" + encodeURIComponent(frm.selUoMStores.value), xmlIsland("OutData"), "dialogHeight:460px;dialogWidth:500px;center:Yes;help:No;resizable:No;status:No", function () {});
		return true;
	}

	function CheckAvailability(obj, callFrom) {
		var response = syncGet("../../Common/CheckAvailabilityOfItem.asp?CallFrom=" + encodeURIComponent(callFrom) + "&Value=" + encodeURIComponent(obj.value) + "&OrgCode=" + encodeURIComponent(fieldValue("hUnitID")));
		var value = trimValue(response.responseText);
		if (value.toUpperCase() !== "NO") {
			alert(value);
			if (!isEditPage()) {
				obj.value = "";
				obj.focus();
			}
		}
	}

	function appendClone(parent, node) {
		return appendImported(parent, node && node.cloneNode ? node.cloneNode(true) : node);
	}

	function copyElementChildren(sourceRoot, targetRoot, tagName) {
		var children = childElements(sourceRoot, tagName);
		for (var i = 0; i < children.length; i += 1) {
			appendClone(targetRoot, children[i]);
		}
	}

	function clearElementChildren(root) {
		while (root && root.firstChild) {
			root.removeChild(root.firstChild);
		}
	}

	function removeLotSerialChildren(root) {
		var storages = selectNodes(root, "//STORAGE", "STORAGE");
		var lots;
		for (var i = 0; i < storages.length; i += 1) {
			lots = childElements(storages[i], "LotSerial");
			for (var jLot = 0; jLot < lots.length; jLot += 1) {
				storages[i].removeChild(lots[jLot]);
			}
		}
	}

	function appendReturnedLotData(targetRoot, returnedRoot) {
		var children;
		if (!targetRoot || !returnedRoot) {
			return;
		}
		children = childElements(returnedRoot);
		if (children.length) {
			for (var i = 0; i < children.length; i += 1) {
				appendClone(targetRoot, children[i]);
			}
		} else if (returnedRoot.nodeType === 1) {
			appendClone(targetRoot, returnedRoot);
		}
	}

	function mergeStoreDetailsInto(root) {
		var storeRoot = xmlRoot("StoreData");
		if (!root || !storeRoot) {
			return;
		}
		removeDirectChildren(root, "STOREDET");
		copyElementChildren(storeRoot, root, "STOREDET");
	}

	function appendStoreStorageNodesInto(root) {
		var storeDetails = childElements(xmlRoot("StoreData"), "STOREDET");
		var storages;
		if (!root) {
			return;
		}
		removeDirectChildren(root, "STORAGE");
		for (var i = 0; i < storeDetails.length; i += 1) {
			storages = childElements(storeDetails[i], "STORAGE");
			for (var jStorage = 0; jStorage < storages.length; jStorage += 1) {
				appendClone(root, storages[jStorage]);
			}
		}
	}

	function editAttributeIds(headTwoOnly) {
		var attrs = childElements(xmlRoot("ItemAttData"), "Attribute");
		var ids = [];
		for (var i = 0; i < attrs.length; i += 1) {
			if (!headTwoOnly || getAttr(attrs[i], "Head") === "2") {
				ids.push(getAttr(attrs[i], "AttID"));
			}
		}
		return ids.length ? ids.join(",") : "NULL";
	}

	function updateEditTaxSelections() {
		setSelectValue(field("selPurTaxType"), fieldValue("hPurTaxType"));
		setSelectValue(field("selSalTaxType"), fieldValue("hSalTaxType"));
	}

	function EditInit() {
		var frm = form();
		var root = xmlRoot("OutData");
		var controls = childElements(root, "CONTROLS");
		var accounts;
		var account;
		var effDate = field("ctlEffDate") || window.ctlEffDate;
		for (var i = 0; i < controls.length; i += 1) {
			accounts = childElements(controls[i], "ACCHEAD");
			for (var jAcc = 0; jAcc < accounts.length; jAcc += 1) {
				account = accounts[jAcc];
				if (trimValue(getAttr(account, "Type")) === "O") {
					frm.hOAH.value = getAttr(account, "Value");
					setText("spanOpenAccHead", getAttr(account, "Name") || " ");
				} else if (trimValue(getAttr(account, "Type")) === "C") {
					frm.hCAH.value = getAttr(account, "Value");
					setText("spanCloseAccHead", getAttr(account, "Name") || " ");
				}
			}
		}
		if (field("hEffFrom") && trimValue(fieldValue("hEffFrom")) !== "" && effDate) {
			if (effDate.setDate) {
				effDate.setDate(fieldValue("hEffFrom"));
			} else if (effDate.SetDate) {
				effDate.SetDate(fieldValue("hEffFrom"));
			} else {
				effDate.value = fieldValue("hEffFrom");
			}
		}
		if (frm.hItemTypeCode && frm.selIType) {
			setSelectValue(frm.selIType, frm.hItemTypeCode.value);
		}
		if (window.LetIType && frm.selIType) {
			window.LetIType(frm.selIType);
		}
		if (window.ChangeLabel && frm.selIType) {
			window.ChangeLabel(frm.selIType);
		}
		return true;
	}

	function EditUploadImage() {
		var itemCode = fieldValue("hItemCode");
		openDialog("ItmUploadImage.asp?ItemCode=" + encodeURIComponent(itemCode), "", "dialogHeight:120px;dialogWidth:275px;Status:No", function () {
			form().submit();
		});
		return true;
	}

	function GetEditAttr() {
		var root = xmlRoot("AttrData");
		var select = field("selAttr");
		var itemType = selectedValue(field("selIType"));
		var nodes = selectNodes(root, "//ROOT/ATTRIBUTES [@ITEMTYPEID = '" + itemType + "']", "ATTRIBUTES");
		var option;
		if (!select) {
			return false;
		}
		select.length = 0;
		for (var i = 0; i < nodes.length; i += 1) {
			if (getAttr(nodes[i], "ITEMTYPEID") !== itemType) {
				continue;
			}
			option = document.createElement("option");
			option.text = getAttr(nodes[i], "ATTRNAME");
			option.value = getAttr(nodes[i], "ATTRID");
			select.add(option);
		}
		return true;
	}

	function GetEditStockDet() {
		var frm = form();
		var root = xmlRoot("StoreData");
		var table = byId("tblData");
		var storageNodes = selectNodes(root, "//STOREDET", "STOREDET");
		var storage;
		var unitStore;
		var orgId;
		var loc;
		var bin;
		var qty;
		var value;
		var monthYear;
		var storeName;
		var row;
		var cell;
		var button;
		var displayIndex = 0;
		ClearTable();
		if (!table) {
			return false;
		}
		for (var i = 0; i < storageNodes.length; i += 1) {
			storage = storageNodes[i];
			unitStore = trimValue(getAttr(storage, "UNITSTORE")).split("-");
			orgId = unitStore[0] || "";
			loc = unitStore[1] || "";
			bin = unitStore[2] || "";
			qty = unitStore[3] || "";
			value = unitStore[4] || "";
			monthYear = unitStore[5] || "";
			storeName = getAttr(storage, "STORE");
			displayIndex += 1;
			if (displayIndex === 1) {
				setText("idStore", storeName);
			}
			row = table.insertRow(state.j + 1);
			appendCell(row, "ExcelSerial", "center", displayIndex);
			appendCell(row, "ExcelDisplayCell", "left", storeName);
			cell = appendCell(row, "ExcelInputCell", "", createTextInput("txtMYX" + loc + "X" + bin, monthYear, 10, isSimplePage() ? "Formelem" : "FormElem", "DoKeyPress('N',6,0)"));
			cell.width = "8";
			cell = appendCell(row, "ExcelInputCell", "right", createTextInput("txtQTYX" + loc + "X" + bin, qty, 12, isSimplePage() ? "Formelem" : "FormElem", "DoKeyPress('Y',7,3)"));
			cell.width = "10";
			cell = appendCell(row, "ExcelInputCell", "right", createTextInput("txtVALX" + loc + "X" + bin, value, 12, isSimplePage() ? "Formelem" : "FormElem", "DoKeyPress('Y',10,2)"));
			cell.width = "10";
			button = document.createElement("input");
			button.type = "button";
			button.value = " Yes ";
			button.name = "btn:" + (fieldValue("hClassCode") || state.iClass || "0") + ":" + orgId + ":" + loc + ":" + bin;
			button.size = 12;
			button.maxLength = 10;
			button.className = "AddButtonX";
			button.onclick = (function (name) {
				return function () {
					if (fieldValue("hReplicateItem") === "Y") {
						GetLotReplicateItem(this, name);
					} else {
						GetLot(this, name);
					}
				};
			}(storeName));
			appendCell(row, "ExcelFieldCell", "center", button);
			state.j += 1;
		}
		updateEditTaxSelections();
		DisplayAttTable();
		return true;
	}

	function editLotValues(obj) {
		var parts = String(obj && obj.name || "").split(":");
		var loc = parts[3] || "";
		var bin = parts[4] || "";
		return {
			classCode: parts[1] || fieldValue("hClassCode") || "0",
			orgId: parts[2] || "",
			loc: loc,
			bin: bin,
			qty: field("txtQTYX" + loc + "X" + bin),
			value: field("txtVALX" + loc + "X" + bin)
		};
	}

	function validateEditLotControls(values) {
		var frm = form();
		if (frm.selRecNum.selectedIndex === 0) {
			alert("Select Receipt Numbering");
			frm.selRecNum.focus();
			return false;
		}
		if (!values.qty || trimValue(values.qty.value) === "") {
			alert("Enter Quantity");
			focusControl(values.qty, true);
			return false;
		}
		if (!values.value || trimValue(values.value.value) === "") {
			alert("Enter Value");
			focusControl(values.value, true);
			return false;
		}
		return updateAttributeInputValues();
	}

	function GetEditLot(obj, storeName) {
		var frm = form();
		var values = editLotValues(obj);
		var attrList;
		var itemCode;
		var tempValues;
		var url = "";
		if (!validateEditLotControls(values)) {
			return false;
		}
		if (num(values.qty.value) === 0) {
			return true;
		}
		attrList = editAttributeIds(false);
		itemCode = fieldValue("hItemCode");
		tempValues = [
			frm.selRecNum.value,
			itemCode,
			values.classCode,
			values.orgId,
			values.loc,
			values.bin,
			values.qty.value,
			String(storeName || "").replace(/&/g, "and"),
			frm.selUoMStores.value,
			values.value.value,
			trimValue(fieldValue("txtItmDesc")).replace(/'/g, "~~"),
			frm.selIType.value,
			attrList
		].join("``");
		if (frm.selRecNum.value === "N" && attrList !== "" && attrList !== "NULL") {
			url = "../../Common/PackingLotSerialEditDetailsForNone.asp?sTemp=" + encodeURIComponent(tempValues);
		} else if (frm.selRecNum.value !== "N") {
			url = "../../Common/PackingLotSerialEditDetails.asp?sTemp=" + encodeURIComponent(tempValues);
		}
		if (!url) {
			return true;
		}
		openDialog(url, xmlIsland("StoreData"), "dialogHeight:580px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No", function (returnedValue) {
			var returnedRoot = xmlRoot(returnedValue);
			var targetRoot = fieldValue("hReplicateItem") === "Y" ? xmlRoot("ReplicateData") : xmlRoot("OutData");
			updateReturnedStorageFields(returnedRoot || xmlRoot("StoreData"));
			appendReturnedLotData(targetRoot, returnedRoot);
			syncPost("XMLSave.asp?SessionFlag=False&Value=LotData&Folder=Master", serializeXml("StoreData"));
		});
		return true;
	}

	function GetLotReplicateItem(obj, storeName) {
		var frm = form();
		var values = editLotValues(obj);
		var attrList;
		var tempValues;
		var url = "";
		var root = xmlRoot("ReplicateData");
		if (!validateEditLotControls(values)) {
			return false;
		}
		if (num(values.qty.value) === 0) {
			return true;
		}
		updateStorageXml(root);
		removeLotSerialChildren(root);
		attrList = editAttributeIds(false);
		if (frm.selIType.value === "GAR") {
			tempValues = [fieldValue("txtitmCode"), values.classCode, values.orgId, values.loc, values.bin, values.qty.value, storeName, frm.selUoMStores.value, values.value.value, trimValue(fieldValue("txtItmDesc")), frm.selIType.value, attrList].join("``");
			url = "CatalogItemLotPopEntry.asp?sTemp=" + encodeURIComponent(tempValues);
		} else if (frm.selRecNum.value === "N" && frm.selIType.value === "STO") {
			tempValues = [fieldValue("txtitmCode"), values.classCode, values.orgId, values.loc, values.bin, values.qty.value, storeName, frm.selUoMStores.value, values.value.value, trimValue(fieldValue("txtItmDesc")), frm.selIType.value, attrList].join("``");
			url = "CatalogItemNonePopEntry.asp?sTemp=" + encodeURIComponent(tempValues);
		} else if (frm.selRecNum.value !== "N") {
			tempValues = [frm.selRecNum.value, "0", values.classCode, values.orgId, values.loc, values.bin, values.qty.value, storeName, frm.selUoMStores.value, values.value.value, trimValue(fieldValue("txtItmDesc")).replace(/'/g, "~~"), frm.selIType.value, attrList].join("``");
			url = "stockLotSerPopNew.asp?sTemp=" + encodeURIComponent(tempValues);
		}
		if (!url) {
			return true;
		}
		openDialog(url, xmlIsland("ReplicateData"), "dialogHeight:580px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No", function (returnedValue) {
			appendReturnedLotData(root, xmlRoot(returnedValue));
			syncPost("XMLSave.asp?SessionFlag=False&Value=ReplicateData&Folder=Master", serializeXml("ReplicateData"));
		});
		return true;
	}

	function appendEditDetailsAndControls(root, attrList, replicate) {
		var frm = form();
		var doc = xmlDocumentFromRoot(root) || docFor(replicate ? "ReplicateData" : "OutData");
		var detail = doc.createElement("DETAILS");
		var controls = doc.createElement("CONTROLS");
		var account;
		setAttr(detail, "ITYPE", fieldValue("selIType"));
		setAttr(detail, "ICODE", replicate ? trimValue(fieldValue("txtitmCode")) : trimValue(fieldValue("hItemCode")));
		if (!replicate) {
			setAttr(detail, "COMPITEMCODE", trimValue(fieldValue("txtitmCode")));
		}
		setAttr(detail, "DESC", trimValue(fieldValue("txtItmDesc")));
		setAttr(detail, "SHDESC", "");
		setAttr(detail, "CATALOUGE", "");
		setAttr(detail, "DRAWVER", "");
		setAttr(detail, "VARIANT", trimValue(fieldValue("txtVariant")));
		setAttr(detail, "ADDDESC", trimValue(fieldValue("txtItmAddDesc")));
		setAttr(detail, "UOM", fieldValue("selUoMStores"));
		setAttr(detail, "CATEGORY", fieldValue("hCategory"));
		setAttr(detail, "ATTRIBUTES", attrList);
		setAttr(detail, "GROUP", fieldValue("hGroup"));
		setAttr(detail, "LEVEL", fieldValue("hLevel"));
		setAttr(detail, "UNIT", fieldValue("selUnit"));
		if (replicate) {
			setAttr(detail, "OPSTOCKUNIT", fieldValue("selUnit"));
		} else {
			setAttr(detail, "ITEMSTATUS", checkedRadioValue("radStatus"));
		}
		setAttr(detail, "MODVAT", checkedRadioValue("radMod"));
		setAttr(detail, "PURTAX", selectedValue(frm.selPurTaxType));
		setAttr(detail, "SALTAX", selectedValue(frm.selSalTaxType));
		root.appendChild(detail);

		setAttr(controls, "RECNUM", fieldValue("selRecNum"));
		setAttr(controls, "ROUTING", fieldValue("selRecRout"));
		setAttr(controls, "ACCOUNTING", fieldValue("selAcc"));
		setAttr(controls, "MODVAT", "0");
		setAttr(controls, "REORDERLEVEL", isSimplePage() ? fieldValue("txtReLvl") : "");
		setAttr(controls, "REORDERQTY", isSimplePage() ? fieldValue("txtReQty") : "");
		setAttr(controls, "ECOORDERQTY", isSimplePage() ? fieldValue("txtEcQty") : "");
		if (radio("radBoM", 1) && radio("radBoM", 1).checked) {
			setAttr(controls, "BOMAPPLICABLE", radio("radBoM", 1).value);
		} else if (radio("radBoM", 0) && radio("radBoM", 0).checked) {
			setAttr(controls, "BOMAPPLICABLE", radio("radBoM", 0).value);
		}
		root.appendChild(controls);

		account = doc.createElement("ACCHEAD");
		setAttr(account, "Name", getText("spanOpenAccHead"));
		setAttr(account, "Type", "O");
		setAttr(account, "Value", fieldValue("hOAH"));
		controls.appendChild(account);

		account = doc.createElement("ACCHEAD");
		setAttr(account, "Name", getText("spanCloseAccHead"));
		setAttr(account, "Type", "C");
		setAttr(account, "Value", fieldValue("hCAH"));
		controls.appendChild(account);
	}

	function prepareEditPayload(rootName, replicate) {
		var root = xmlRoot(rootName);
		var attrList;
		if (!root || !updateAttributeInputValues()) {
			return null;
		}
		removeDirectChildren(root, "DETAILS");
		removeDirectChildren(root, "CONTROLS");
		removeDirectChildren(root, "ATTRIBUTE");
		removeDirectChildren(root, "PRICING");
		attrList = editAttributeIds(!replicate);
		appendAttributeData(root);
		appendEditDetailsAndControls(root, attrList, replicate);
		mergeStoreDetailsInto(root);
		if (replicate) {
			appendStoreStorageNodesInto(root);
		}
		updateStorageXml(root);
		appendPricing(root);
		return root;
	}

	function validateEditForm(requireClass, skipStoreCheck) {
		var frm = form();
		if (trimValue(fieldValue("hCategory")) === "") {
			alert("Select Category");
			return false;
		}
		if (!isSimplePage() && frm.selIType.selectedIndex === 0) {
			alert("Select Item Type");
			frm.selIType.focus();
			return false;
		}
		if (requireClass && trimValue(fieldValue("txtClass")) === "") {
			alert("Select Classification");
			return false;
		}
		if (trimValue(fieldValue("txtitmCode")) === "") {
			alert("Enter Item Code");
			focusControl(frm.txtitmCode, true);
			return false;
		}
		if (trimValue(fieldValue("txtItmDesc")) === "") {
			alert("Enter Item Description");
			focusControl(frm.txtItmDesc, true);
			return false;
		}
		if (frm.selUoMStores.selectedIndex === 0) {
			alert("Select Unit Of Measurement");
			frm.selUoMStores.focus();
			return false;
		}
		if (frm.selRecNum.value === "select") {
			alert("Select Receipt Numbering");
			frm.selRecNum.focus();
			return false;
		}
		if (frm.selRecRout.value === "select") {
			alert("Select Receipt Routing");
			frm.selRecRout.focus();
			return false;
		}
		if (frm.selAcc.value === "select") {
			alert("Select Accounting Type");
			frm.selAcc.focus();
			return false;
		}
		if (!skipStoreCheck && getText("idStore") === "") {
			alert("Select Storage");
			GetStore();
			return false;
		}
		return true;
	}

	function CheckEditSubmitDetails(finFrom, finTo) {
		var frm = form();
		var root;
		var response;
		if (fieldValue("hReplicateItem") === "Y") {
			return ReplicateItemSave(finFrom, finTo);
		}
		if (!validateEditForm(false)) {
			return false;
		}
		if (frm.selUnit) {
			frm.selUnit.value = fieldValue("hUnitID");
		}
		if (state.j > 0 && !validateStorageRows(finFrom, finTo)) {
			return false;
		}
		if (frm.hRecNum) {
			frm.hRecNum.value = frm.selRecNum.value;
		}
		root = prepareEditPayload("OutData", false);
		if (!root || !storageHasLotSerial(root)) {
			return false;
		}
		if (frm.B2) {
			frm.B2.disabled = true;
		}
		response = syncPost("ItmEditInsert.asp", serializeXml("OutData"));
		if (trimValue(response.responseText) === "") {
			alert("Item has been Amend Successfully.");
			window.location.href = "ItemListEntry.asp?ACTN=L:EDIT=S";
		} else if (response.responseText === "N") {
			alert("Item Name or Code Already Exists");
			if (frm.B2) {
				frm.B2.disabled = false;
			}
		} else {
			alert(response.responseText);
			if (frm.B2) {
				frm.B2.disabled = false;
			}
		}
		return true;
	}

	function clearStorageChildNodes(root) {
		var storages = selectNodes(root, "//STORAGE", "STORAGE");
		for (var i = 0; i < storages.length; i += 1) {
			clearElementChildren(storages[i]);
		}
	}

	function ReplicateItemSave(finFrom, finTo) {
		var frm = form();
		var root;
		var response;
		var parts;
		if (!validateEditForm(true)) {
			return false;
		}
		if (frm.selUnit) {
			frm.selUnit.value = fieldValue("hUnitID");
		}
		if (state.j > 0 && !validateStorageRows(finFrom, finTo)) {
			return false;
		}
		if (frm.hRecNum) {
			frm.hRecNum.value = frm.selRecNum.value;
		}
		root = prepareEditPayload("ReplicateData", true);
		if (!root) {
			return false;
		}
		if (trimValue(fieldValue("hRecNum")) === "N") {
			clearStorageChildNodes(root);
		}
		if (frm.B2) {
			frm.B2.disabled = true;
		}
		response = syncPost("ItmCreationDefinitionInsert.asp", serializeXml("ReplicateData"));
		if (response.responseText.substring(0, 13) === "ItemClassCode") {
			alert("Item has been created and defined Successfully.");
			parts = response.responseText.split(":");
			window.location.href = (isSimplePage() ? "ItemEditSimple.asp" : "ItmEditEntry.asp") + "?hItemCode=" + encodeURIComponent(parts[1] || "") + "&ClassCode=" + encodeURIComponent(parts[2] || "");
		} else if (response.responseText === "N") {
			alert("Item Name or Code Already Exists");
			if (frm.B2) {
				frm.B2.disabled = false;
			}
		} else {
			alert(response.responseText);
			if (frm.B2) {
				frm.B2.disabled = false;
			}
		}
		return true;
	}

	function SaveReplicateXML() {
		syncPost("XMLSave.asp?SessionFlag=False&Value=ReplicateData&Folder=Master", serializeXml("ReplicateData"));
		return true;
	}

	function Replicate_Item() {
		var frm = form();
		var storeRoot = xmlRoot("StoreData");
		var unit = fieldValue("hUnitID");
		if (storeRoot) {
			clearElementChildren(storeRoot);
		}
		openDialog("ItmEditEntry_ReplicateItem.asp?sUnit=" + encodeURIComponent(unit), xmlIsland("StoreData"), "dialogHeight:375px;dialogWidth:375px;center:Yes;help:No;resizable:No;status:No", function () {
			var nodes = selectNodes(xmlRoot("StoreData"), "//STOREDET", "STOREDET");
			var parts;
			var orgId = "";
			for (var i = 0; i < nodes.length; i += 1) {
				parts = trimValue(getAttr(nodes[i], "UNITSTORE")).split("-");
				orgId = parts[0] || orgId;
				if (orgId) {
					frm.hUnitID.value = orgId;
					break;
				}
			}
			if (trimValue(orgId) !== "") {
				if (frm.But_Replicate) {
					frm.But_Replicate.disabled = true;
				}
				if (frm.hReplicateItem) {
					frm.hReplicateItem.value = "Y";
				}
			}
			GetEditStockDet();
			if (frm.selUnit) {
				frm.selUnit.value = fieldValue("hUnitID");
			}
			prepareEditPayload("ReplicateData", true);
		});
		return true;
	}

	function GetEditStore() {
		var frm = form();
		if (!validateEditForm(false, true)) {
			return false;
		}
		openDialog("StorageSelectPopEdit.asp?sUnit=" + encodeURIComponent(fieldValue("hUnitID")), xmlIsland("StoreData"), "dialogHeight:375px;dialogWidth:375px;center:Yes;help:No;resizable:No;status:No", function () {
			GetEditStockDet();
		});
		return true;
	}

	function EditUoMDetails() {
		var frm = form();
		if (frm.selUoMStores.selectedIndex === 0) {
			alert("Select Unit of Measurement");
			frm.selUoMStores.focus();
			return false;
		}
		openDialog("ItmEditOpUoMEntry.asp?UOM=" + encodeURIComponent(frm.selUoMStores.value), xmlIsland("OutData"), "dialogHeight:460px;dialogWidth:500px;center:Yes;help:No;resizable:No;status:No", function () {});
		return true;
	}

	function EditItemSpecPop() {
		var frm = form();
		if (trimValue(fieldValue("txtClass")) === "") {
			alert("Select the Category");
			return false;
		}
		openDialog("ItmSpecsEditPop.asp?ItemType=" + encodeURIComponent(selectedValue(frm.selIType)) + "&ClassCode=" + encodeURIComponent(fieldValue("hClassCode")) + "&ItemCode=" + encodeURIComponent(fieldValue("hItemCode")), xmlIsland("OutData"), "dialogHeight:300px;dialogWeight:250px;Status:No", function () {});
		return true;
	}

	function EditClass() {
		return popClass();
	}

	function GetClass() {
		return popClass();
	}

	window.AssaignValue = AssaignValue;
	window.CalcValue = CalcValue;
	window.ItemSpecPop = ItemSpecPop;
	window.ChangeBOM = ChangeBOM;
	window.ChangeAttribute = ChangeAttribute;
	window.UploadImage = UploadImage;
	window.DeleteImage = DeleteImage;
	window.Init = Init;
	window.CheckNoSeries = CheckNoSeries;
	window.SelectAccHead = SelectAccHead;
	window.ManageAttribute = ManageAttribute;
	window.ClearAttTable = ClearAttTable;
	window.DisplayAttTable = DisplayAttTable;
	window.ClearTable = ClearTable;
	window.popClass = popClass;
	window.EditClass = EditClass;
	window.GetClass = GetClass;
	window.GetDetails = GetDetails;
	window.GetAttr = GetAttr;
	window.GetStockDet = GetStockDet;
	window.GetLot = GetLot;
	window.GetLotReplicateItem = GetLotReplicateItem;
	window.SaveReplicateXML = SaveReplicateXML;
	window.ReplicateItemSave = ReplicateItemSave;
	window.Replicate_Item = Replicate_Item;
	window.CheckSubmitDetails = CheckSubmitDetails;
	window.GetStore = GetStore;
	window.UoMDetails = UoMDetails;
	window.CheckAvailability = CheckAvailability;
}(window, document));
