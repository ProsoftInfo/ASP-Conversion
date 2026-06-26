(function (window, document) {
	"use strict";

	var config = window.__itmsPartySelectorConfig || {};
	var buttonPressed = "";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.formname || document.forms.formname || document.forms.FormName || document.forms[0] || null;
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] || null : null;
	}

	function fieldValue(name) {
		var item = field(name);
		return item ? item.value : "";
	}

	function setFieldValue(name, value) {
		var item = field(name);
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function byId(id) {
		return document.getElementById(id) || document.getElementsByName(id)[0] || window[id] || null;
	}

	function setText(id, value) {
		var element = byId(id);
		if (element) {
			if ("innerText" in element) {
				element.innerText = value == null ? "" : String(value);
			} else {
				element.textContent = value == null ? "" : String(value);
			}
		}
	}

	function asArray(collection) {
		return Array.prototype.slice.call(collection || []);
	}

	function childElements(node, name) {
		var wanted = name && String(name).toLowerCase();
		return asArray(node && node.childNodes).filter(function (child) {
			return child.nodeType === 1 && (!wanted || String(child.nodeName).toLowerCase() === wanted);
		});
	}

	function attr(node, name) {
		return node && node.getAttribute ? node.getAttribute(name) || "" : "";
	}

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function xmlObject(name) {
		ensureCompat();
		return window[name] || document[name] || byId(name) || null;
	}

	function xmlDocument(nameOrObject) {
		var object = typeof nameOrObject === "string" ? xmlObject(nameOrObject) : nameOrObject;
		return object && object.XMLDocument || object && object._doc || object || null;
	}

	function xmlRoot(nameOrObject) {
		var object = typeof nameOrObject === "string" ? xmlObject(nameOrObject) : nameOrObject;
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || null;
	}

	function createXmlElement(xmlName, nodeName) {
		var doc = xmlDocument(xmlName);
		if (doc && doc.createElement) {
			return doc.createElement(nodeName);
		}
		return document.implementation.createDocument("", "", null).createElement(nodeName);
	}

	function loadXml(xmlName, xmlText) {
		var object = xmlObject(xmlName);
		if (object && object.loadXML) {
			object.loadXML(xmlText || "<Root/>");
			return object;
		}
		if (object) {
			object._doc = new DOMParser().parseFromString(xmlText || "<Root/>", "text/xml");
		}
		return object;
	}

	function createHttp() {
		if (window.CreateObject) {
			try {
				return window.CreateObject("MSXML2.XMLHTTP");
			} catch (ignore) {}
		}
		return new XMLHttpRequest();
	}

	function getXml(url) {
		var xhr = createHttp();
		xhr.open("GET", url, false);
		xhr.send(null);
		if (xhr.responseXML && xhr.responseXML.documentElement) {
			return new XMLSerializer().serializeToString(xhr.responseXML);
		}
		if (trim(xhr.responseText)) {
			return xhr.responseText;
		}
		return "";
	}

	function dialogId() {
		var match = String(window.location.search || "").match(/[?&]__itmsDialogId=([^&]+)/);
		return match ? decodeURIComponent(match[1]) : "";
	}

	function returnValue(value) {
		var id;
		window.returnValue = value;
		window.returnvalue = value;
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(value);
			return;
		}
		id = dialogId();
		if (id && window.opener && window.opener.ITMSModernCompat && window.opener.ITMSModernCompat._receiveDialogValue) {
			window.opener.ITMSModernCompat._receiveDialogValue(id, value);
		}
	}

	function closeWithPartyData() {
		returnValue(xmlRoot("PartyData"));
		window.close();
	}

	function isMultiSelect() {
		return trim(fieldValue("hSelectMode")).toUpperCase() === "M";
	}

	function table() {
		return byId("tblItem");
	}

	function clearTable() {
		var tbl = table();
		if (!tbl || !tbl.rows) {
			return;
		}
		while (tbl.rows.length > 5) {
			tbl.deleteRow(2);
		}
	}

	function addCell(row, className, align, text) {
		var cell = row.insertCell(-1);
		if (className) {
			cell.className = className;
		}
		if (align) {
			cell.align = align;
		}
		if (text != null) {
			cell.innerHTML = String(text);
		}
		return cell;
	}

	function insertDataRow() {
		var tbl = table();
		return tbl.insertRow(Math.max(2, tbl.rows.length - 3));
	}

	function setControlTabs() {
		var index = 6;
		var currentPage = field("txtCurrPage");
		var addButton = field("btnAddToList");
		var doneButton = field("btnDone");
		if (currentPage) {
			currentPage.tabIndex = index;
		}
		index += 1;
		if (addButton) {
			addButton.tabIndex = index;
		}
		index += 1;
		if (doneButton) {
			doneButton.tabIndex = index;
		}
	}

	function tempRoot() {
		return xmlRoot("TempItem");
	}

	function partyRoot() {
		return xmlRoot("PartyData");
	}

	function selectedEntries() {
		return childElements(partyRoot(), "Entry");
	}

	function removeAllEntries() {
		var root = partyRoot();
		selectedEntries().forEach(function (entry) {
			root.removeChild(entry);
		});
	}

	function findLoadedNodeByCode(code) {
		var wanted = trim(code);
		var root = tempRoot();
		var nodeName = config.kind === "employee" ? "Emp" : "Party";
		var attrName = config.kind === "employee" ? "EmpID" : "PartyCode";
		var found = null;
		childElements(root, nodeName).forEach(function (node) {
			if (!found && trim(attr(node, attrName)) === wanted) {
				found = node;
			}
		});
		return found;
	}

	function entryExists(code) {
		var wanted = trim(code);
		return selectedEntries().some(function (entry) {
			return trim(attr(entry, "RetField1")) === wanted;
		});
	}

	function removeEntry(code) {
		var root = partyRoot();
		var wanted = trim(code);
		selectedEntries().forEach(function (entry) {
			if (trim(attr(entry, "RetField1")) === wanted) {
				root.removeChild(entry);
			}
		});
	}

	function selectedSubtypeFor(partyCode) {
		var typeFlag = fieldValue("hPartyTypeZ" + partyCode);
		var select = field("SelPartyTypeZ" + partyCode);
		if (!trim(typeFlag) || trim(typeFlag).toUpperCase() === "N" || !select || select.selectedIndex < 0) {
			return ["", ""];
		}
		return String(select.options[select.selectedIndex].value || "").split("|");
	}

	function addEntryFromParts(parts, checked) {
		var code = parts[1] || "";
		var rowNode;
		var entry;
		var subtype;
		if (!checked) {
			removeEntry(code);
			return;
		}
		if (entryExists(code)) {
			return;
		}
		rowNode = findLoadedNodeByCode(code);
		if (!rowNode) {
			return;
		}
		entry = createXmlElement("PartyData", "Entry");
		if (config.kind === "employee") {
			setAttr(entry, "RetField0", attr(rowNode, "EmpFullName"));
			setAttr(entry, "RetField1", attr(rowNode, "EmpID"));
			setAttr(entry, "RetField2", attr(rowNode, "EmpCode"));
			setAttr(entry, "RetField3", attr(rowNode, "EmpName"));
		} else {
			subtype = selectedSubtypeFor(code);
			setAttr(entry, "RetField0", attr(rowNode, "PartyName"));
			setAttr(entry, "RetField1", attr(rowNode, "PartyCode"));
			setAttr(entry, "RetField2", attr(rowNode, "OrgnPartyCode"));
			setAttr(entry, "RetField3", subtype[0] || "");
			setAttr(entry, "RetField4", subtype[1] || "");
			setAttr(entry, "RetField5", "0");
		}
		partyRoot().appendChild(entry);
	}

	function xmlFun(obj) {
		var parts = String(obj && obj.value || "").split(":");
		if (isMultiSelect()) {
			addEntryFromParts(parts, !!(obj && obj.checked));
		} else {
			removeAllEntries();
			if (obj && obj.checked) {
				addEntryFromParts(parts, true);
			}
		}
		displayList();
	}

	function addFun() {
		var count = parseInt(fieldValue("hChkCount"), 10) || 0;
		var check;
		for (var i = 1; i <= count; i += 1) {
			check = field("ChkZ" + i);
			if (check && check.checked) {
				xmlFun(check);
				check.checked = false;
			}
		}
	}

	function escapeHtml(value) {
		return String(value == null ? "" : value)
			.replace(/&/g, "&amp;")
			.replace(/</g, "&lt;")
			.replace(/>/g, "&gt;")
			.replace(/"/g, "&quot;")
			.replace(/'/g, "&#39;");
	}

	function displayList() {
		var host = byId("idSelList");
		var html = '<br><TABLE class="TableOutLineOnly" cellspacing="1" width="100%">';
		selectedEntries().forEach(function (entry) {
			html += '<tr><td class="ExcelDisplayCell">';
			html += '<input type="checkbox" name="chk" value="' + escapeHtml(attr(entry, "RetField1")) + '" checked onclick="RemoveNode(this)">';
			html += '</td><td class="ExcelDisplayCell">' + escapeHtml(attr(entry, "RetField2")) + '</td>';
			html += '<td class="ExcelDisplayCell">' + escapeHtml(attr(entry, "RetField0").replace(/~~/g, '"')) + '</td></tr>';
		});
		html += "</table><br>";
		if (host) {
			host.innerHTML = html;
		}
	}

	function removeNode(obj) {
		var code = obj && obj.value || "";
		if (obj && obj.checked) {
			return;
		}
		removeEntry(code);
		asArray(document.querySelectorAll('input[type="checkbox"][name^="ChkZ"]')).forEach(function (check) {
			var parts = String(check.value || "").split(":");
			if (trim(parts[1]) === trim(code)) {
				check.checked = false;
			}
		});
		displayList();
	}

	function buildRequest(baseArguments) {
		var request = String(baseArguments || "");
		if (config.kind === "employee") {
			request += "&EmpCode=" + encodeURIComponent(fieldValue("txtEmpCode"));
			request += "&EmpName=" + encodeURIComponent(fieldValue("txtEmpName"));
		} else {
			request += "&OrgnPartyCode=" + encodeURIComponent(fieldValue("txtOrgnPartyCode"));
			request += "&PartyName=" + encodeURIComponent(fieldValue("txtPartyName"));
			request += "&ParType=" + encodeURIComponent(fieldValue("selParType"));
		}
		return request;
	}

	function loadData(query) {
		var xml = getXml(config.dataUrl + "?" + query);
		if (trim(xml)) {
			loadXml("TempItem", xml);
		}
	}

	function populateSubtypeSelect(cell, partyCode) {
		var orgCode = fieldValue("hOrgID");
		var xml = getXml("PartySubType.asp?ParCode=" + encodeURIComponent(partyCode) + "&OrgCode=" + encodeURIComponent(orgCode));
		var subtypeRoot;
		var select;
		var hidden;
		if (trim(xml)) {
			loadXml("XMLPartySubType", xml);
		}
		subtypeRoot = xmlRoot("XMLPartySubType");
		if (childElements(subtypeRoot).length) {
			select = document.createElement("select");
			select.name = "SelPartyTypeZ" + partyCode;
			select.className = "FormElem";
			childElements(subtypeRoot, "Party").forEach(function (node) {
				var option = document.createElement("option");
				option.value = attr(node, "SubType");
				option.text = node.textContent || "";
				select.appendChild(option);
			});
			cell.appendChild(select);
			hidden = document.createElement("input");
			hidden.type = "hidden";
			hidden.name = "hPartyTypeZ" + partyCode;
			hidden.value = "Y";
			cell.appendChild(hidden);
		} else {
			cell.innerHTML = "N/A";
			hidden = document.createElement("input");
			hidden.type = "hidden";
			hidden.name = "hPartyTypeZ" + partyCode;
			hidden.value = "N";
			cell.appendChild(hidden);
		}
	}

	function renderRows() {
		var root = tempRoot();
		var nodes = childElements(root, config.kind === "employee" ? "Emp" : "Party");
		var currentPage = attr(root, "CurrPage") || "0";
		var totalPage = attr(root, "TotPage") || "0";
		var lastSerial = 0;
		clearTable();
		if (!nodes.length) {
			var emptyRow = insertDataRow();
			var emptyCell = addCell(emptyRow, "ExcelDisplayCell", "center", "No Records Found");
			emptyCell.colSpan = config.kind === "employee" ? 3 : 4;
			currentPage = "0";
			totalPage = "0";
		} else {
			nodes.forEach(function (node) {
				var serial = attr(node, config.kind === "employee" ? "SNo" : "SNO");
				var counter = attr(node, "Counter");
				var code = attr(node, config.kind === "employee" ? "EmpCode" : "OrgnPartyCode");
				var id = attr(node, config.kind === "employee" ? "EmpID" : "PartyCode");
				var name = attr(node, config.kind === "employee" ? "EmpFullName" : "PartyName");
				var row = insertDataRow();
				var selectCell = addCell(row, "ExcelDisplayCell", "center");
				var input = document.createElement("input");
				input.type = isMultiSelect() ? "checkbox" : "radio";
				input.name = "ChkZ" + (isMultiSelect() ? serial : counter);
				input.value = code + ":" + id;
				if (!isMultiSelect()) {
					input.onclick = function () {
						xmlFun(input);
					};
				}
				selectCell.appendChild(input);
				addCell(row, "ExcelDisplayCell", "Left", escapeHtml(code));
				addCell(row, "ExcelDisplayCell", "Left", escapeHtml(name));
				if (config.kind === "party") {
					populateSubtypeSelect(addCell(row, "ExcelDisplayCell", "Left"), id);
				}
				lastSerial = parseInt(serial, 10) || lastSerial;
			});
		}
		setFieldValue("hChkCount", lastSerial);
		setFieldValue("txtCurrPage", currentPage);
		setFieldValue("hPage", currentPage);
		setText("spanTotPage", totalPage);
		setControlTabs();
	}

	function showPage(argumentsText) {
		var request = buildRequest(argumentsText);
		buttonPressed = "Page";
		setAttr(partyRoot(), "Action", "Page");
		setAttr(partyRoot(), "PassQuery", request);
		loadData(request);
		renderRows();
		displayList();
	}

	function init() {
		loadData(fieldValue("hTemp"));
		renderRows();
		displayList();
	}

	function sendValue() {
		buttonPressed = "Done";
		setAttr(partyRoot(), "Action", "Done");
		closeWithPartyData();
	}

	function callSearchMain(eventArg) {
		var eventObj = eventArg || window.event || {};
		var key = eventObj.keyCode || eventObj.which;
		var request = fieldValue("hRequest");
		var page = parseFloat(fieldValue("hPage")) || 0;
		var lastPage = parseFloat((byId("spanTotPage") || {}).innerText || (byId("spanTotPage") || {}).textContent || "0") || 0;
		if (key === 13) {
			showPage(request + "&Page=" + fieldValue("txtCurrPage"));
		} else if (key === 33) {
			if (page > 1) {
				page -= 1;
			}
			showPage(request + "&Page=" + page);
		} else if (key === 34) {
			if (page < lastPage) {
				page += 1;
			}
			showPage(request + "&Page=" + page);
		}
	}

	function callSearch() {
		showPage(fieldValue("hRequest") + "&Page=" + fieldValue("txtCurrPage"));
	}

	function displaySubType() {
		var select = field("selParType");
		var xml;
		var subtypeRoot;
		if (!select) {
			return;
		}
		xml = getXml("PartySubType.asp?OrgCode=" + encodeURIComponent(fieldValue("hOrgID")));
		if (trim(xml)) {
			loadXml("XMLPartySubType", xml);
		}
		subtypeRoot = xmlRoot("XMLPartySubType");
		select.length = 0;
		var defaultOption = document.createElement("option");
		defaultOption.value = "0|0";
		defaultOption.text = "Select";
		select.appendChild(defaultOption);
		childElements(subtypeRoot).forEach(function (node) {
			var option = document.createElement("option");
			option.value = attr(node, "SubType");
			option.text = node.textContent || "";
			select.appendChild(option);
		});
	}

	window.AddFun = addFun;
	window.XmlFun = xmlFun;
	window.XMLFun = xmlFun;
	window.DispList = displayList;
	window.SendValue = sendValue;
	window.showpage = showPage;
	window.ShowPage = showPage;
	window.CallSearchMain = callSearchMain;
	window.CallSearch = callSearch;
	window.Init = init;
	window.ClearTable = clearTable;
	window.RemoveNode = removeNode;
	window.DisplaySubType = displaySubType;
	window.window_onunload = function () {
		if (!buttonPressed) {
			setAttr(partyRoot(), "Action", attr(partyRoot(), "Action") || "CLOSE");
		}
		returnValue(partyRoot());
	};
	window.addEventListener("beforeunload", function () {
		window.window_onunload();
	});
}(window, document));
